#!/usr/bin/env python3
"""
DB Viewer NG 3 - obecný prehliadač SQLite databáz.

Funguje s ľubovoľnou .db (SPEED.db, ODO.db, obed.db, vykaz.db, financie.db ...).

Hlavné vlastnosti
- DB sa otvára LEN NA ČÍTANIE (mode=ro, query_only) - nič sa nedá pokaziť.
- Po otvorení sa načíta len schéma; počty riadkov sa dopočítajú na pozadí.
- Tabuľky aj pohľady (views), stĺpce s typom, PK, indexy, CREATE SQL.
- Filter bez SQL (viac podmienok, AND/OR, obsahuje, IN, BETWEEN, IS NULL ...).
- SQL editor (Ctrl+Enter, zvýraznenie, história, plán dotazu).
- Všetky dotazy bežia vo vlákne, dajú sa zastaviť (Esc).
- Mriežka po stranách, triedenie klikom na hlavičku, detail riadku, Ctrl+C.
- Graf (Plotly): voliteľná os X, prevod času (ms presnosť), agregácia,
  rozdelenie podľa kategórie, grafy pod sebou, min/max decimácia (zachová špičky).
- Štatistika všetkých stĺpcov + časová analýza (Δt, medzery, duplicity).
- Export CSV / TSV / XLSX, kopírovanie do schránky.

Spustenie:  python db_viewer_ng.py [cesta.db]
Závislosti: pandas, numpy, plotly (openpyxl len pre XLSX export)

Author: MEL
"""

from __future__ import annotations

import json
import math
import queue
import os
import re
import sqlite3
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import webbrowser
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from typing import Any, Callable, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

APP_NAME = "DB Viewer NG"
APP_VERSION = "3.0"
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".db_viewer_ng.json")

DEFAULT_ROW_LIMIT = 200_000     # max. riadkov načítaných do pamäte
PAGE_SIZE = 1000                # riadkov na stranu v mriežke
FETCH_CHUNK = 20_000            # po koľkých riadkoch sa číta z DB
DEFAULT_MAX_POINTS = 50_000     # max. bodov na krivku v grafe
MAX_SPLIT_GROUPS = 30           # max. skupín pri "Rozdeliť podľa"
HISTORY_SIZE = 50
NULL_TEXT = "NULL"

# Kategorická paleta (pevné poradie, farba patrí sérii, nie poradiu)
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
           "#e87ba4", "#008300", "#4a3aa7", "#e34948"]

TIME_NAME_PRIORITY = ["time", "timestamp", "datetime", "date_time", "date", "datum",
                      "ts", "utc", "time_utc", "month", "mesiac", "day", "den", "cas", "čas"]
TIME_NAME_RE = re.compile(r"(^|_)(time|timestamp|datetime|date|datum|ts|utc|month|cas|čas)($|_)", re.I)
TIME_TYPE_RE = re.compile(r"date|time", re.I)

SQL_KEYWORDS = (
    "select from where and or not in is null like glob between order by group having limit offset "
    "as on join left right inner outer cross natural using union all distinct case when then else end "
    "asc desc with recursive values exists cast collate escape pragma explain query plan attach "
    "database detach insert update delete create drop table view index into set primary key "
    "count sum avg min max total group_concat coalesce ifnull nullif round abs length lower upper "
    "substr replace trim date time datetime julianday strftime unixepoch printf instr over "
    "partition row_number rank dense_rank lag lead first_value last_value iif"
).split()


# =============================================================================
#  Pomocné funkcie
# =============================================================================

def qi(name: str) -> str:
    """Bezpečné ohraničenie identifikátora (tabuľka/stĺpec) - "NAV-TIME" atď."""
    return '"' + str(name).replace('"', '""') + '"'


def sql_literal(v: Any) -> str:
    """Hodnota -> SQL literál (len na zobrazenie / kopírovanie do editora)."""
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):
        return repr(v)
    return "'" + str(v).replace("'", "''") + "'"


def parse_value(text: str) -> Any:
    """Text z políčka -> int / float / str."""
    t = text.strip()
    if re.fullmatch(r"[+-]?\d+", t):
        try:
            return int(t)
        except ValueError:
            pass
    if re.fullmatch(r"[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?", t):
        try:
            return float(t)
        except ValueError:
            pass
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "'\"":
        return t[1:-1]
    return t


def split_list(text: str) -> List[str]:
    return [p.strip() for p in re.split(r"[;,]", text) if p.strip()]


def fmt_int(n: Optional[int]) -> str:
    return "?" if n is None else f"{n:,}".replace(",", " ")


def fmt_num(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (float, np.floating)):
        return f"{v:.6g}"
    return str(v)


def fmt_bytes(n: int) -> str:
    for unit in ("B", "kB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def fmt_timedelta(td: pd.Timedelta) -> str:
    if pd.isna(td):
        return ""
    s = td.total_seconds()
    if abs(s) < 1:
        return f"{s * 1000:.3f} ms"
    if abs(s) < 120:
        return f"{s:.3f} s"
    return str(td)


def cell_text(v: Any) -> str:
    """Hodnota bunky -> text do mriežky."""
    if v is None:
        return NULL_TEXT
    if isinstance(v, float):                 # aj np.float64
        return NULL_TEXT if v != v else f"{float(v):.15g}"
    if isinstance(v, (bytes, bytearray, memoryview)):
        return f"<BLOB {len(v)} B>"
    try:
        if pd.isna(v):
            return NULL_TEXT
    except (TypeError, ValueError):
        pass
    return str(v)


def dedupe_columns(cols: List[str]) -> List[str]:
    """JOIN môže vrátiť 2x rovnaký názov stĺpca - pandas to nemá rád."""
    seen: Dict[str, int] = {}
    out = []
    for c in cols:
        c = str(c)
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out


def is_numeric(s: pd.Series) -> bool:
    return pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s)


def is_datetime(s: pd.Series) -> bool:
    return pd.api.types.is_datetime64_any_dtype(s)


def to_numeric(s: pd.Series) -> pd.Series:
    if is_numeric(s):
        return s.astype(float)
    return pd.to_numeric(s, errors="coerce")


_PANDAS_MAJOR = int(pd.__version__.split(".")[0])


def _parse_datetime_text(s: pd.Series) -> pd.Series:
    txt = s.astype("string")
    if _PANDAS_MAJOR >= 2:     # pandas 1.x (Debian 12) "ISO8601" nepozná -> všetko NaT
        try:
            return pd.to_datetime(txt, errors="coerce", format="ISO8601")
        except (TypeError, ValueError):
            pass
    return pd.to_datetime(txt, errors="coerce")


def looks_like_datetime_text(s: pd.Series) -> bool:
    """Textový stĺpec, ktorý vyzerá ako dátum/čas (2026-09-20, 2026-09, 12:00:00.123 ...)."""
    if is_numeric(s) or is_datetime(s):
        return False
    sample = s.dropna().head(200)
    if sample.empty:
        return False
    sample = sample.astype(str)
    if not sample.str.contains(r"\d[-/.:]\d", regex=True).all():
        return False
    parsed = _parse_datetime_text(sample)
    return parsed.notna().mean() >= 0.9


X_CONVERSIONS = ["Auto", "Bez prevodu", "Text → dátum/čas", "Unix [s]", "Unix [ms]", "Unix [µs]"]


def convert_x(s: pd.Series, mode: str) -> pd.Series:
    """Prevod osi X. Milisekundy ostávajú zachované (datetime64[ns])."""
    if mode == "Bez prevodu" or is_datetime(s):
        return s
    if mode == "Auto":
        return _parse_datetime_text(s) if looks_like_datetime_text(s) else s
    if mode == "Text → dátum/čas":
        return _parse_datetime_text(s)
    unit = {"Unix [s]": "s", "Unix [ms]": "ms", "Unix [µs]": "us"}[mode]
    return pd.to_datetime(to_numeric(s), unit=unit, errors="coerce")


def guess_time_column(columns: List[str], types: Optional[Dict[str, str]] = None) -> Optional[str]:
    """Odhad časového stĺpca podľa názvu (TIME, date, month ...) alebo typu."""
    lower = {c.lower(): c for c in columns}
    for name in TIME_NAME_PRIORITY:
        if name in lower:
            return lower[name]
    for c in columns:
        if TIME_NAME_RE.search(c):
            return c
    if types:
        for c in columns:
            if TIME_TYPE_RE.search(types.get(c, "") or ""):
                return c
    return None


def minmax_decimate(x: np.ndarray, y: np.ndarray, max_points: int) -> Tuple[np.ndarray, np.ndarray]:
    """Zníži počet bodov tak, že z každého okna vezme min aj max -> špičky ostanú.
    Očakáva dáta zoradené podľa x."""
    n = len(y)
    if n <= max_points or max_points < 4:
        return x, y
    ok = np.isfinite(y)
    idx_all = np.nonzero(ok)[0]
    if len(idx_all) <= max_points:
        return x[idx_all], y[idx_all]
    yv = y[idx_all]
    m = len(yv)
    nb = max(1, (max_points - 2) // 2)
    bucket = (np.arange(m, dtype=np.int64) * nb) // m
    order = np.lexsort((yv, bucket))          # v rámci okna podľa hodnoty
    b_sorted = bucket[order]
    first = np.r_[0, np.nonzero(np.diff(b_sorted))[0] + 1]   # min okna
    last = np.r_[first[1:] - 1, m - 1]                        # max okna
    pick = np.unique(np.concatenate([order[first], order[last], [0, m - 1]]))
    sel = idx_all[pick]
    return x[sel], y[sel]


# =============================================================================
#  Databáza (len na čítanie)
# =============================================================================

def _path_to_uri(path: str) -> str:
    p = os.path.abspath(path).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p                          # C:/x.db -> /C:/x.db
    return "file://" + urllib.parse.quote(p, safe="/:")   # UNC //srv/x -> file:////srv/x


def open_readonly(path: str) -> Tuple[sqlite3.Connection, str]:
    """Otvorí DB len na čítanie. Vráti (spojenie, popis režimu)."""
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    uri = _path_to_uri(path)
    attempts = [(uri + "?mode=ro", True, "len na čítanie"),
                (uri + "?mode=ro&immutable=1", True, "len na čítanie (immutable)"),
                (path, False, "query_only")]
    last_err: Optional[Exception] = None
    for target, use_uri, label in attempts:
        try:
            conn = sqlite3.connect(target, uri=use_uri, check_same_thread=False, timeout=10)
            conn.execute("SELECT count(*) FROM sqlite_master").fetchone()
            for pragma in ("PRAGMA query_only = 1", "PRAGMA cache_size = -131072",
                           "PRAGMA temp_store = MEMORY", "PRAGMA mmap_size = 268435456"):
                try:
                    conn.execute(pragma)
                except sqlite3.Error:
                    pass
            return conn, label
        except sqlite3.Error as e:
            last_err = e
            try:
                conn.close()
            except Exception:
                pass
    raise last_err or sqlite3.OperationalError("Nepodarilo sa otvoriť databázu.")


@dataclass
class ColumnInfo:
    name: str
    type: str
    notnull: bool
    default: Any
    pk: int


@dataclass
class ObjectInfo:
    name: str
    kind: str                       # 'table' | 'view'
    sql: str
    columns: List[ColumnInfo] = field(default_factory=list)
    indexes: List[Tuple[str, List[str], bool]] = field(default_factory=list)
    rows: Optional[int] = None
    has_rowid: bool = True

    @property
    def column_names(self) -> List[str]:
        return [c.name for c in self.columns]

    @property
    def column_types(self) -> Dict[str, str]:
        return {c.name: c.type for c in self.columns}

    def indexed_columns(self) -> set:
        s = {c.name for c in self.columns if c.pk}
        for _, cols, _ in self.indexes:
            if cols:
                s.add(cols[0])
        return s


def read_schema(conn: sqlite3.Connection) -> Dict[str, ObjectInfo]:
    rows = conn.execute(
        "SELECT name, type, COALESCE(sql,'') FROM sqlite_master "
        "WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY type, name COLLATE NOCASE"
    ).fetchall()
    schema: Dict[str, ObjectInfo] = {}
    for name, kind, sql in rows:
        obj = ObjectInfo(name=name, kind=kind, sql=sql)
        try:
            for cid, cname, ctype, notnull, dflt, pk in conn.execute(f"PRAGMA table_info({qi(name)})"):
                obj.columns.append(ColumnInfo(cname, ctype or "", bool(notnull), dflt, pk))
        except sqlite3.Error:
            pass
        if kind == "table":
            obj.has_rowid = "WITHOUT ROWID" not in sql.upper()
            try:
                for idx in conn.execute(f"PRAGMA index_list({qi(name)})").fetchall():
                    iname, unique = idx[1], bool(idx[2])
                    icols = [r[2] for r in conn.execute(f"PRAGMA index_info({qi(iname)})")]
                    obj.indexes.append((iname, icols, unique))
            except sqlite3.Error:
                pass
        else:
            obj.has_rowid = False
        schema[name] = obj
    return schema


# =============================================================================
#  Nastavenia (nedávne súbory, história SQL)
# =============================================================================

class Settings:
    def __init__(self):
        self.data: Dict[str, Any] = {"recent": [], "sql_history": [], "row_limit": DEFAULT_ROW_LIMIT,
                                     "geometry": "1400x860"}
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                self.data.update(json.load(f))
        except Exception:
            pass

    def save(self):
        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def add_recent(self, path: str):
        path = os.path.abspath(path)
        rec = [p for p in self.data.get("recent", []) if os.path.normcase(p) != os.path.normcase(path)]
        self.data["recent"] = [path] + rec[:9]
        self.save()

    def add_sql(self, sql: str):
        sql = sql.strip()
        if not sql:
            return
        hist = [s for s in self.data.get("sql_history", []) if s != sql]
        self.data["sql_history"] = [sql] + hist[:HISTORY_SIZE - 1]
        self.save()


# =============================================================================
#  Spúšťanie dotazov na pozadí
# =============================================================================

@dataclass
class QueryResult:
    df: pd.DataFrame
    sql: str
    elapsed: float
    truncated: bool = False
    total: Optional[int] = None
    sample_step: int = 1
    message: str = ""
    table: Optional[str] = None


class UiDispatcher:
    """Vlákna nesmú volať Tk priamo. Pošlú funkciu do fronty a hlavné vlákno
    ju vykoná (kontrola každých 30 ms)."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.q: "queue.Queue[Callable[[], Any]]" = queue.Queue()
        self._poll()

    def post(self, fn: Callable[[], Any]):
        self.q.put(fn)

    def _poll(self):
        try:
            for _ in range(500):
                fn = self.q.get_nowait()
                try:
                    fn()
                except Exception:
                    self.root.report_callback_exception(*sys.exc_info())
        except queue.Empty:
            pass
        try:
            self.root.after(30, self._poll)
        except tk.TclError:
            pass    # okno zatvorené


class QueryRunner:
    """Jeden dotaz naraz. Každý beží vo vlastnom vlákne a vlastnom spojení,
    dá sa prerušiť cez sqlite3.Connection.interrupt()."""

    def __init__(self, ui: UiDispatcher):
        self.ui = ui
        self._lock = threading.Lock()
        self._active: Optional[sqlite3.Connection] = None
        self._gen = 0

    @property
    def busy(self) -> bool:
        return self._active is not None

    def cancel(self):
        with self._lock:
            self._gen += 1
            if self._active is not None:
                try:
                    self._active.interrupt()
                except Exception:
                    pass

    def submit(self, db_path: str, job: Callable[[sqlite3.Connection, Callable[[str], None]], Any],
               on_done: Callable[[Any], None], on_error: Callable[[str], None],
               on_progress: Optional[Callable[[str], None]] = None):
        self.cancel()
        with self._lock:
            self._gen += 1
            gen = self._gen

        def alive() -> bool:
            return gen == self._gen

        def progress(msg: str):
            if on_progress and alive():
                self.ui.post(lambda: alive() and on_progress(msg))

        def worker():
            conn = None
            try:
                conn, _ = open_readonly(db_path)
                with self._lock:
                    if not alive():
                        return
                    self._active = conn
                result = job(conn, progress)
                self.ui.post(lambda: alive() and on_done(result))
            except sqlite3.OperationalError as e:
                msg = "Dotaz bol zastavený." if "interrupt" in str(e).lower() else f"SQL chyba: {e}"
                self.ui.post(lambda: alive() and on_error(msg))
            except Exception as e:
                detail = f"{type(e).__name__}: {e}"
                self.ui.post(lambda: alive() and on_error(detail))
            finally:
                with self._lock:
                    if self._active is conn:
                        self._active = None
                if conn is not None:
                    try:
                        conn.close()
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()


def fetch_dataframe(conn: sqlite3.Connection, sql: str, params: List[Any], limit: int,
                    progress: Callable[[str], None]) -> Tuple[pd.DataFrame, bool, str]:
    """Spustí dotaz, číta po kúskoch (FETCH_CHUNK) až do limitu.
    Vráti (df, skrátené?, správa pre príkazy bez výsledku)."""
    cur = conn.execute(sql, params)
    if cur.description is None:
        return pd.DataFrame(), False, f"Príkaz vykonaný (ovplyvnené riadky: {cur.rowcount})."
    cols = dedupe_columns([d[0] for d in cur.description])
    frames: List[pd.DataFrame] = []
    n = 0
    truncated = False
    last_report = 0.0
    while True:
        chunk = cur.fetchmany(min(FETCH_CHUNK, max(1, limit - n)))
        if not chunk:
            break
        frames.append(pd.DataFrame.from_records(chunk, columns=cols))
        n += len(chunk)
        now = time.perf_counter()
        if now - last_report > 0.25:
            progress(f"Načítavam... {fmt_int(n)} riadkov")
            last_report = now
        if n >= limit:
            truncated = cur.fetchone() is not None
            break
    if not frames:
        return pd.DataFrame(columns=cols), False, ""
    df = frames[0] if len(frames) == 1 else pd.concat(frames, ignore_index=True, copy=False)
    df = df.infer_objects()
    return df, truncated, ""


# =============================================================================
#  Filter bez SQL
# =============================================================================

OPERATORS = ["=", "!=", ">", ">=", "<", "<=", "obsahuje", "neobsahuje", "začína na",
             "končí na", "LIKE", "IN", "NOT IN", "BETWEEN", "IS NULL", "IS NOT NULL"]
NO_VALUE_OPS = {"IS NULL", "IS NOT NULL"}


@dataclass
class Condition:
    column: str
    op: str
    value: str


class SqlBuilder:
    """Zbiera SQL text + parametre. inline=True dá literály (na zobrazenie)."""

    def __init__(self, inline: bool):
        self.inline = inline
        self.params: List[Any] = []

    def lit(self, v: Any) -> str:
        if self.inline:
            return sql_literal(v)
        self.params.append(v)
        return "?"


def build_where(conds: List[Condition], combine: str, b: SqlBuilder) -> str:
    parts = []
    for c in conds:
        if not c.column or not c.op:
            continue
        col = qi(c.column)
        op, raw = c.op, c.value
        if op in NO_VALUE_OPS:
            parts.append(f"{col} {op}")
            continue
        if raw.strip() == "" and op not in ("=", "!="):
            continue
        if op in ("=", "!=", ">", ">=", "<", "<="):
            parts.append(f"{col} {op} {b.lit(parse_value(raw))}")
        elif op in ("obsahuje", "neobsahuje"):
            neg = "NOT " if op == "neobsahuje" else ""
            parts.append(f"{col} {neg}LIKE {b.lit('%' + raw.strip() + '%')}")
        elif op == "začína na":
            parts.append(f"{col} LIKE {b.lit(raw.strip() + '%')}")
        elif op == "končí na":
            parts.append(f"{col} LIKE {b.lit('%' + raw.strip())}")
        elif op == "LIKE":
            parts.append(f"{col} LIKE {b.lit(raw.strip())}")
        elif op in ("IN", "NOT IN"):
            items = split_list(raw)
            if items:
                parts.append(f"{col} {op} ({', '.join(b.lit(parse_value(i)) for i in items)})")
        elif op == "BETWEEN":
            items = split_list(raw)
            if len(items) == 2:
                parts.append(f"{col} BETWEEN {b.lit(parse_value(items[0]))} AND {b.lit(parse_value(items[1]))}")
            else:
                raise ValueError(f"BETWEEN pre '{c.column}': zadaj dve hodnoty oddelené čiarkou, napr. 10, 20")
    if not parts:
        return ""
    joiner = f" {combine} "
    return " WHERE " + joiner.join(f"({p})" if len(parts) > 1 and combine == "OR" else p for p in parts)


def build_select(table: str, columns: List[str], conds: List[Condition], combine: str,
                 order_col: str, order_desc: bool, limit: int, inline: bool,
                 sample_step: int = 1) -> Tuple[str, List[Any]]:
    b = SqlBuilder(inline)
    col_sql = ", ".join(qi(c) for c in columns) if columns else "*"
    where = build_where(conds, combine, b)
    order = f" ORDER BY {qi(order_col)}{' DESC' if order_desc else ''}" if order_col else ""
    if sample_step > 1:
        over = f"ORDER BY {qi(order_col)}{' DESC' if order_desc else ''}" if order_col else ""
        # pri SELECT * sa vráti aj __rn - zahodí sa v pandas
        sql = (f"SELECT {col_sql} FROM (\n  SELECT {col_sql}, ROW_NUMBER() OVER ({over}) AS __rn\n"
               f"  FROM {qi(table)}{where}\n) WHERE (__rn - 1) % {sample_step} = 0 ORDER BY __rn")
    else:
        sql = f"SELECT {col_sql} FROM {qi(table)}{where}{order}"
        if limit > 0:
            sql += f" LIMIT {limit + (0 if inline else 1)}"
    return sql, b.params


# =============================================================================
#  UI komponenty
# =============================================================================

class ScrolledTree(ttk.Frame):
    """Treeview so scrollbarmi."""

    def __init__(self, parent, **kw):
        super().__init__(parent)
        self.tree = ttk.Treeview(self, **kw)
        vs = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)


class SchemaPanel(ttk.Frame):
    """Ľavý panel: strom tabuliek/pohľadov + zoznam stĺpcov so zaškrtávaním."""

    def __init__(self, parent, app: "ViewerApp"):
        super().__init__(parent)
        self.app = app
        self.schema: Dict[str, ObjectInfo] = {}
        self.current: Optional[str] = None
        self.checked: Dict[str, bool] = {}
        self._build()

    def _build(self):
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # --- tabuľky
        top = ttk.Frame(paned)
        hdr = ttk.Frame(top)
        hdr.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(hdr, text="Tabuľky a pohľady", style="Header.TLabel").pack(side=tk.LEFT)
        self.obj_filter = tk.StringVar()
        srch = ttk.Frame(top)
        srch.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(srch, text="Hľadať:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(srch, textvariable=self.obj_filter).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.obj_filter.trace_add("write", lambda *_: self._fill_objects())

        st = ScrolledTree(top, columns=("rows",), selectmode="browse")
        st.pack(fill=tk.BOTH, expand=True)
        self.obj_tree = st.tree
        self.obj_tree.heading("#0", text="Názov")
        self.obj_tree.heading("rows", text="Riadkov")
        self.obj_tree.column("#0", width=200, stretch=True)
        self.obj_tree.column("rows", width=90, anchor=tk.E, stretch=False)
        self.obj_tree.tag_configure("view", foreground="#4a3aa7")
        self.obj_tree.tag_configure("col", foreground="#555555")
        self.obj_tree.bind("<<TreeviewSelect>>", self._on_obj_select)
        self.obj_tree.bind("<Double-1>", self._on_obj_double)
        self.obj_tree.bind("<Button-3>", self._on_obj_menu)
        paned.add(top, weight=2)

        # --- stĺpce
        bottom = ttk.Frame(paned)
        hdr2 = ttk.Frame(bottom)
        hdr2.pack(fill=tk.X, pady=(4, 2))
        self.cols_label = ttk.Label(hdr2, text="Stĺpce", style="Header.TLabel")
        self.cols_label.pack(side=tk.LEFT)
        self.col_filter = tk.StringVar()
        srch2 = ttk.Frame(bottom)
        srch2.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(srch2, text="Hľadať:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(srch2, textvariable=self.col_filter).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.col_filter.trace_add("write", lambda *_: self._fill_columns())

        btns = ttk.Frame(bottom)
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Všetky", command=lambda: self._set_visible(True)).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(btns, text="Žiadne", command=lambda: self._set_visible(False)).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        st2 = ScrolledTree(bottom, columns=("type",), selectmode="browse", show="tree")
        st2.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        self.col_tree = st2.tree
        self.col_tree.column("#0", width=200, stretch=True)
        self.col_tree.column("type", width=90, stretch=False)
        self.col_tree.tag_configure("on", foreground="#1a1a1a")
        self.col_tree.tag_configure("off", foreground="#777777")
        self.col_tree.bind("<Button-1>", self._on_col_click)
        self.col_tree.bind("<space>", self._on_col_space)
        self.col_tree.bind("<Double-1>", self._on_col_double)
        ttk.Label(bottom, text="Nič nevybraté = všetky stĺpce. Dvojklik vloží názov do SQL.",
                  style="Muted.TLabel", wraplength=280).pack(anchor=tk.W, pady=(2, 0))
        paned.add(bottom, weight=3)

    # ---------- plnenie
    def set_schema(self, schema: Dict[str, ObjectInfo]):
        self.schema = schema
        self.current = None
        self.checked = {}
        self._fill_objects()
        self._fill_columns()
        first = next(iter(sorted(schema.values(), key=lambda o: (o.kind != "table", o.name.lower()))), None)
        if first:
            self.obj_tree.selection_set(first.name)
            self.obj_tree.see(first.name)
            self._on_obj_select()     # hneď, nie až po udalosti

    def _fill_objects(self):
        needle = self.obj_filter.get().strip().lower()
        sel = self.current
        self.obj_tree.delete(*self.obj_tree.get_children())
        for obj in sorted(self.schema.values(), key=lambda o: (o.kind != "table", o.name.lower())):
            if needle and needle not in obj.name.lower():
                continue
            icon = "▦ " if obj.kind == "table" else "◫ "
            self.obj_tree.insert("", tk.END, iid=obj.name, text=icon + obj.name,
                                 values=(fmt_int(obj.rows) if obj.rows is not None else "…",),
                                 tags=(obj.kind,))
            pk_cols = {c.name for c in obj.columns if c.pk}
            idx_cols = obj.indexed_columns()
            for c in obj.columns:
                mark = "  [PK]" if c.name in pk_cols else ("  [idx]" if c.name in idx_cols else "")
                self.obj_tree.insert(obj.name, tk.END, iid=f"{obj.name}\x1f{c.name}",
                                     text=f"{c.name}{mark}", values=(c.type,), tags=("col",))
        if sel and self.obj_tree.exists(sel):
            self.obj_tree.selection_set(sel)

    def update_row_count(self, name: str, rows: Optional[int]):
        if name in self.schema:
            self.schema[name].rows = rows
        if self.obj_tree.exists(name):
            self.obj_tree.set(name, "rows", fmt_int(rows) if rows is not None else "chyba")

    def _fill_columns(self):
        self.col_tree.delete(*self.col_tree.get_children())
        obj = self.schema.get(self.current or "")
        if not obj:
            self.cols_label.config(text="Stĺpce")
            return
        needle = self.col_filter.get().strip().lower()
        for c in obj.columns:
            if needle and needle not in c.name.lower():
                continue
            on = self.checked.get(c.name, False)
            self.col_tree.insert("", tk.END, iid=c.name, text=("☑ " if on else "☐ ") + c.name,
                                 values=(c.type,), tags=("on" if on else "off",))
        self._update_cols_label()

    def _update_cols_label(self):
        obj = self.schema.get(self.current or "")
        if not obj:
            return
        n = len(self.selected_columns())
        self.cols_label.config(text=f"Stĺpce ({n}/{len(obj.columns)} vybraté)" if n
                               else f"Stĺpce (všetky {len(obj.columns)})")

    def _toggle(self, col: str):
        on = not self.checked.get(col, False)
        self.checked[col] = on
        if self.col_tree.exists(col):
            self.col_tree.item(col, text=("☑ " if on else "☐ ") + col, tags=("on" if on else "off",))
        self._update_cols_label()
        self.app.on_columns_changed()

    def _set_visible(self, value: bool):
        for iid in self.col_tree.get_children():
            self.checked[iid] = value
        self._fill_columns()
        self.app.on_columns_changed()

    # ---------- udalosti
    def _on_col_click(self, event):
        """Klik = zaškrtnúť/odškrtnúť. V SQL záložke klik len označí (dvojklik vkladá názov)."""
        if self.col_tree.identify_region(event.x, event.y) not in ("tree", "cell"):
            return
        iid = self.col_tree.identify_row(event.y)
        if not iid or self.app.active_tab() == "sql":
            return
        self.col_tree.selection_set(iid)
        self.col_tree.focus(iid)
        self._toggle(iid)
        return "break"

    def _on_col_space(self, _event):
        for iid in self.col_tree.selection():
            self._toggle(iid)
        return "break"

    def _on_col_double(self, event):
        iid = self.col_tree.identify_row(event.y)
        if not iid:
            return
        if self.app.active_tab() == "sql":
            self.app.sql_panel.insert_text(qi(iid) if not re.fullmatch(r"[A-Za-z_]\w*", iid) else iid)
        else:
            self._toggle(iid)   # 2. klik dvojkliku (Tk pošle <Double-1> namiesto <Button-1>)
        return "break"

    def _on_obj_select(self, _event=None):
        sel = self.obj_tree.selection()
        if not sel:
            return
        name = sel[0].split("\x1f")[0]
        if name != self.current:
            self.current = name
            self.col_filter.set("")
            self._fill_columns()
            self.app.on_table_changed(name)

    def _on_obj_double(self, event):
        iid = self.obj_tree.identify_row(event.y)
        if not iid:
            return
        if "\x1f" in iid:
            tbl, col = iid.split("\x1f", 1)
            if self.app.active_tab() == "sql":
                self.app.sql_panel.insert_text(col if re.fullmatch(r"[A-Za-z_]\w*", col) else qi(col))
            return "break"
        self.app.browse_table(iid)
        return "break"

    def _on_obj_menu(self, event):
        iid = self.obj_tree.identify_row(event.y)
        if not iid:
            return
        name = iid.split("\x1f")[0]
        self.obj_tree.selection_set(iid)
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="Zobraziť dáta", command=lambda: self.app.browse_table(name))
        m.add_command(label="Štruktúra a indexy…", command=lambda: self.app.show_structure(name))
        m.add_command(label="SELECT do SQL editora",
                      command=lambda: self.app.sql_panel.set_text(f"SELECT *\nFROM {qi(name)}\nLIMIT 1000;"))
        m.add_command(label="Kopírovať názov", command=lambda: self.app.copy_to_clipboard(name))
        m.tk_popup(event.x_root, event.y_root)

    # ---------- API
    def selected_columns(self) -> List[str]:
        obj = self.schema.get(self.current or "")
        if not obj:
            return []
        return [c for c in obj.column_names if self.checked.get(c, False)]


class ConditionRow(ttk.Frame):
    def __init__(self, parent, panel: "FilterPanel"):
        super().__init__(parent)
        self.panel = panel
        self.col = tk.StringVar()
        self.op = tk.StringVar(value="=")
        self.val = tk.StringVar()
        self.col_cb = ttk.Combobox(self, textvariable=self.col, state="readonly", width=22)
        self.col_cb.pack(side=tk.LEFT, padx=(0, 4))
        self.op_cb = ttk.Combobox(self, textvariable=self.op, values=OPERATORS, state="readonly", width=11)
        self.op_cb.pack(side=tk.LEFT, padx=(0, 4))
        self.val_cb = ttk.Combobox(self, textvariable=self.val, width=30)
        self.val_cb.pack(side=tk.LEFT, padx=(0, 4), fill=tk.X, expand=True)
        ttk.Button(self, text="hodnoty…", width=9, command=self._load_values).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(self, text="✕", width=3, command=self._remove).pack(side=tk.LEFT)
        for v in (self.col, self.op, self.val):
            v.trace_add("write", lambda *_: panel.refresh_preview())
        self.op.trace_add("write", lambda *_: self._op_changed())
        self.val_cb.bind("<Return>", lambda e: panel.app.run_filter())

    def set_columns(self, cols: List[str]):
        self.col_cb["values"] = cols
        if self.col.get() not in cols:
            self.col.set(cols[0] if cols else "")

    def _op_changed(self):
        self.val_cb.configure(state="disabled" if self.op.get() in NO_VALUE_OPS else "normal")

    def _load_values(self):
        col = self.col.get()
        table = self.panel.app.schema_panel.current
        if not col or not table:
            return
        self.val_cb["values"] = ["(načítavam…)"]

        def job(conn, progress):
            rows = conn.execute(
                f"SELECT DISTINCT {qi(col)} FROM {qi(table)} WHERE {qi(col)} IS NOT NULL "
                f"ORDER BY {qi(col)} LIMIT 1000").fetchall()
            return [cell_text(r[0]) for r in rows]

        def done(values):
            self.val_cb["values"] = values
            self.panel.app.set_status(f"{len(values)} rôznych hodnôt v '{col}'"
                                      + (" (prvých 1000)" if len(values) >= 1000 else "")
                                      + " - rozbaľ zoznam pri hodnote")
            self.val_cb.focus_set()

        self.panel.app.submit_job(job, done, busy_text=f"Hľadám hodnoty '{col}'…")

    def _remove(self):
        self.panel.remove_row(self)

    def get(self) -> Condition:
        return Condition(self.col.get(), self.op.get(), self.val.get())


class FilterPanel(ttk.Frame):
    """Záložka Filter: podmienky, triedenie, limit, vzorkovanie, náhľad SQL."""

    def __init__(self, parent, app: "ViewerApp"):
        super().__init__(parent, padding=6)
        self.app = app
        self.rows: List[ConditionRow] = []
        self.columns: List[str] = []
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X)
        ttk.Label(top, text="Podmienky", style="Header.TLabel").pack(side=tk.LEFT)
        self.combine = tk.StringVar(value="AND")
        ttk.Radiobutton(top, text="všetky platia (AND)", variable=self.combine, value="AND",
                        command=self.refresh_preview).pack(side=tk.LEFT, padx=(12, 4))
        ttk.Radiobutton(top, text="aspoň jedna (OR)", variable=self.combine, value="OR",
                        command=self.refresh_preview).pack(side=tk.LEFT)
        ttk.Button(top, text="+ podmienka", command=self.add_row).pack(side=tk.RIGHT)

        self.rows_frame = ttk.Frame(self)
        self.rows_frame.pack(fill=tk.X, pady=(4, 4))
        self.empty_label = ttk.Label(self.rows_frame, text="Žiadne podmienky - načíta sa celá tabuľka (po limit).",
                                     style="Muted.TLabel")
        self.empty_label.pack(anchor=tk.W)

        opts = ttk.Frame(self)
        opts.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(opts, text="Zoradiť podľa:").pack(side=tk.LEFT)
        self.order_col = tk.StringVar()
        self.order_cb = ttk.Combobox(opts, textvariable=self.order_col, state="readonly", width=22)
        self.order_cb.pack(side=tk.LEFT, padx=(4, 4))
        self.order_desc = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="zostupne", variable=self.order_desc,
                        command=self.refresh_preview).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(opts, text="Limit riadkov:").pack(side=tk.LEFT)
        self.limit = tk.StringVar(value=str(self.app.settings.data.get("row_limit", DEFAULT_ROW_LIMIT)))
        ttk.Entry(opts, textvariable=self.limit, width=10).pack(side=tk.LEFT, padx=(4, 12))
        self.sample = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="vzorkovať, ak je viac riadkov",
                        variable=self.sample, command=self.refresh_preview).pack(side=tk.LEFT)
        for v in (self.order_col, self.limit):
            v.trace_add("write", lambda *_: self.refresh_preview())

        act = ttk.Frame(self)
        act.pack(fill=tk.X, pady=(2, 4))
        ttk.Button(act, text="▶ Načítať  (F5)", style="Accent.TButton",
                   command=self.app.run_filter).pack(side=tk.LEFT)
        ttk.Button(act, text="Spočítať riadky", command=self.app.count_filter).pack(side=tk.LEFT, padx=6)
        ttk.Button(act, text="→ do SQL editora", command=self._to_sql).pack(side=tk.LEFT)
        ttk.Button(act, text="Vymazať podmienky", command=self.clear_rows).pack(side=tk.RIGHT)

        self.preview = tk.Text(self, height=4, wrap=tk.WORD, relief=tk.FLAT,
                               background="#f4f6f9", font=self.app.mono_font, padx=6, pady=4)
        self.preview.pack(fill=tk.BOTH, expand=True)
        self.preview.configure(state="disabled")

    def set_columns(self, cols: List[str], time_col: Optional[str]):
        self.columns = cols
        for r in list(self.rows):
            if r.col.get() and r.col.get() not in cols:
                self.remove_row(r)      # stĺpec v novej tabuľke nie je -> podmienka preč
            else:
                r.set_columns(cols)
        self.order_cb["values"] = [""] + cols
        if self.order_col.get() not in cols:
            self.order_col.set(time_col or "")
        self.refresh_preview()

    def add_row(self):
        if not self.rows:
            self.empty_label.pack_forget()
        r = ConditionRow(self.rows_frame, self)
        r.set_columns(self.columns)
        r.pack(fill=tk.X, pady=1)
        self.rows.append(r)
        self.refresh_preview()
        r.val_cb.focus_set()

    def remove_row(self, row: ConditionRow):
        self.rows.remove(row)
        row.destroy()
        if not self.rows:
            self.empty_label.pack(anchor=tk.W)
        self.refresh_preview()

    def clear_rows(self):
        for r in list(self.rows):
            self.remove_row(r)

    def conditions(self) -> List[Condition]:
        return [r.get() for r in self.rows]

    def row_limit(self) -> int:
        try:
            v = int(str(self.limit.get()).replace(" ", ""))
            return max(1, v)
        except ValueError:
            return DEFAULT_ROW_LIMIT

    def refresh_preview(self):
        table = self.app.schema_panel.current
        text = ""
        if table:
            try:
                sql, _ = build_select(table, self.app.effective_columns(), self.conditions(), self.combine.get(),
                                      self.order_col.get(), self.order_desc.get(), self.row_limit(), inline=True)
                text = sql + ";"
                if self.sample.get():
                    text += "\n-- ak by bolo viac riadkov než limit, načíta sa každý N-tý (rovnomerne)"
            except ValueError as e:
                text = f"-- {e}"
        self.preview.configure(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    def _to_sql(self):
        table = self.app.schema_panel.current
        if not table:
            return
        try:
            sql, _ = build_select(table, self.app.effective_columns(), self.conditions(), self.combine.get(),
                                  self.order_col.get(), self.order_desc.get(), self.row_limit(), inline=True)
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return
        self.app.sql_panel.set_text(sql + ";")
        self.app.select_tab("sql")


class SqlPanel(ttk.Frame):
    """Záložka SQL: editor so zvýraznením, história, plán dotazu."""

    def __init__(self, parent, app: "ViewerApp"):
        super().__init__(parent, padding=6)
        self.app = app
        self._hl_job = None
        self._build()

    def _build(self):
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(bar, text="▶ Spustiť  (Ctrl+Enter)", style="Accent.TButton",
                   command=self.app.run_sql).pack(side=tk.LEFT)
        ttk.Button(bar, text="Plán dotazu", command=self.app.explain_sql).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="Formátovať", command=self._format).pack(side=tk.LEFT)
        ttk.Label(bar, text="História:").pack(side=tk.LEFT, padx=(16, 4))
        self.hist = tk.StringVar()
        self.hist_cb = ttk.Combobox(bar, textvariable=self.hist, state="readonly", width=60)
        self.hist_cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.hist_cb.bind("<<ComboboxSelected>>", self._on_hist)
        self.refresh_history()

        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True)
        self.text = tk.Text(frame, wrap=tk.NONE, undo=True, font=self.app.mono_font, height=8,
                            padx=6, pady=4, relief=tk.FLAT, background="#fbfbfd",
                            insertbackground="#1a1a1a")
        vs = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.text.yview)
        hs = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.text.xview)
        self.text.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        bold = tkfont.Font(font=self.app.mono_font)
        bold.configure(weight="bold")
        self.text.tag_configure("kw", foreground="#1f5fbf", font=bold)
        self.text.tag_configure("str", foreground="#1b7f3a")
        self.text.tag_configure("num", foreground="#a0522d")
        self.text.tag_configure("com", foreground="#8a8a8a")
        self.text.tag_configure("ident", foreground="#6b3fa0")
        self.text.bind("<Control-Return>", lambda e: (self.app.run_sql(), "break")[1])
        self.text.bind("<KeyRelease>", lambda e: self._schedule_highlight())
        self.text.bind("<Tab>", lambda e: (self.text.insert(tk.INSERT, "    "), "break")[1])
        ttk.Label(self, text="Tip: označený text sa spustí samostatne. DB je len na čítanie. "
                              "Iné DB: ATTACH 'C:/cesta/druha.db' AS b;",
                  style="Muted.TLabel").pack(anchor=tk.W, pady=(4, 0))

    def refresh_history(self):
        hist = self.app.settings.data.get("sql_history", [])
        self.hist_cb["values"] = [re.sub(r"\s+", " ", h)[:160] for h in hist]

    def _on_hist(self, _e):
        i = self.hist_cb.current()
        hist = self.app.settings.data.get("sql_history", [])
        if 0 <= i < len(hist):
            self.set_text(hist[i])

    def _format(self):
        sql = self.get_sql(selection=False)
        kws = ["FROM", "WHERE", "GROUP BY", "ORDER BY", "HAVING", "LIMIT", "LEFT JOIN", "INNER JOIN",
               "JOIN", "UNION ALL", "UNION"]
        out = re.sub(r"\s+", " ", sql).strip()
        for k in kws:
            pattern = r"(?<!\n)\s+(" + k.replace(" ", r"\s+") + r")\b"
            out = re.sub(pattern, r"\n\1", out, flags=re.I)
        out = re.sub(r"\s+\b(AND|OR)\b\s+", r"\n  \1 ", out, flags=re.I)
        self.set_text(out)

    def get_sql(self, selection: bool = True) -> str:
        if selection:
            try:
                sel = self.text.get(tk.SEL_FIRST, tk.SEL_LAST)
                if sel.strip():
                    return sel.strip()
            except tk.TclError:
                pass
        return self.text.get("1.0", tk.END).strip()

    def set_text(self, sql: str):
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", sql)
        self.highlight()

    def insert_text(self, s: str):
        self.text.insert(tk.INSERT, s)
        self.text.focus_set()
        self._schedule_highlight()

    def _schedule_highlight(self):
        if self._hl_job:
            self.after_cancel(self._hl_job)
        self._hl_job = self.after(150, self.highlight)

    def highlight(self):
        self._hl_job = None
        t = self.text
        content = t.get("1.0", tk.END)
        if len(content) > 200_000:
            return
        for tag in ("kw", "str", "num", "com", "ident"):
            t.tag_remove(tag, "1.0", tk.END)
        kw_re = r"\b(" + "|".join(SQL_KEYWORDS) + r")\b"
        rules = [("num", r"\b\d+(\.\d+)?\b"), ("kw", kw_re), ("ident", r'"[^"\n]*"|`[^`\n]*`|\[[^\]\n]*\]'),
                 ("str", r"'(?:[^']|'')*'"), ("com", r"--[^\n]*")]
        for tag, pattern in rules:
            for m in re.finditer(pattern, content, flags=re.I):
                t.tag_add(tag, f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        t.tag_raise("str")
        t.tag_raise("com")


class ResultGrid(ttk.Frame):
    """Mriežka výsledkov po stranách. Dáta sú v DataFrame, v Treeview len aktuálna strana."""

    def __init__(self, parent, app: "ViewerApp"):
        super().__init__(parent)
        self.app = app
        self.df: pd.DataFrame = pd.DataFrame()
        self.page = 0
        self.sort_col: Optional[str] = None
        self.sort_desc = False
        self._build()

    def _build(self):
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, pady=(0, 2))
        self.info = ttk.Label(bar, text="Žiadne dáta.", style="Muted.TLabel")
        self.info.pack(side=tk.LEFT)
        pager = ttk.Frame(bar)
        pager.pack(side=tk.RIGHT)
        ttk.Button(pager, text="⏮", width=3, command=lambda: self.goto(0)).pack(side=tk.LEFT)
        ttk.Button(pager, text="◀", width=3, command=lambda: self.goto(self.page - 1)).pack(side=tk.LEFT)
        self.page_var = tk.StringVar(value="1")
        pe = ttk.Entry(pager, textvariable=self.page_var, width=6, justify=tk.CENTER)
        pe.pack(side=tk.LEFT, padx=2)
        pe.bind("<Return>", lambda e: self._goto_entry())
        self.pages_label = ttk.Label(pager, text="/ 1")
        self.pages_label.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(pager, text="▶", width=3, command=lambda: self.goto(self.page + 1)).pack(side=tk.LEFT)
        ttk.Button(pager, text="⏭", width=3, command=lambda: self.goto(10 ** 12)).pack(side=tk.LEFT)

        st = ScrolledTree(self, show="headings", selectmode="extended")
        st.pack(fill=tk.BOTH, expand=True)
        self.tree = st.tree
        self.tree.tag_configure("odd", background="#f6f8fb")
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Button-3>", self._on_right)
        self.tree.bind("<Control-c>", lambda e: self.copy_selection())
        self.tree.bind("<Control-C>", lambda e: self.copy_selection())
        self.tree.bind("<Control-a>", lambda e: (self.tree.selection_set(self.tree.get_children()), "break")[1])

    # ---------- dáta
    def set_data(self, df: pd.DataFrame, info: str):
        self.df = df
        self.sort_col = None
        self.sort_desc = False
        self.page = 0
        self._info_base = info
        self._setup_columns()
        self._render()

    def _setup_columns(self):
        cols = [str(c) for c in self.df.columns]
        self.tree["columns"] = cols
        f = self.app.ui_font
        sample = self.df.head(60)
        for c in cols:
            try:
                vals = [cell_text(v) for v in sample[c].tolist()]
            except Exception:
                vals = []
            w = max([f.measure(c + " ▲") + 18] + [f.measure(v[:40]) + 14 for v in vals])
            w = int(min(max(w, 60), 320))
            anchor = tk.E if c in self.df.columns and is_numeric(self.df[c]) else tk.W
            self.tree.heading(c, text=c, anchor=tk.W, command=lambda cc=c: self.sort_by(cc))
            self.tree.column(c, width=w, minwidth=40, stretch=False, anchor=anchor)
        self.tree.configure(displaycolumns="#all")

    def pages(self) -> int:
        return max(1, math.ceil(len(self.df) / PAGE_SIZE))

    def goto(self, page: int):
        self.page = min(max(0, page), self.pages() - 1)
        self._render()

    def _goto_entry(self):
        try:
            self.goto(int(self.page_var.get()) - 1)
        except ValueError:
            self.page_var.set(str(self.page + 1))

    def _render(self):
        self.tree.delete(*self.tree.get_children())
        n = len(self.df)
        start = self.page * PAGE_SIZE
        chunk = self.df.iloc[start:start + PAGE_SIZE]
        if len(chunk):
            values = chunk.to_numpy(dtype=object)
            insert = self.tree.insert
            for i, row in enumerate(values):
                insert("", tk.END, iid=str(start + i), values=[cell_text(v) for v in row],
                       tags=("odd",) if i % 2 else ())
        self.page_var.set(str(self.page + 1))
        self.pages_label.config(text=f"/ {self.pages()}")
        rng = f"riadky {fmt_int(start + 1)}–{fmt_int(min(n, start + PAGE_SIZE))}" if n else "0 riadkov"
        sort = f"  |  zoradené: {self.sort_col} {'▼' if self.sort_desc else '▲'}" if self.sort_col else ""
        self.info.config(text=f"{getattr(self, '_info_base', '')}  |  {rng}{sort}")

    def sort_by(self, col: str):
        if self.df.empty or col not in self.df.columns:
            return
        self.sort_desc = (not self.sort_desc) if self.sort_col == col else False
        self.sort_col = col
        try:
            self.df = self.df.sort_values(col, ascending=not self.sort_desc, kind="mergesort",
                                          na_position="last", ignore_index=True)
        except TypeError:  # zmiešané typy (text + čísla)
            key = self.df[col].map(lambda v: (0, v) if isinstance(v, (int, float)) and v == v else (1, str(v)))
            order = sorted(range(len(key)), key=lambda i: key.iat[i], reverse=self.sort_desc)
            self.df = self.df.iloc[order].reset_index(drop=True)
        for c in self.tree["columns"]:
            mark = (" ▼" if self.sort_desc else " ▲") if c == col else ""
            self.tree.heading(c, text=c + mark)
        self.app.results_changed(self.df, keep_plot_settings=True)
        self.goto(0)

    # ---------- interakcie
    def selected_positions(self) -> List[int]:
        return sorted(int(i) for i in self.tree.selection())

    def copy_selection(self):
        pos = self.selected_positions()
        if not pos:
            return "break"
        sub = self.df.iloc[pos]
        self.app.copy_to_clipboard(sub.to_csv(sep="\t", index=False))
        self.app.set_status(f"Skopírovaných {len(pos)} riadkov (s hlavičkou) - vlož do Excelu.")
        return "break"

    def _on_double(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            RowDetail(self.app, self.df, int(iid))

    def _on_right(self, event):
        region = self.tree.identify_region(event.x, event.y)
        col_id = self.tree.identify_column(event.x)
        cols = list(self.tree["displaycolumns"])
        col = None
        try:
            col = cols[int(col_id.replace("#", "")) - 1]
        except (ValueError, IndexError):
            pass
        m = tk.Menu(self, tearoff=0)
        if col:
            m.add_command(label=f"Štatistika: {col}", command=lambda: self.app.show_column_stats(col))
            m.add_command(label=f"Histogram: {col}", command=lambda: self.app.quick_histogram(col))
            m.add_command(label="Použiť ako os X v grafe", command=lambda: self.app.plot_panel.set_x(col))
            m.add_command(label="Kopírovať názov stĺpca", command=lambda: self.app.copy_to_clipboard(col))
        if region == "cell":
            iid = self.tree.identify_row(event.y)
            if iid and iid not in self.tree.selection():
                self.tree.selection_set(iid)
            if col and iid:
                val = cell_text(self.df.iloc[int(iid)][col])
                m.add_separator()
                m.add_command(label="Kopírovať hodnotu", command=lambda: self.app.copy_to_clipboard(val))
                m.add_command(label=f"Filter: {col} = {val[:30]}",
                              command=lambda: self.app.add_filter_condition(col, "=", val))
            m.add_command(label="Kopírovať vybrané riadky (Ctrl+C)", command=self.copy_selection)
            m.add_command(label="Detail riadku…", command=lambda: iid and RowDetail(self.app, self.df, int(iid)))
        m.tk_popup(event.x_root, event.y_root)


class PlotPanel(ttk.Frame):
    """Záložka Graf: nastavenia grafu nad načítanými dátami."""

    CHART_TYPES = ["Čiara", "Body", "Čiara + body", "Schodovitá čiara", "Stĺpce", "Histogram", "Box plot"]
    AGG_MODES = ["žiadna", "podľa hodnoty X", "1 s", "1 min", "1 h", "deň", "týždeň", "mesiac", "rok"]
    AGG_FUNCS = {"súčet": "sum", "priemer": "mean", "medián": "median", "min": "min", "max": "max", "počet": "count"}
    ROW_ORDER = "(poradie riadku)"
    NO_SPLIT = "(nerozdeľovať)"

    def __init__(self, parent, app: "ViewerApp"):
        super().__init__(parent, padding=6)
        self.app = app
        self.columns: List[str] = []
        self._build()

    def _build(self):
        left = ttk.Frame(self)
        left.pack(side=tk.LEFT, fill=tk.Y)
        g = ttk.Frame(left)
        g.pack(fill=tk.X)

        def row(r, label, widget):
            ttk.Label(g, text=label).grid(row=r, column=0, sticky=tk.W, pady=2, padx=(0, 6))
            widget.grid(row=r, column=1, sticky=tk.EW, pady=2)

        self.x = tk.StringVar(value=self.ROW_ORDER)
        self.x_cb = ttk.Combobox(g, textvariable=self.x, state="readonly", width=26)
        row(0, "Os X:", self.x_cb)
        self.xconv = tk.StringVar(value="Auto")
        row(1, "Prevod X:", ttk.Combobox(g, textvariable=self.xconv, values=X_CONVERSIONS,
                                          state="readonly", width=26))
        self.kind = tk.StringVar(value="Čiara")
        row(2, "Typ grafu:", ttk.Combobox(g, textvariable=self.kind, values=self.CHART_TYPES,
                                           state="readonly", width=26))
        self.split = tk.StringVar(value=self.NO_SPLIT)
        self.split_cb = ttk.Combobox(g, textvariable=self.split, state="readonly", width=26)
        row(3, "Rozdeliť podľa:", self.split_cb)
        self.agg = tk.StringVar(value="žiadna")
        row(4, "Agregácia X:", ttk.Combobox(g, textvariable=self.agg, values=self.AGG_MODES,
                                             state="readonly", width=26))
        self.func = tk.StringVar(value="súčet")
        row(5, "Funkcia:", ttk.Combobox(g, textvariable=self.func, values=list(self.AGG_FUNCS),
                                         state="readonly", width=26))
        self.max_points = tk.StringVar(value=str(DEFAULT_MAX_POINTS))
        row(6, "Max. bodov/krivka:", ttk.Entry(g, textvariable=self.max_points, width=12))
        self.subplots = tk.BooleanVar(value=False)
        ttk.Checkbutton(g, text="Každý stĺpec vo vlastnom grafe (pod sebou, spoločná os X)",
                        variable=self.subplots).grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(4, 2))
        g.columnconfigure(1, weight=1)
        ttk.Button(left, text="Vykresliť graf  (F9)", style="Accent.TButton",
                   command=self.app.plot).pack(fill=tk.X, pady=(8, 0))
        ttk.Label(left, text="Graf sa otvorí v prehliadači. Kreslia sa všetky načítané riadky;\n"
                             "pri veľa bodoch sa z každého okna berie min aj max (špičky ostanú).",
                  style="Muted.TLabel").pack(anchor=tk.W, pady=(6, 0))

        right = ttk.Frame(self)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))
        hdr = ttk.Frame(right)
        hdr.pack(fill=tk.X)
        ttk.Label(hdr, text="Os Y (vyber 1 alebo viac):").pack(side=tk.LEFT)
        ttk.Button(hdr, text="Číselné", command=self._select_numeric).pack(side=tk.RIGHT)
        ttk.Button(hdr, text="Žiadne", command=lambda: self.y_list.selection_clear(0, tk.END)).pack(side=tk.RIGHT)
        lf = ttk.Frame(right)
        lf.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        self.y_list = tk.Listbox(lf, selectmode=tk.EXTENDED, exportselection=False, activestyle="none",
                                 font=self.app.ui_font, relief=tk.FLAT, highlightthickness=1,
                                 highlightcolor="#c8d0dc", highlightbackground="#d5dbe3",
                                 selectbackground="#2a78d6", selectforeground="white")
        vs = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=self.y_list.yview)
        self.y_list.configure(yscrollcommand=vs.set)
        self.y_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vs.pack(side=tk.RIGHT, fill=tk.Y)

    def set_columns(self, df: pd.DataFrame, time_col: Optional[str], keep: bool = False):
        cols = [str(c) for c in df.columns]
        prev_y = set(self.selected_y()) if keep else set()
        self.columns = cols
        self.x_cb["values"] = [self.ROW_ORDER] + cols
        self.split_cb["values"] = [self.NO_SPLIT] + cols
        if self.x.get() not in cols or not keep:
            self.x.set(time_col if time_col in cols else self.ROW_ORDER)
        if self.split.get() not in cols:
            self.split.set(self.NO_SPLIT)
        self.y_list.delete(0, tk.END)
        num = []
        for i, c in enumerate(cols):
            numeric = is_numeric(df[c])
            self.y_list.insert(tk.END, c + ("" if numeric else "   (text)"))
            if numeric and c != self.x.get():
                num.append(i)
        if keep and prev_y:
            for i, c in enumerate(cols):
                if c in prev_y:
                    self.y_list.selection_set(i)
        else:
            for i in num[:8]:
                self.y_list.selection_set(i)

    def _select_numeric(self):
        self.y_list.selection_clear(0, tk.END)
        df = self.app.grid.df
        for i, c in enumerate(self.columns):
            if c in df.columns and is_numeric(df[c]) and c != self.x.get():
                self.y_list.selection_set(i)

    def set_x(self, col: str):
        if col in self.columns:
            self.x.set(col)
            self.app.select_tab("plot")

    def selected_y(self) -> List[str]:
        return [self.columns[i] for i in self.y_list.curselection() if i < len(self.columns)]

    def get_max_points(self) -> int:
        try:
            return max(100, int(str(self.max_points.get()).replace(" ", "")))
        except ValueError:
            return DEFAULT_MAX_POINTS


class RowDetail(tk.Toplevel):
    """Detail jedného riadku - stĺpce pod sebou (praktické pri 90+ stĺpcoch)."""

    def __init__(self, app: "ViewerApp", df: pd.DataFrame, pos: int):
        super().__init__(app.root)
        self.app, self.df, self.pos = app, df, pos
        self.geometry("560x640")
        self.transient(app.root)
        top = ttk.Frame(self, padding=6)
        top.pack(fill=tk.X)
        ttk.Button(top, text="◀", width=3, command=lambda: self.show(self.pos - 1)).pack(side=tk.LEFT)
        ttk.Button(top, text="▶", width=3, command=lambda: self.show(self.pos + 1)).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(top, text="Hľadať:").pack(side=tk.LEFT, padx=(0, 4))
        self.flt = tk.StringVar()
        ttk.Entry(top, textvariable=self.flt, width=20).pack(side=tk.LEFT)
        self.flt.trace_add("write", lambda *_: self.show(self.pos))
        ttk.Button(top, text="Kopírovať", command=self._copy).pack(side=tk.RIGHT)
        st = ScrolledTree(self, columns=("val",), show="tree headings")
        st.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self.tree = st.tree
        self.tree.heading("#0", text="Stĺpec")
        self.tree.heading("val", text="Hodnota")
        self.tree.column("#0", width=200, stretch=False)
        self.tree.column("val", width=330)
        self.tree.tag_configure("odd", background="#f6f8fb")
        self.bind("<Left>", lambda e: self.show(self.pos - 1))
        self.bind("<Right>", lambda e: self.show(self.pos + 1))
        self.bind("<Escape>", lambda e: self.destroy())
        self.show(pos)

    def show(self, pos: int):
        if not (0 <= pos < len(self.df)):
            return
        self.pos = pos
        self.title(f"Riadok {pos + 1:,} z {len(self.df):,}".replace(",", " "))
        needle = self.flt.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        row = self.df.iloc[pos]
        i = 0
        for c in self.df.columns:
            if needle and needle not in str(c).lower():
                continue
            self.tree.insert("", tk.END, text=str(c), values=(cell_text(row[c]),), tags=("odd",) if i % 2 else ())
            i += 1

    def _copy(self):
        row = self.df.iloc[self.pos]
        self.app.copy_to_clipboard("\n".join(f"{c}\t{cell_text(row[c])}" for c in self.df.columns))


class TextWindow(tk.Toplevel):
    def __init__(self, app: "ViewerApp", title: str, text: str, size="760x520"):
        super().__init__(app.root)
        self.title(title)
        self.geometry(size)
        self.transient(app.root)
        t = tk.Text(self, wrap=tk.NONE, font=app.mono_font, padx=8, pady=6, relief=tk.FLAT)
        vs = ttk.Scrollbar(self, orient=tk.VERTICAL, command=t.yview)
        t.configure(yscrollcommand=vs.set)
        vs.pack(side=tk.RIGHT, fill=tk.Y)
        t.pack(fill=tk.BOTH, expand=True)
        t.insert("1.0", text)
        t.configure(state="disabled")
        self.bind("<Escape>", lambda e: self.destroy())


class TableWindow(tk.Toplevel):
    """Okno s tabuľkou (štatistiky) + voliteľný text pod ňou."""

    def __init__(self, app: "ViewerApp", title: str, df: pd.DataFrame, extra: str = "", size="1100x560"):
        super().__init__(app.root)
        self.app, self.df = app, df
        self.title(title)
        self.geometry(size)
        self.transient(app.root)
        bar = ttk.Frame(self, padding=6)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="Kopírovať tabuľku", command=self._copy).pack(side=tk.LEFT)
        ttk.Label(bar, text="(vloží sa do Excelu)", style="Muted.TLabel").pack(side=tk.LEFT, padx=6)
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        st = ScrolledTree(paned, columns=list(df.columns), show="headings")
        tree = st.tree
        f = app.ui_font
        for c in df.columns:
            w = max([f.measure(str(c)) + 20] + [f.measure(str(v)[:40]) + 14 for v in df[c].head(200)])
            tree.heading(c, text=c, anchor=tk.W)
            tree.column(c, width=min(max(w, 50), 300), stretch=False,
                        anchor=tk.W if c in (df.columns[0], "Typ", "Najčastejšia") else tk.E)
        tree.tag_configure("odd", background="#f6f8fb")
        for i, row in enumerate(df.itertuples(index=False)):
            tree.insert("", tk.END, values=[str(v) for v in row], tags=("odd",) if i % 2 else ())
        paned.add(st, weight=3)
        if extra:
            t = tk.Text(paned, wrap=tk.WORD, font=app.mono_font, height=10, padx=8, pady=6, relief=tk.FLAT,
                        background="#f4f6f9")
            t.insert("1.0", extra)
            t.configure(state="disabled")
            paned.add(t, weight=1)
        self.bind("<Escape>", lambda e: self.destroy())

    def _copy(self):
        self.app.copy_to_clipboard(self.df.to_csv(sep="\t", index=False))


# =============================================================================
#  Štatistiky
# =============================================================================

def describe_all(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    n = len(df)
    for c in df.columns:
        s = df[c]
        nn = int(s.notna().sum())
        rec: Dict[str, Any] = {"Stĺpec": c, "Typ": str(s.dtype), "Počet": fmt_int(nn),
                               "NULL": fmt_int(n - nn), "Unikátnych": ""}
        try:
            rec["Unikátnych"] = fmt_int(int(s.nunique(dropna=True)))
        except TypeError:
            pass
        for k in ("Min", "Max", "Priemer", "Sm. odch.", "Medián", "P1", "P99", "Najčastejšia"):
            rec[k] = ""
        clean = s.dropna()
        if len(clean):
            if is_numeric(s):
                v = clean.astype(float)
                q = v.quantile([0.01, 0.5, 0.99])
                rec.update({"Min": fmt_num(v.min()), "Max": fmt_num(v.max()), "Priemer": fmt_num(v.mean()),
                            "Sm. odch.": fmt_num(v.std()), "Medián": fmt_num(q[0.5]),
                            "P1": fmt_num(q[0.01]), "P99": fmt_num(q[0.99])})
            elif is_datetime(s):
                rec.update({"Min": str(clean.min()), "Max": str(clean.max())})
            else:
                try:
                    rec.update({"Min": cell_text(clean.min())[:40], "Max": cell_text(clean.max())[:40]})
                except TypeError:
                    pass
            try:
                vc = clean.value_counts()
                if len(vc) and int(vc.iloc[0]) > 1:
                    rec["Najčastejšia"] = f"{cell_text(vc.index[0])[:30]} ({fmt_int(int(vc.iloc[0]))}×)"
            except TypeError:
                pass
        rows.append(rec)
    return pd.DataFrame(rows)


def time_analysis(df: pd.DataFrame, col: str) -> str:
    """Analýza časového stĺpca: rozsah, perióda vzorkovania, medzery, duplicity."""
    s = convert_x(df[col], "Auto")
    lines = [f"Časová analýza stĺpca '{col}'"]
    if is_datetime(s):
        t = s.dropna()
        if len(t) < 2:
            return lines[0] + ": málo hodnôt."
        dt = t.diff().dropna()
        mono = (dt >= pd.Timedelta(0)).all()
        ts = t.sort_values() if not mono else t
        d = ts.diff().dropna()
        med = d.median()
        lines += [f"  Prvý:            {t.min()}",
                  f"  Posledný:        {t.max()}",
                  f"  Rozsah:          {t.max() - t.min()}",
                  f"  Hodnôt:          {fmt_int(len(t))}   (neprevediteľných: {fmt_int(int(df[col].notna().sum() - len(t)))})",
                  f"  Medián Δt:       {fmt_timedelta(med)}" + (f"   (≈ {1 / med.total_seconds():.4g} Hz)"
                                                                 if med.total_seconds() > 0 else ""),
                  f"  Min / max Δt:    {fmt_timedelta(d.min())} / {fmt_timedelta(d.max())}"]
        if med > pd.Timedelta(0):
            gaps = d[d > 3 * med]
            lines.append(f"  Medzery > 3×medián: {fmt_int(len(gaps))}")
            for i, (idx, g) in enumerate(gaps.sort_values(ascending=False).head(5).items()):
                pos = ts.index.get_loc(idx)
                start = ts.iloc[pos - 1] if isinstance(pos, int) and pos > 0 else "?"
                lines.append(f"     {fmt_timedelta(g):>14}  od {start}")
        lines.append(f"  Duplicitné časy (Δt = 0): {fmt_int(int((d == pd.Timedelta(0)).sum()))}")
        lines.append(f"  Rastúce poradie v dátach:  {'áno' if mono else 'NIE - ' + fmt_int(int((dt < pd.Timedelta(0)).sum())) + ' skokov späť'}")
        return "\n".join(lines)
    v = to_numeric(s).dropna()
    if len(v) < 2:
        return lines[0] + ": nie je dátum ani číslo."
    d = v.diff().dropna()
    med = d.median()
    lines += ["  (číselný čas - jednotky ako v DB)",
              f"  Prvý / posledný: {fmt_num(v.min())} / {fmt_num(v.max())}",
              f"  Medián Δ:        {fmt_num(med)}",
              f"  Min / max Δ:     {fmt_num(d.min())} / {fmt_num(d.max())}",
              f"  Medzery > 3×medián: {fmt_int(int((d > 3 * med).sum())) if med > 0 else '-'}",
              f"  Duplicity (Δ = 0):  {fmt_int(int((d == 0).sum()))}",
              f"  Skoky späť (Δ < 0): {fmt_int(int((d < 0).sum()))}"]
    return "\n".join(lines)


# =============================================================================
#  Graf
# =============================================================================

def bucket_x(x: pd.Series, mode: str) -> pd.Series:
    if mode == "podľa hodnoty X":
        return x
    if not is_datetime(x):
        raise ValueError("Časová agregácia (s, min, h, deň…) potrebuje os X ako dátum/čas.\n"
                         "Nastav 'Prevod X' alebo zvoľ 'podľa hodnoty X'.")
    if mode == "1 s":
        return x.dt.floor(pd.Timedelta(seconds=1))
    if mode == "1 min":
        return x.dt.floor(pd.Timedelta(minutes=1))
    if mode == "1 h":
        return x.dt.floor(pd.Timedelta(hours=1))
    if mode == "deň":
        return x.dt.normalize()
    period = {"týždeň": "W", "mesiac": "M", "rok": "Y"}[mode]
    return x.dt.to_period(period).dt.start_time


def build_figure(df: pd.DataFrame, x_col: str, ys: List[str], kind: str, xconv: str, split: Optional[str],
                 agg: str, func: str, subplots: bool, max_points: int, title: str) -> Tuple[go.Figure, str]:
    notes: List[str] = []
    n = len(df)
    if x_col == PlotPanel.ROW_ORDER:
        x = pd.Series(np.arange(n), name="riadok")
        x_label = "poradie riadku"
    else:
        x = convert_x(df[x_col], xconv).reset_index(drop=True)
        x_label = x_col
    x_is_dt = is_datetime(x)

    data = pd.DataFrame({"__x": x})
    for y in ys:
        data[y] = df[y].reset_index(drop=True)
    if split:
        data["__s"] = df[split].reset_index(drop=True).map(cell_text)

    fname = PlotPanel.AGG_FUNCS[func]
    aggregated = agg != "žiadna" and kind not in ("Histogram", "Box plot")
    if aggregated:
        data["__x"] = bucket_x(data["__x"], agg)
        for y in ys:
            if fname != "count":
                data[y] = to_numeric(data[y])
        keys = ["__x"] + (["__s"] if split else [])
        data = data.dropna(subset=["__x"]).groupby(keys, sort=True)[ys].agg(fname).reset_index()
        notes.append(f"agregácia {agg}, {func}")
    else:
        for y in ys:
            if not is_numeric(data[y]) and kind not in ("Histogram", "Box plot"):
                conv = to_numeric(data[y])
                if conv.notna().sum() == 0:
                    notes.append(f"'{y}' je text - vynechaný (použi agregáciu 'počet' alebo histogram)")
                data[y] = conv

    ys = [y for y in ys if y in data.columns and (kind in ("Histogram",) or data[y].notna().any())]
    if not ys:
        raise ValueError("Žiadny zo zvolených stĺpcov Y nemá číselné hodnoty.")
    if kind == "Stĺpce" and len(data) > 20_000:
        raise ValueError(f"{fmt_int(len(data))} stĺpcov je priveľa na stĺpcový graf.\n"
                         "Zapni agregáciu (napr. deň/mesiac) alebo použi čiaru.")

    # skupiny
    if split:
        counts = data["__s"].value_counts()
        groups = list(counts.index[:MAX_SPLIT_GROUPS])
        if len(counts) > MAX_SPLIT_GROUPS:
            notes.append(f"zobrazených {MAX_SPLIT_GROUPS} najčastejších z {len(counts)} skupín")
        groups = sorted(groups, key=lambda g: str(g))
    else:
        groups = [None]

    line_kinds = ("Čiara", "Body", "Čiara + body", "Schodovitá čiara")
    if kind in line_kinds and not aggregated:
        if not data["__x"].is_monotonic_increasing:
            data = data.sort_values("__x", kind="mergesort")

    series_count = len(ys) * len(groups)
    use_sub = subplots and len(ys) > 1 and kind != "Histogram"
    rows = len(ys) if use_sub else 1
    if use_sub:
        fig = make_subplots(rows=rows, cols=1, shared_xaxes=True, vertical_spacing=min(0.04, 0.3 / rows),
                            subplot_titles=ys)
    else:
        fig = go.Figure()

    color_keys = groups if split else ys
    color_of = {k: PALETTE[i % len(PALETTE)] for i, k in enumerate(color_keys)}
    if len(color_keys) > len(PALETTE):
        notes.append(f"{len(color_keys)} sérií - farby sa opakujú, rozlišuj podľa legendy/hover")

    decimated = False
    total_points = 0
    for yi, y in enumerate(ys):
        for g in groups:
            sub = data if g is None else data[data["__s"] == g]
            name = y if g is None else (str(g) if len(ys) == 1 or use_sub else f"{y} · {g}")
            color = color_of[g if split else y]
            show_legend = True if not use_sub else (split and yi == 0) or (not split)
            legendgroup = str(g) if split else y
            xv = sub["__x"].to_numpy()
            yv_s = sub[y]
            if kind == "Histogram":
                vals = yv_s.dropna()
                tr = go.Histogram(x=vals, name=name, marker=dict(color=color, line=dict(color="white", width=1)),
                                  opacity=0.75 if series_count > 1 else 1.0, legendgroup=legendgroup)
            elif kind == "Box plot":
                tr = go.Box(y=to_numeric(yv_s), name=name, marker_color=color, boxmean=True,
                            legendgroup=legendgroup, showlegend=show_legend)
            elif kind == "Stĺpce":
                yv = to_numeric(yv_s).to_numpy(dtype=float)
                tr = go.Bar(x=xv, y=yv, name=name, marker=dict(color=color, line=dict(width=0)),
                            legendgroup=legendgroup, showlegend=show_legend)
                total_points += len(yv)
            else:
                yv = to_numeric(yv_s).to_numpy(dtype=float)
                before = len(yv)
                xv2, yv2 = minmax_decimate(xv, yv, max_points)
                decimated |= len(yv2) < before
                total_points += len(yv2)
                mode = {"Čiara": "lines", "Body": "markers", "Čiara + body": "lines+markers",
                        "Schodovitá čiara": "lines"}[kind]
                cls = go.Scattergl if len(yv2) > 2000 else go.Scatter
                kw: Dict[str, Any] = dict(x=xv2, y=yv2, mode=mode, name=name, legendgroup=legendgroup,
                                          showlegend=show_legend,
                                          line=dict(color=color, width=2 if len(yv2) < 5000 else 1.2),
                                          marker=dict(color=color, size=8 if len(yv2) < 2000 else 4))
                if kind == "Schodovitá čiara" and cls is go.Scatter:
                    kw["line"]["shape"] = "hv"
                elif kind == "Schodovitá čiara":
                    kw["line"]["shape"] = "hv"
                    cls = go.Scatter   # Scattergl nevie 'hv'
                tr = cls(**kw)
            if use_sub:
                fig.add_trace(tr, row=yi + 1, col=1)
            else:
                fig.add_trace(tr)

    if decimated:
        notes.append(f"min/max decimácia na ≤ {fmt_int(max_points)} bodov/krivku")

    ink, muted, grid = "#1f2328", "#5b6470", "#e7eaee"
    fig.update_layout(
        template="plotly_white",
        title=dict(text=title + (f"<br><sup style='color:{muted}'>{' · '.join(notes)}</sup>" if notes else ""),
                   font=dict(size=16, color=ink), x=0.01),
        font=dict(family="Segoe UI, Inter, Arial, sans-serif", size=12, color=ink),
        hovermode="x unified" if kind in line_kinds + ("Stĺpce",) and series_count <= 12 else "closest",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(color=muted)),
        showlegend=series_count > 1,
        margin=dict(l=64, r=24, t=90, b=56),
        plot_bgcolor="white", paper_bgcolor="white",
        barmode="overlay" if kind == "Histogram" else "group",
        bargap=0.08,
        height=max(520, 260 * rows) if use_sub else None,
    )
    fig.update_xaxes(showgrid=True, gridcolor=grid, zeroline=False, linecolor="#c9ced6",
                     tickfont=dict(color=muted))
    fig.update_yaxes(showgrid=True, gridcolor=grid, zeroline=False, linecolor="#c9ced6",
                     tickfont=dict(color=muted))
    if kind == "Histogram":
        fig.update_xaxes(title_text=", ".join(ys) if len(ys) <= 3 else "hodnota")
        fig.update_yaxes(title_text="počet")
    elif kind != "Box plot":
        fig.update_xaxes(title_text=x_label, row=rows if use_sub else None, col=1 if use_sub else None)
        if x_is_dt:
            coarse = aggregated and agg in ("deň", "týždeň", "mesiac", "rok")
            fig.update_xaxes(hoverformat="%Y-%m-%d" if coarse else "%Y-%m-%d %H:%M:%S.%L")
        if not use_sub and len(ys) == 1:
            fig.update_yaxes(title_text=ys[0] + (f" ({func})" if aggregated else ""))
    return fig, " · ".join(notes)


# =============================================================================
#  Hlavné okno
# =============================================================================

class ViewerApp:
    def __init__(self, root: tk.Tk, db_path: Optional[str] = None):
        self.root = root
        self.settings = Settings()
        self.db_path: Optional[str] = None
        self.db_mode = ""
        self.schema: Dict[str, ObjectInfo] = {}
        self.last_result: Optional[QueryResult] = None
        self.ui = UiDispatcher(root)
        self.runner = QueryRunner(self.ui)
        self.count_runner = QueryRunner(self.ui)
        self._busy_since: Optional[float] = None
        _cleanup_old_plot_dirs()
        self._plot_dir = tempfile.mkdtemp(prefix="dbviewer_")

        self._setup_style()
        root.title(f"{APP_NAME} {APP_VERSION}")
        root.geometry(self.settings.data.get("geometry", "1400x860"))
        root.minsize(900, 560)
        self._build_menu()
        self._build()
        self._bind_keys()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._tick()

        if db_path:
            self.open_db(db_path)

    # ---------- vzhľad
    def _setup_style(self):
        style = ttk.Style(self.root)
        names = style.theme_names()
        for theme in ("vista", "clam"):
            if theme in names:
                style.theme_use(theme)
                break
        family = "Segoe UI" if sys.platform.startswith("win") else "DejaVu Sans"
        mono = "Consolas" if sys.platform.startswith("win") else "DejaVu Sans Mono"
        self.ui_font = tkfont.nametofont("TkDefaultFont")
        self.ui_font.configure(family=family, size=10)
        for name in ("TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=family, size=10)
            except tk.TclError:
                pass
        self.mono_font = tkfont.Font(family=mono, size=10)
        bold = tkfont.Font(family=family, size=10, weight="bold")
        rowh = int(self.ui_font.metrics("linespace") * 1.45)
        style.configure("Treeview", rowheight=rowh)
        style.configure("Treeview.Heading", font=bold)
        style.configure("Header.TLabel", font=bold)
        style.configure("Muted.TLabel", foreground="#6b7280")
        style.configure("Status.TLabel", foreground="#374151")
        style.configure("Badge.TLabel", foreground="#1b7f3a")
        style.configure("Accent.TButton", font=bold)
        self.root.configure(background=style.lookup("TFrame", "background") or "#f0f0f0")

    def _build_menu(self):
        mb = tk.Menu(self.root)
        fm = tk.Menu(mb, tearoff=0)
        fm.add_command(label="Otvoriť databázu…", accelerator="Ctrl+O", command=self.open_dialog)
        self.recent_menu = tk.Menu(fm, tearoff=0)
        fm.add_cascade(label="Nedávne", menu=self.recent_menu)
        fm.add_command(label="Znovu načítať schému", accelerator="Ctrl+R", command=self.reload_schema)
        fm.add_separator()
        fm.add_command(label="Export výsledku…", accelerator="Ctrl+E", command=self.export)
        fm.add_separator()
        fm.add_command(label="Koniec", command=self._on_close)
        mb.add_cascade(label="Súbor", menu=fm)

        dm = tk.Menu(mb, tearoff=0)
        dm.add_command(label="Načítať (Filter) / Spustiť (SQL)", accelerator="F5", command=self.run_active)
        dm.add_command(label="Zastaviť dotaz", accelerator="Esc", command=self.cancel)
        dm.add_separator()
        dm.add_command(label="Vykresliť graf", accelerator="F9", command=self.plot)
        dm.add_command(label="Štatistika všetkých stĺpcov", accelerator="Ctrl+Shift+S", command=self.show_stats)
        dm.add_separator()
        dm.add_command(label="Informácie o databáze…", command=self.show_db_info)
        dm.add_command(label="Kontrola integrity (quick_check)", command=self.integrity_check)
        mb.add_cascade(label="Dáta", menu=dm)

        hm = tk.Menu(mb, tearoff=0)
        hm.add_command(label="Klávesové skratky", command=self._show_help)
        mb.add_cascade(label="Pomoc", menu=hm)
        self.root.config(menu=mb)
        self._refresh_recent()

    def _refresh_recent(self):
        self.recent_menu.delete(0, tk.END)
        rec = self.settings.data.get("recent", [])
        if not rec:
            self.recent_menu.add_command(label="(prázdne)", state="disabled")
        for p in rec:
            self.recent_menu.add_command(label=p, command=lambda pp=p: self.open_db(pp))

    def _build(self):
        top = ttk.Frame(self.root, padding=(8, 6))
        top.pack(fill=tk.X)
        ttk.Button(top, text="Otvoriť…", command=self.open_dialog).pack(side=tk.LEFT)
        self.db_label = ttk.Label(top, text="Žiadna databáza - otvor .db súbor (Ctrl+O)", style="Header.TLabel")
        self.db_label.pack(side=tk.LEFT, padx=10)
        self.db_badge = ttk.Label(top, text="", style="Badge.TLabel")
        self.db_badge.pack(side=tk.LEFT)
        ttk.Button(top, text="Export…", command=self.export).pack(side=tk.RIGHT)
        ttk.Button(top, text="Štatistika", command=self.show_stats).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Graf", command=self.plot).pack(side=tk.RIGHT)

        main = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.schema_panel = SchemaPanel(main, self)
        main.add(self.schema_panel, weight=1)

        right = ttk.PanedWindow(main, orient=tk.VERTICAL)
        main.add(right, weight=4)
        self.notebook = ttk.Notebook(right)
        self.filter_panel = FilterPanel(self.notebook, self)
        self.sql_panel = SqlPanel(self.notebook, self)
        self.plot_panel = PlotPanel(self.notebook, self)
        self.notebook.add(self.filter_panel, text="  Filter  ")
        self.notebook.add(self.sql_panel, text="  SQL  ")
        self.notebook.add(self.plot_panel, text="  Graf  ")
        right.add(self.notebook, weight=2)
        self.grid = ResultGrid(right, self)
        right.add(self.grid, weight=3)

        status = ttk.Frame(self.root, padding=(8, 2, 8, 6))
        status.pack(fill=tk.X)
        self.status = tk.StringVar(value="Pripravený.")
        ttk.Label(status, textvariable=self.status, style="Status.TLabel").pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(status, text="■ Zastaviť (Esc)", command=self.cancel)
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=140)
        self.elapsed = ttk.Label(status, text="", style="Muted.TLabel")
        self.elapsed.pack(side=tk.RIGHT)

    def _bind_keys(self):
        r = self.root
        r.bind("<Control-o>", lambda e: self.open_dialog())
        r.bind("<Control-r>", lambda e: self.reload_schema())
        r.bind("<Control-e>", lambda e: self.export())
        r.bind("<F5>", lambda e: self.run_active())
        r.bind("<F9>", lambda e: self.plot())
        r.bind("<Escape>", lambda e: self.cancel())
        r.bind("<Control-S>", lambda e: self.show_stats())

    # ---------- stav
    def set_status(self, text: str):
        self.status.set(text)

    def _set_busy(self, text: Optional[str]):
        if text:
            self.status.set(text)
            self._busy_since = time.perf_counter()
            self.progress.pack(side=tk.RIGHT, padx=6)
            self.stop_btn.pack(side=tk.RIGHT)
            self.progress.start(12)
            self.root.configure(cursor="watch")
        else:
            self._busy_since = None
            self.progress.stop()
            self.progress.pack_forget()
            self.stop_btn.pack_forget()
            self.elapsed.config(text="")
            self.root.configure(cursor="")

    def _tick(self):
        if self._busy_since is not None:
            self.elapsed.config(text=f"{time.perf_counter() - self._busy_since:.1f} s")
        self.root.after(200, self._tick)

    def active_tab(self) -> str:
        idx = self.notebook.index(self.notebook.select())
        return ["filter", "sql", "plot"][idx]

    def select_tab(self, name: str):
        self.notebook.select({"filter": 0, "sql": 1, "plot": 2}[name])

    def copy_to_clipboard(self, text: str):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    # ---------- DB
    def open_dialog(self):
        initial = os.path.dirname(self.db_path) if self.db_path else None
        path = filedialog.askopenfilename(title="Otvoriť SQLite databázu", initialdir=initial,
                                          filetypes=[("SQLite", "*.db *.sqlite *.sqlite3 *.db3"),
                                                     ("Všetky súbory", "*.*")])
        if path:
            self.open_db(path)

    def open_db(self, path: str):
        self.cancel()
        self.count_runner.cancel()
        try:
            conn, mode = open_readonly(path)
            try:
                schema = read_schema(conn)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Databázu sa nepodarilo otvoriť:\n{path}\n\n{e}")
            return
        self.db_path = os.path.abspath(path)
        self.db_mode = mode
        self.schema = schema
        size = os.path.getsize(path)
        self.db_label.config(text=f"{os.path.basename(path)}   ({fmt_bytes(size)})")
        self.db_badge.config(text=f"● {mode}")
        self.root.title(f"{APP_NAME} {APP_VERSION} - {os.path.basename(path)}")
        self.settings.add_recent(self.db_path)
        self._refresh_recent()
        self.schema_panel.set_schema(schema)
        n_t = sum(1 for o in schema.values() if o.kind == "table")
        self.set_status(f"Otvorené: {self.db_path}  |  {n_t} tabuliek, {len(schema) - n_t} pohľadov")
        self._count_rows()
        if not schema:
            messagebox.showinfo(APP_NAME, "Databáza neobsahuje žiadne tabuľky.")

    def reload_schema(self):
        if self.db_path:
            cur = self.schema_panel.current
            self.open_db(self.db_path)
            if cur and cur in self.schema and self.schema_panel.obj_tree.exists(cur):
                self.schema_panel.obj_tree.selection_set(cur)

    def _count_rows(self):
        """Počty riadkov na pozadí - tabuľky od najmenšej (podľa poradia), pohľady sa nepočítajú."""
        names = [o.name for o in self.schema.values() if o.kind == "table"]
        if not names or not self.db_path:
            return
        panel = self.schema_panel

        def job(conn, progress):
            for name in names:
                try:
                    n = conn.execute(f"SELECT COUNT(*) FROM {qi(name)}").fetchone()[0]
                except sqlite3.OperationalError as e:
                    if "interrupt" in str(e).lower():
                        raise
                    n = None
                self.ui.post(lambda nm=name, nn=n: panel.update_row_count(nm, nn))
            return None

        self.count_runner.submit(self.db_path, job, on_done=lambda _: None, on_error=lambda _: None)

    def submit_job(self, job, on_done, busy_text: str = "Pracujem…"):
        if not self.db_path:
            messagebox.showwarning(APP_NAME, "Najprv otvor databázu.")
            return
        self._set_busy(busy_text)

        def done(res):
            self._set_busy(None)
            on_done(res)

        def err(msg):
            self._set_busy(None)
            self.set_status(msg if msg == "Dotaz bol zastavený." else "Chyba.")
            if msg != "Dotaz bol zastavený.":
                messagebox.showerror(APP_NAME, msg)

        self.runner.submit(self.db_path, job, done, err, on_progress=self.set_status)

    def cancel(self):
        if self.runner.busy:
            self.runner.cancel()
            self._set_busy(None)
            self.set_status("Dotaz bol zastavený.")

    # ---------- udalosti z panelov
    def on_table_changed(self, name: str):
        obj = self.schema.get(name)
        if not obj:
            return
        self.filter_panel.set_columns(obj.column_names, guess_time_column(obj.column_names, obj.column_types))

    def on_columns_changed(self):
        self.filter_panel.refresh_preview()

    def effective_columns(self) -> List[str]:
        """Vybraté stĺpce + automaticky časový stĺpec (ak existuje), aby sa dalo kresliť."""
        sel = self.schema_panel.selected_columns()
        obj = self.schema.get(self.schema_panel.current or "")
        if not sel or not obj:
            return sel
        tcol = guess_time_column(obj.column_names, obj.column_types)
        if tcol and tcol not in sel:
            sel = [tcol] + sel
        return sel

    def add_filter_condition(self, col: str, op: str, value: str):
        if col not in self.filter_panel.columns:
            messagebox.showinfo(APP_NAME, f"Stĺpec '{col}' nie je v aktuálnej tabuľke.")
            return
        self.filter_panel.add_row()
        r = self.filter_panel.rows[-1]
        r.col.set(col)
        r.op.set("IS NULL" if value == NULL_TEXT else op)
        r.val.set("" if value == NULL_TEXT else value)
        self.select_tab("filter")

    # ---------- dotazy
    def run_active(self):
        tab = self.active_tab()
        if tab == "sql":
            self.run_sql()
        elif tab == "plot":
            self.plot()
        else:
            self.run_filter()

    def browse_table(self, name: str):
        if self.schema_panel.current != name and self.schema_panel.obj_tree.exists(name):
            self.schema_panel.obj_tree.selection_set(name)
            self.schema_panel._on_obj_select()
        self.select_tab("filter")
        self.run_filter()

    def run_filter(self):
        table = self.schema_panel.current
        if not table or not self.db_path:
            messagebox.showwarning(APP_NAME, "Najprv otvor databázu a vyber tabuľku.")
            return
        fp = self.filter_panel
        cols = self.effective_columns()
        conds, combine = fp.conditions(), fp.combine.get()
        order_col, desc = fp.order_col.get(), fp.order_desc.get()
        limit = fp.row_limit()
        sample = fp.sample.get()
        self.settings.data["row_limit"] = limit
        try:
            build_where(conds, combine, SqlBuilder(True))
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return

        def job(conn, progress):
            t0 = time.perf_counter()
            total = None
            step = 1
            if sample:
                progress("Počítam riadky…")
                b = SqlBuilder(False)
                where = build_where(conds, combine, b)
                total = conn.execute(f"SELECT COUNT(*) FROM {qi(table)}{where}", b.params).fetchone()[0]
                if total > limit:
                    step = math.ceil(total / limit)
            sql, params = build_select(table, cols, conds, combine, order_col, desc, limit, inline=False,
                                       sample_step=step)
            df, truncated, msg = fetch_dataframe(conn, sql, params, limit, progress)
            if "__rn" in df.columns:
                df = df.drop(columns="__rn")
            shown, _ = build_select(table, cols, conds, combine, order_col, desc, limit, inline=True,
                                    sample_step=step)
            return QueryResult(df=df, sql=shown, elapsed=time.perf_counter() - t0, truncated=truncated,
                               total=total, sample_step=step, message=msg, table=table)

        self.submit_job(job, self._on_result, busy_text=f"Načítavam '{table}'…")

    def count_filter(self):
        table = self.schema_panel.current
        if not table:
            return
        conds, combine = self.filter_panel.conditions(), self.filter_panel.combine.get()

        def job(conn, progress):
            b = SqlBuilder(False)
            where = build_where(conds, combine, b)
            return conn.execute(f"SELECT COUNT(*) FROM {qi(table)}{where}", b.params).fetchone()[0]

        self.submit_job(job, lambda n: self.set_status(f"Podmienke vyhovuje {fmt_int(n)} riadkov."),
                        busy_text="Počítam…")

    def run_sql(self):
        sql = self.sql_panel.get_sql()
        if not sql:
            return
        if not self.db_path:
            messagebox.showwarning(APP_NAME, "Najprv otvor databázu.")
            return
        self.settings.add_sql(sql)
        self.sql_panel.refresh_history()
        limit = self.filter_panel.row_limit()
        clean = sql.strip().rstrip(";").strip()

        def job(conn, progress):
            t0 = time.perf_counter()
            # viac príkazov (napr. ATTACH ...; SELECT ...): všetky okrem posledného
            # sa vykonajú, posledný vráti dáta
            parts = _split_sql(clean)
            if not parts:
                raise ValueError("Prázdny dotaz (len komentár).")
            for p in parts[:-1]:
                conn.execute(p)
            df, truncated, msg = fetch_dataframe(conn, parts[-1], [], limit, progress)
            return QueryResult(df=df, sql=sql, elapsed=time.perf_counter() - t0, truncated=truncated,
                               message=msg)

        self.submit_job(job, self._on_result, busy_text="Spúšťam SQL…")

    def explain_sql(self):
        sql = self.sql_panel.get_sql().rstrip(";")
        if not sql:
            return

        def job(conn, progress):
            rows = conn.execute("EXPLAIN QUERY PLAN " + sql).fetchall()
            return rows

        def done(rows):
            lines = []
            parents = {}
            for r in rows:
                rid, parent, detail = r[0], r[1], r[-1]
                depth = parents.get(parent, -1) + 1
                parents[rid] = depth
                lines.append("   " * depth + "└─ " + str(detail))
            txt = "\n".join(lines)
            txt += ("\n\nSCAN = prechádza celú tabuľku (pomalé pri veľkých tabuľkách).\n"
                    "SEARCH ... USING INDEX = používa index (rýchle).\n"
                    "USE TEMP B-TREE = triedenie/zoskupenie v pamäti.")
            TextWindow(self, "Plán dotazu", txt, size="760x380")

        self.submit_job(job, done, busy_text="Plán dotazu…")

    def _on_result(self, res: QueryResult):
        self.last_result = res
        df = res.df
        if res.message and df.empty and len(df.columns) == 0:
            self.set_status(res.message)
            return
        info = f"{fmt_int(len(df))} riadkov × {len(df.columns)} stĺpcov"
        extra = []
        if res.sample_step > 1:
            extra.append(f"vzorkované: každý {res.sample_step}. z {fmt_int(res.total)}")
        elif res.total is not None:
            extra.append(f"celkom {fmt_int(res.total)}")
        if res.truncated:
            extra.append(f"SKRÁTENÉ na limit {fmt_int(len(df))} (zvýš limit alebo zapni vzorkovanie)")
        info_full = info + ("  (" + ", ".join(extra) + ")" if extra else "")
        self.grid.set_data(df, info_full)
        self.results_changed(df)
        self.set_status(f"{info_full}  |  {res.elapsed:.2f} s")

    def results_changed(self, df: pd.DataFrame, keep_plot_settings: bool = False):
        types = {}
        if self.last_result and self.last_result.table and self.last_result.table in self.schema:
            types = self.schema[self.last_result.table].column_types
        tcol = guess_time_column([str(c) for c in df.columns], types)
        self.plot_panel.set_columns(df, tcol, keep=keep_plot_settings)

    # ---------- graf, štatistiky, export
    def _need_data(self) -> Optional[pd.DataFrame]:
        df = self.grid.df
        if df is None or df.empty:
            messagebox.showwarning(APP_NAME, "Najprv načítaj dáta (Filter alebo SQL).")
            return None
        return df

    def plot(self):
        df = self._need_data()
        if df is None:
            return
        pp = self.plot_panel
        ys = pp.selected_y()
        if not ys:
            self.select_tab("plot")
            messagebox.showwarning(APP_NAME, "Vyber aspoň jeden stĺpec pre os Y (záložka Graf).")
            return
        split = pp.split.get()
        split = None if split == PlotPanel.NO_SPLIT else split
        src = self.last_result.table if self.last_result and self.last_result.table else "SQL dotaz"
        title = f"{os.path.basename(self.db_path or '')} · {src} · {fmt_int(len(df))} riadkov"
        args = (df, pp.x.get(), ys, pp.kind.get(), pp.xconv.get(), split, pp.agg.get(), pp.func.get(),
                pp.subplots.get(), pp.get_max_points(), title)
        self._set_busy("Kreslím graf…")

        def worker():
            try:
                fig, notes = build_figure(*args)
                path = os.path.join(self._plot_dir, f"graf_{int(time.time() * 1000)}.html")
                fig.write_html(path, include_plotlyjs="directory", full_html=True,
                               config={"scrollZoom": True, "displaylogo": False,
                                       "toImageButtonOptions": {"format": "png", "scale": 2}})
                self.ui.post(lambda: self._plot_done(path, notes))
            except Exception as e:
                traceback.print_exc()
                msg = str(e) if isinstance(e, ValueError) else f"{type(e).__name__}: {e}"
                self.ui.post(lambda: self._plot_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _plot_done(self, path: str, notes: str):
        self._set_busy(None)
        self.set_status("Graf otvorený v prehliadači." + (f"  ({notes})" if notes else ""))
        webbrowser.open(path_to_file_url(path))

    def _plot_error(self, msg: str):
        self._set_busy(None)
        self.set_status("Graf sa nepodarilo vytvoriť.")
        messagebox.showerror(APP_NAME, msg)

    def quick_histogram(self, col: str):
        df = self._need_data()
        if df is None:
            return
        pp = self.plot_panel
        pp.kind.set("Histogram")
        pp.split.set(PlotPanel.NO_SPLIT)
        pp.y_list.selection_clear(0, tk.END)
        if col in pp.columns:
            pp.y_list.selection_set(pp.columns.index(col))
        self.plot()

    def show_stats(self):
        df = self._need_data()
        if df is None:
            return
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            table = describe_all(df)
            tcol = guess_time_column([str(c) for c in df.columns])
            extra = time_analysis(df, tcol) if tcol else ""
        finally:
            self.root.configure(cursor="")
        TableWindow(self, f"Štatistika - {fmt_int(len(df))} riadkov", table, extra)

    def show_column_stats(self, col: str):
        df = self._need_data()
        if df is None or col not in df.columns:
            return
        table = describe_all(df[[col]]).T.reset_index()
        table.columns = ["Vlastnosť", "Hodnota"]
        extra = ""
        s = df[col]
        if not is_numeric(s):
            try:
                vc = s.value_counts(dropna=False).head(15)
                extra = "Najčastejšie hodnoty:\n" + "\n".join(
                    f"  {fmt_int(int(n)):>10}×  {cell_text(v)[:80]}" for v, n in vc.items())
            except TypeError:
                pass
        if guess_time_column([col]) == col or looks_like_datetime_text(s):
            extra = (extra + "\n\n" if extra else "") + time_analysis(df, col)
        TableWindow(self, f"Štatistika: {col}", table, extra, size="560x560")

    def export(self):
        df = self._need_data()
        if df is None:
            return
        base = (self.last_result.table if self.last_result and self.last_result.table else "vysledok")
        path = filedialog.asksaveasfilename(
            title="Export", initialfile=f"{base}.csv", defaultextension=".csv",
            filetypes=[("CSV (oddeľovač ;)", "*.csv"), ("TSV (tabulátor)", "*.tsv *.txt"),
                       ("Excel", "*.xlsx"), ("Všetky", "*.*")])
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        self._set_busy("Exportujem…")

        def worker():
            try:
                if ext == ".xlsx":
                    if len(df) > 1_048_575:
                        raise ValueError("Excel má limit 1 048 576 riadkov. Použi CSV.")
                    out = df.copy()
                    for c in out.columns:
                        if isinstance(out[c].dtype, pd.DatetimeTZDtype):
                            out[c] = out[c].dt.tz_localize(None)
                        elif out[c].dtype == object:
                            out[c] = out[c].map(lambda v: f"<BLOB {len(v)} B>"
                                                if isinstance(v, (bytes, bytearray, memoryview)) else v)
                    out.to_excel(path, index=False, engine="openpyxl")
                else:
                    df.to_csv(path, sep=";" if ext == ".csv" else "\t", index=False, encoding="utf-8-sig")
                self.ui.post(lambda: (self._set_busy(None),
                                      self.set_status(f"Exportované {fmt_int(len(df))} riadkov → {path}")))
            except Exception as e:
                msg = f"Export zlyhal: {e}"
                self.ui.post(lambda: (self._set_busy(None), messagebox.showerror(APP_NAME, msg)))

        threading.Thread(target=worker, daemon=True).start()

    # ---------- info okná
    def show_structure(self, name: str):
        obj = self.schema.get(name)
        if not obj:
            return
        lines = [f"{'TABUĽKA' if obj.kind == 'table' else 'POHĽAD'}: {obj.name}",
                 f"Riadkov: {fmt_int(obj.rows) if obj.rows is not None else '?'}", "",
                 f"{'#':>3}  {'Stĺpec':<32} {'Typ':<14} {'NOT NULL':<9} {'PK':<3} Default"]
        for i, c in enumerate(obj.columns):
            lines.append(f"{i:>3}  {c.name:<32} {c.type:<14} {'áno' if c.notnull else '':<9} "
                         f"{c.pk or '':<3} {'' if c.default is None else c.default}")
        if obj.indexes:
            lines += ["", "Indexy:"]
            for iname, icols, uniq in obj.indexes:
                lines.append(f"  {'UNIQUE ' if uniq else ''}{iname} ({', '.join(icols)})")
        elif obj.kind == "table":
            lines += ["", "Indexy: žiadne (filtrovanie/triedenie prechádza celú tabuľku)"]
        lines += ["", "CREATE SQL:", obj.sql]
        TextWindow(self, f"Štruktúra - {name}", "\n".join(lines))

    def show_db_info(self):
        if not self.db_path:
            return

        def job(conn, progress):
            info = {}
            for p in ("page_size", "page_count", "freelist_count", "journal_mode", "encoding",
                      "user_version", "application_id", "auto_vacuum"):
                try:
                    info[p] = conn.execute(f"PRAGMA {p}").fetchone()[0]
                except sqlite3.Error:
                    info[p] = "?"
            info["sqlite"] = sqlite3.sqlite_version
            return info

        def done(info):
            st = os.stat(self.db_path)
            wal = self.db_path + "-wal"
            lines = [f"Súbor:          {self.db_path}",
                     f"Veľkosť:        {fmt_bytes(st.st_size)}",
                     f"Zmenený:        {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_mtime))}",
                     f"Režim:          {self.db_mode}",
                     f"WAL súbor:      {fmt_bytes(os.path.getsize(wal)) if os.path.exists(wal) else 'nie je'}",
                     f"Journal mode:   {info['journal_mode']}",
                     f"Encoding:       {info['encoding']}",
                     f"Page size:      {info['page_size']} B × {info['page_count']} strán",
                     f"Voľné strany:   {info['freelist_count']}  (VACUUM by uvoľnil "
                     f"{fmt_bytes(int(info['freelist_count'] or 0) * int(info['page_size'] or 0))})",
                     f"user_version:   {info['user_version']}",
                     f"SQLite knižnica:{info['sqlite']}", "",
                     "Objekty:"]
            for o in sorted(self.schema.values(), key=lambda o: (o.kind, o.name.lower())):
                lines.append(f"  {o.kind:<6} {o.name:<36} {fmt_int(o.rows) if o.rows is not None else '':>14}"
                             f"  {len(o.columns)} stĺpcov, {len(o.indexes)} indexov")
            TextWindow(self, "Informácie o databáze", "\n".join(lines))

        self.submit_job(job, done, busy_text="Čítam informácie…")

    def integrity_check(self):
        def job(conn, progress):
            return [r[0] for r in conn.execute("PRAGMA quick_check").fetchall()]

        def done(rows):
            ok = rows == ["ok"]
            (messagebox.showinfo if ok else messagebox.showwarning)(
                APP_NAME, "Databáza je v poriadku (quick_check = ok)." if ok else "\n".join(rows[:30]))

        self.submit_job(job, done, busy_text="Kontrolujem integritu…")

    def _show_help(self):
        TextWindow(self, "Klávesové skratky", "\n".join([
            "Ctrl+O          otvoriť databázu",
            "Ctrl+R          znovu načítať schému",
            "F5              Filter: načítať / SQL: spustiť",
            "Ctrl+Enter      spustiť SQL (označený text = len ten)",
            "Esc             zastaviť bežiaci dotaz",
            "F9              vykresliť graf",
            "Ctrl+Shift+S    štatistika všetkých stĺpcov",
            "Ctrl+E          export",
            "",
            "Mriežka:",
            "  klik na hlavičku      zoradiť (znova = opačne)",
            "  pravý klik            štatistika / histogram / os X / filter podľa hodnoty",
            "  dvojklik na riadok    detail riadku (stĺpce pod sebou, ←/→ ďalší)",
            "  Ctrl+C                kopírovať vybrané riadky (do Excelu)",
            "",
            "Strom tabuliek:",
            "  dvojklik na tabuľku   zobraziť dáta",
            "  pravý klik            štruktúra, indexy, SELECT do editora",
            "  [PK] = primárny kľúč, [idx] = indexovaný stĺpec (rýchle filtrovanie)",
            "",
            "Stĺpce: klik na ☐ = vybrať, medzerník = prepnúť, dvojklik v SQL záložke = vložiť názov.",
        ]), size="620x520")

    def _on_close(self):
        try:
            self.settings.data["geometry"] = self.root.geometry()
            self.settings.save()
        except Exception:
            pass
        self.runner.cancel()
        self.count_runner.cancel()
        self.root.destroy()


def path_to_file_url(path: str) -> str:
    return _path_to_uri(path)


def _split_sql(sql: str) -> List[str]:
    """Rozdelí text na príkazy podľa ';' mimo reťazcov a komentárov."""
    out, buf = [], []
    i, n = 0, len(sql)
    quote = None
    while i < n:
        ch = sql[i]
        if quote:
            buf.append(ch)
            if ch == quote:
                if i + 1 < n and sql[i + 1] == quote:
                    buf.append(sql[i + 1])
                    i += 1
                else:
                    quote = None
        elif ch in "'\"`":
            quote = ch
            buf.append(ch)
        elif ch == "-" and sql[i:i + 2] == "--":
            j = sql.find("\n", i)
            j = n if j < 0 else j
            buf.append(sql[i:j])
            i = j
            continue
        elif ch == ";":
            out.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    if "".join(buf).strip():
        out.append("".join(buf))
    # vynechať časti, kde je len komentár / prázdno
    return [p for p in out if re.sub(r"--[^\n]*", "", p).strip()]


def _enable_dpi_awareness():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def _cleanup_old_plot_dirs(days: int = 7):
    """Zmaže staré priečinky s grafmi (%TEMP%\\dbviewer_*) staršie ako `days` dní."""
    import glob
    import shutil
    limit = time.time() - days * 86400
    for d in glob.glob(os.path.join(tempfile.gettempdir(), "dbviewer_*")):
        try:
            if os.path.isdir(d) and os.path.getmtime(d) < limit:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def _resource_path(name: str) -> str:
    """Cesta k súboru vedľa skriptu alebo vo vnútri .exe (PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def main():
    # .exe / .pyw nemá konzolu -> sys.stdout/stderr = None, výpis chyby by spadol
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    _enable_dpi_awareness()
    root = tk.Tk()
    ico = _resource_path("db_viewer.ico")
    if sys.platform.startswith("win") and os.path.isfile(ico):
        try:
            root.iconbitmap(default=ico)
        except tk.TclError:
            pass

    db_path = sys.argv[1] if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]) else None
    if db_path is None:
        try:
            from LIB.config import CONFIG
            base = CONFIG.get("paths", {}).get("db", "")
            candidate = os.path.join(base, "TEST_FUTURE_speed_FULL2.db") if base else None
            if candidate and os.path.exists(candidate):
                db_path = candidate
        except Exception:
            pass

    def report(exc, val, tb):
        traceback.print_exception(exc, val, tb)
        messagebox.showerror(APP_NAME, f"Neočakávaná chyba:\n{exc.__name__}: {val}")

    root.report_callback_exception = report
    ViewerApp(root, db_path)
    root.mainloop()


if __name__ == "__main__":
    main()
