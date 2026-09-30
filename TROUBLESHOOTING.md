# DB Viewer – inštalácia, riešenie problémov, tipy

## Kde čo je

| Čo | Windows | Linux |
|---|---|---|
| Program | `C:\Users\eliasmichal\Dev\Projects\Personal\DBViewer\db_viewer_ng.pyw` | `~/Tools/DBViewer/db_viewer_ng.pyw` |
| Ikona | `db_viewer.ico` (vedľa programu) | `db_viewer.png` (vedľa programu) |
| Python (env) | `C:\Users\eliasmichal\Dev\Env\Personal\Scripts\pythonw.exe` | `python3` zo systému |
| Odkaz | ikona **DB Viewer** na ploche | `~/.local/share/applications/db-viewer.desktop` + plocha |
| Nastavenia (nedávne DB, história SQL, veľkosť okna, limit) | `%USERPROFILE%\.db_viewer_ng.json` | `~/.db_viewer_ng.json` |
| Grafy (dočasné HTML) | `%TEMP%\dbviewer_*` (staršie ako 7 dní sa mažú samé) | `/tmp/dbviewer_*` |

Potrebné balíky: `pandas`, `numpy`, `plotly`, `tkinter` (súčasť Pythonu). Voliteľne `openpyxl` (len export do XLSX).
Otestované: Python 3.10 – 3.13, pandas 2.2 a 3.0.

---

## 1. Inštalácia na nový / preinštalovaný Windows

**Čo si zálohovať pred preinštalovaním:**
- priečinok `DBViewer` (`db_viewer_ng.pyw`, `db_viewer.ico`, tento návod),
- voliteľne `%USERPROFILE%\.db_viewer_ng.json` (história SQL dotazov a nedávne súbory).

Env (`Dev\Env\Personal`) **nezálohuj**. Env sa nedá presunúť, vždy sa robí nanovo (krok 3).

**Postup:**

1. **Python** z python.org (rovnaký, aký používaš pre ostatné projekty).
   Pri inštalácii nechaj zaškrtnuté **„tcl/tk and IDLE“**. Bez toho nebude `tkinter` a viewer sa nespustí.
   Verziu z Microsoft Store nepoužívaj, `pythonw.exe` je tam schovaný na zvláštnom mieste.

2. **Skopíruj** priečinok `DBViewer` späť do `C:\Users\eliasmichal\Dev\Projects\Personal\`.

3. **Env** (ak ho ešte nemáš), v PowerShell alebo cmd:
   ```
   python -m venv C:\Users\eliasmichal\Dev\Env\Personal
   C:\Users\eliasmichal\Dev\Env\Personal\Scripts\python.exe -m pip install pandas numpy plotly openpyxl
   ```

4. **Kontrola**, že env má všetko (musí vypísať `OK`):
   ```
   C:\Users\eliasmichal\Dev\Env\Personal\Scripts\python.exe -c "import tkinter, pandas, numpy, plotly, openpyxl; print('OK')"
   ```

5. **Odkaz na ploche:** pravý klik na plochu → *Nový* → *Odkaz* → do poľa vlož (jeden riadok):
   ```
   "C:\Users\eliasmichal\Dev\Env\Personal\Scripts\pythonw.exe" "C:\Users\eliasmichal\Dev\Projects\Personal\DBViewer\db_viewer_ng.pyw"
   ```
   Názov: **DB Viewer**. Potom *Vlastnosti* → *Zmeniť ikonu* → *Prehľadávať* → `db_viewer.ico`.

6. **Test:** dvojklik na ikonu. Otvorí sa viewer a **nesmie** sa otvoriť terminál.

## 2. Inštalácia na Linux

```
mkdir -p ~/Tools && cp -r DBViewer ~/Tools/
sudo apt install python3-tk python3-pandas python3-plotly python3-openpyxl   # alebo pip do venv
sh ~/Tools/DBViewer/install_linux.sh
```
Skript skontroluje balíky a vytvorí spúšťač v menu aj na ploche. Ak GNOME pri prvom spustení z plochy zobrazí upozornenie: pravý klik → *Povoliť spúšťanie*.
Ak používaš venv, spusti skript s aktivovaným venv (`source venv/bin/activate`). Skript zapíše ten Python, ktorý nájde ako `python3`.

---

## 3. Riešenie problémov

### Po dvojkliku na ikonu sa nič neotvorí
`pythonw` chyby nezobrazuje. Spusti to isté s `python.exe` v termináli a uvidíš príčinu:
```
C:\Users\eliasmichal\Dev\Env\Personal\Scripts\python.exe C:\Users\eliasmichal\Dev\Projects\Personal\DBViewer\db_viewer_ng.pyw
```

| Chyba v termináli | Príčina | Riešenie |
|---|---|---|
| `ModuleNotFoundError: No module named 'pandas'` (alebo `plotly`, `numpy`) | odkaz používa Python bez balíkov | skontroluj cestu v odkaze, alebo `...\Scripts\python.exe -m pip install pandas numpy plotly` |
| `ModuleNotFoundError: No module named 'tkinter'` | Python bez tcl/tk | spusti inštalátor Pythonu → *Modify* → zaškrtni *tcl/tk and IDLE* |
| `SyntaxError` / `IndentationError` + číslo riadku | chyba pri ručnej úprave kódu | oprav riadok, alebo vráť pôvodný súbor |
| `...\pythonw.exe` sa nenašiel (hlási Windows) | env neexistuje alebo sa premenoval | vytvor env (časť 1, krok 3), alebo oprav cestu v odkaze |

Kontrola syntaxe po úprave kódu (nič nevypíše = OK):
```
...\Scripts\python.exe -m py_compile ...\DBViewer\db_viewer_ng.pyw
```

### Spolu s viewerom sa otvára terminál
Odkaz spúšťa `python.exe`. V *Vlastnosti* → *Cieľ* klikni do poľa, stlač **Home** a prepíš `python.exe` na **`pythonw.exe`**.

### Okno sa otvorí mimo obrazovky / má zlú veľkosť
Stáva sa po zmene monitorov (napr. doma bez externého monitora). Viewer si pamätá poslednú polohu. Zmaž `%USERPROFILE%\.db_viewer_ng.json`. Prídeš pritom o históriu SQL a nedávne súbory.
Rýchla záchrana bez mazania: klikni na viewer na lište, **Alt+Medzerník** → *Presunúť* → šípky.

### Databázu sa nepodarí otvoriť

| Hláška | Príčina | Riešenie |
|---|---|---|
| `file is not a database` | súbor nie je SQLite (alebo je šifrovaný) | skontroluj, či je to správny súbor |
| `unable to open database file` | sieťový disk nedostupný, zlá cesta, chýbajú práva | skopíruj DB lokálne a otvor kópiu |
| `database is locked` | iný program práve robí veľký zápis | počkaj a daj *Súbor → Znovu načítať schému* (Ctrl+R) |
| Pri názve DB je **„len na čítanie (immutable)“** | súbor je na mieste bez práva zápisu (sieť, CD, read-only priečinok) | funguje, ale **nemusí vidieť posledné zápisy** bežiacej apky (WAL). Na aktuálne dáta skopíruj DB aj s `-wal` súborom lokálne. |

Viewer DB **nikdy nemení**. Otvára ju len na čítanie, `DELETE` / `UPDATE` v SQL skončia chybou „readonly database“.

### Graf sa neotvorí
- Graf je HTML súbor, otvára ho **predvolený prehliadač** Windows. Skontroluj *Nastavenia → Aplikácie → Predvolené aplikácie*.
- Funguje offline (plotly.js je vedľa HTML v `%TEMP%\dbviewer_*`).
- Hláška „priveľa stĺpcov na stĺpcový graf“: zapni *Agregácia X* (deň, mesiac…) alebo použi čiaru.
- Hláška „časová agregácia potrebuje os X ako dátum/čas“: nastav *Prevod X* (text → dátum/čas, Unix s/ms).

### Je to pomalé
| Situácia | Riešenie |
|---|---|
| DB je na sieťovom disku | skopíruj ju lokálne, sieť je 10–100× pomalšia |
| Veľká tabuľka (milióny riadkov) | vyber len potrebné stĺpce vľavo a nechaj zapnuté *vzorkovať* |
| Filter trvá dlho | filtruj podľa stĺpca s **[idx]** alebo **[PK]** (má index). SQL záložka → *Plán dotazu*: `SCAN` = pomalé, `SEARCH … USING INDEX` = rýchle. |
| Dotaz beží priveľmi dlho | **Esc** ho zastaví |
| Prvé otvorenie veľkej DB je pomalé, ďalšie rýchle | normálne, Windows si súbor načíta do cache |

### Neúplné dáta
- V stavovom riadku je **„SKRÁTENÉ na limit“**: zvýš *Limit riadkov* alebo zapni *vzorkovať*.
- Je tam **„vzorkované: každý N. z …“**: v mriežke je len každý N-tý riadok. Na presné hodnoty daj filter na užší rozsah.

---

## 4. Tipy a triky

**Otváranie**
- Pretiahni `.db` súbor na ikonu na ploche a otvorí sa rovno.
- *Súbor → Nedávne* ponúka posledných 10 DB.
- Z terminálu: `pythonw db_viewer_ng.pyw cesta\k\suboru.db`.

**Filter**
- Tlačidlo **hodnoty…** načíta existujúce hodnoty stĺpca (napr. RUN_ID) do zoznamu.
- `IN`: hodnoty oddeľ čiarkou (`2, 3, 5`). `BETWEEN`: dve hodnoty (`10, 20`).
- Časy v texte porovnávaj celým formátom: `2026-03-01 08:30:00.000`.
- Pravý klik na bunku → *Filter: stĺpec = hodnota*.
- **→ do SQL editora** skopíruje vygenerovaný SQL, v editore ho potom dorobíš (JOIN, GROUP BY).

**SQL**
- **Ctrl+Enter** spustí dotaz. Keď označíš časť textu, spustí sa len tá.
- Dvojklik na stĺpec vľavo (v SQL záložke) vloží jeho názov.
- Porovnanie dvoch DB:
  ```
  ATTACH 'C:/cesta/druha.db' AS b;
  SELECT a.TIME, a.SPEED, b.SPEED AS SPEED_B FROM SENSOR a JOIN b.SENSOR b USING (TIME);
  ```
- Hodiny po projektoch a mesiacoch (Výkaz):
  ```
  SELECT substr(date,1,7) AS month, project, SUM(hours) AS hours FROM entries GROUP BY 1,2;
  ```

**Graf**
- Rôzne jednotky (rýchlosť vs. poloha): zaškrtni **Každý stĺpec vo vlastnom grafe**. Nekresli ich do jedného grafu.
- **Rozdeliť podľa** kategórie (projekt, RUN_ID) = jedna krivka alebo farba na hodnotu.
- V grafe: koliesko myši = zoom, dvojklik = pôvodný pohľad, ikona fotoaparátu = PNG.
- Špičky ostávajú aj pri miliónoch bodov (min/max decimácia).

**Mriežka a štatistika**
- Klik na hlavičku = zoradiť, druhý klik = opačne.
- Dvojklik na riadok = detail (stĺpce pod sebou, šípky ←/→ = ďalší riadok).
- **Ctrl+C** skopíruje vybrané riadky aj s hlavičkou, vložíš ich priamo do Excelu.
- **Štatistika** (Ctrl+Shift+S) ukáže aj časovú analýzu: frekvenciu vzorkovania, medzery v dátach, duplicitné a spätné časy.

**Skratky:** Ctrl+O otvoriť · F5 načítať/spustiť · Esc zastaviť · F9 graf · Ctrl+E export · Ctrl+R znovu načítať schému.

---

## 5. Nová verzia programu
1. Prepíš `db_viewer_ng.pyw` novým súborom. Odkaz ani nastavenia sa nemenia.
2. Ak si v kóde robil vlastné úpravy, pred prepísaním si starý súbor premenuj (napr. `db_viewer_ng_old.pyw`).
3. Kontrola: `...\Scripts\python.exe -m py_compile db_viewer_ng.pyw`, potom dvojklik na ikonu.
