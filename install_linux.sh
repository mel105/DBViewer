#!/bin/sh
# Vytvorí spúšťač "DB Viewer" v menu aplikácií a na ploche (bez sudo).
DIR="$(cd "$(dirname "$0")" && pwd)"
PY="$(command -v python3)"
"$PY" -c "import tkinter, pandas, numpy, plotly" 2>/dev/null || {
  echo "Chýba niektorý balík. Skontroluj: $PY -c 'import tkinter, pandas, numpy, plotly'"
  echo "tkinter na Ubuntu/Debian: sudo apt install python3-tk"
  exit 1; }
APPS="$HOME/.local/share/applications"
mkdir -p "$APPS"
cat > "$APPS/db-viewer.desktop" <<EOD
[Desktop Entry]
Type=Application
Name=DB Viewer
Comment=Prehliadač SQLite databáz
Exec="$PY" "$DIR/db_viewer_ng.pyw" %f
Icon=$DIR/db_viewer.png
Terminal=false
Categories=Development;Database;
MimeType=application/vnd.sqlite3;application/x-sqlite3;
EOD
chmod +x "$APPS/db-viewer.desktop"
update-desktop-database "$APPS" 2>/dev/null
DESK="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")"
if [ -d "$DESK" ]; then
  cp "$APPS/db-viewer.desktop" "$DESK/"
  chmod +x "$DESK/db-viewer.desktop"
  gio set "$DESK/db-viewer.desktop" metadata::trusted true 2>/dev/null
fi
echo "Hotovo: DB Viewer je v menu aplikácií a na ploche ($DESK)."
