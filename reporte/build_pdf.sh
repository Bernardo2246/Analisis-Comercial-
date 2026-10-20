#!/bin/bash
# Genera reporte/Reporte_Comercial_Q3-2026.pdf desde reporte/reporte.html (Chrome headless).
# Uso: bash reporte/build_pdf.sh   (requiere conexión para el mapa)
set -e
cd "$(dirname "$0")/.."
python3 data/build.py
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
"$CHROME" --headless=new --disable-gpu --no-pdf-header-footer --virtual-time-budget=20000 \
  --user-data-dir="$(mktemp -d)" --print-to-pdf="reporte/Reporte_Comercial_Q3-2026.pdf" \
  "file://$PWD/reporte/reporte.html" 2>/dev/null
echo "PDF: reporte/Reporte_Comercial_Q3-2026.pdf"
