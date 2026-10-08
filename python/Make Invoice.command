#!/bin/zsh -l
# Double-click this file, pick your filled-in invoice .xlsx, get the PDF.
cd "$(dirname "$0")"

XLSX=$(osascript -e 'POSIX path of (choose file with prompt "Choose the invoice Excel file" of type {"org.openxmlformats.spreadsheetml.sheet"} default location (POSIX file "'"$PWD"'"))' 2>/dev/null)
if [ -z "$XLSX" ]; then
  echo "No file chosen."; exit 0
fi

PY=python3
for p in "$HOME/anaconda3/bin/python3" "/opt/anaconda3/bin/python3" "$HOME/miniconda3/bin/python3"; do
  [ -x "$p" ] && PY="$p" && break
done

"$PY" -c "import reportlab, openpyxl" 2>/dev/null || \
  "$PY" -m pip install --quiet reportlab openpyxl || \
  "$PY" -m pip install --quiet --user reportlab openpyxl

"$PY" make_invoice.py "$XLSX" && open output
echo
echo "You can close this window."
