@echo off
REM Drag a filled-in invoice .xlsx onto this file to make the PDF.
cd /d "%~dp0"
python -m pip install --quiet --disable-pip-version-check reportlab openpyxl
if "%~1"=="" (
  echo Drag your invoice .xlsx file onto "Make Invoice.bat".
) else (
  python make_invoice.py "%~1"
  if not errorlevel 1 start "" "%~dp0output"
)
pause
