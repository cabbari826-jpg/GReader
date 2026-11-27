@echo off
setlocal
cd /d "%~dp0"

if not exist .venv (
  py -3.11 -m venv .venv
  if errorlevel 1 goto :error
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements-dev.txt
if errorlevel 1 goto :error

REM Optional: set TESSERACT_DIR to a portable Tesseract directory.
if defined TESSERACT_DIR (
  if exist "%TESSERACT_DIR%\tesseract.exe" (
    if not exist resources\tesseract mkdir resources\tesseract
    xcopy "%TESSERACT_DIR%\*" "resources\tesseract\" /E /I /Y >nul
  ) else (
    echo WARNING: TESSERACT_DIR was set but tesseract.exe was not found.
  )
)

set "TESSERACT_DATA_ARG="
if exist "resources\tesseract\tesseract.exe" set TESSERACT_DATA_ARG=--add-data "resources\tesseract;resources\tesseract"

pyinstaller --noconfirm --clean --windowed --onedir ^
  --name PDFOrderValidator ^
  --collect-all PySide6 --collect-all fitz --collect-all openpyxl --collect-all PIL ^
  --hidden-import pytesseract --hidden-import win32com.client --hidden-import pythoncom ^
  %TESSERACT_DATA_ARG% ^
  app.py
if errorlevel 1 goto :error

echo.
echo Build complete: dist\PDFOrderValidator\PDFOrderValidator.exe
echo Compile installer.iss with Inno Setup 6 to create the setup executable.
pause
exit /b 0

:error
echo Build failed. See the preceding error.
pause
exit /b 1
