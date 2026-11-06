@echo off
setlocal
cd /d "%~dp0"
echo Building DOCX to Markdown Converter...
py -m pip install --upgrade pip
if errorlevel 1 goto failed
py -m pip install -r requirements-dev.txt
if errorlevel 1 goto failed
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "DOCX-to-Markdown-Converter" app.py
if errorlevel 1 goto failed
echo.
echo Build completed:
echo %CD%\dist\DOCX-to-Markdown-Converter.exe
pause
exit /b 0
:failed
echo.
echo Build failed. Review the error output above.
pause
exit /b 1
