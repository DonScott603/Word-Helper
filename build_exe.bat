@echo off
rem Build a standalone single-file WordHelper.exe into the dist\ folder.
rem Requires: py -m pip install pyinstaller
cd /d "%~dp0"
py -m PyInstaller --noconfirm --onefile --windowed --name "WordHelper" ^
    --collect-all customtkinter --collect-data docx word_helper.py
echo.
echo Done. The executable is at: dist\WordHelper.exe
pause
