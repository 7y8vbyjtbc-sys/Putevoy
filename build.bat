@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul || (echo Python not found. Install Python 3.12 from python.org and tick "Add to PATH". & pause & exit /b 1)
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PutevyeListy --collect-all tkinterdnd2 --add-data "shops.json;." putevoy_gui.py
if errorlevel 1 (echo Build failed. & pause & exit /b 1)
copy /y shops.json dist\shops.json >nul
echo.
echo Done: dist\PutevyeListy.exe
pause
