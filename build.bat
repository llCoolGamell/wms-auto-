@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === WMS Automator: установка зависимостей ===
python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

echo.
echo === Сборка exe ===
python -m PyInstaller --noconfirm --onefile --windowed ^
  --name "WMS_Automator" ^
  --distpath "dist" ^
  wms_auto.py

if errorlevel 1 (
  echo Сборка не удалась.
  exit /b 1
)

echo.
echo Готово: dist\WMS_Automator.exe
pause
