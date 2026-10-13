@echo off
cd /d "%~dp0"
echo ==========================================
echo   Creation de FDL.exe (une seule fois)
echo ==========================================
py -m pip install --upgrade pyinstaller python-docx openpyxl segno
if errorlevel 1 goto erreur
py -m PyInstaller --noconfirm --clean --onefile --windowed --name FDL --icon "assets\fdl.ico" --add-data "assets;assets" --collect-data docx --collect-data openpyxl --collect-all segno fdl_app.py
if errorlevel 1 goto erreur
echo.
echo Termine !  Le programme est ici :  dist\FDL.exe
echo Vous pouvez copier FDL.exe ou vous voulez (bureau, cle USB...).
pause
exit /b 0
:erreur
echo.
echo Une erreur est survenue. Envoyez-moi une capture de cette fenetre.
pause
exit /b 1
