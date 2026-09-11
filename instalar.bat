@echo off
echo Instalando dependencias para Registro de Produccion...
cd /d "%~dp0"
pip install -r requirements.txt
echo.
echo Instalacion completada. Presiona cualquier tecla para cerrar.
pause > nul
