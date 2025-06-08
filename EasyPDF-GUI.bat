@echo off
echo Iniciando EasyPDF Extractor GUI...
cd /d "%~dp0"
python gui.py
if errorlevel 1 (
    echo.
    echo Erro ao executar a aplicacao!
    echo Verifique se o Python esta instalado e as dependencias estao corretas.
    pause
)