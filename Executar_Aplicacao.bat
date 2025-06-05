@echo off
title Extrator de Notas Fiscais

echo =====================================================
echo  EXTRATOR DE NOTAS FISCAIS PDF PARA EXCEL
echo =====================================================
echo.

REM Verifica se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python nao esta instalado ou nao foi encontrado no PATH.
    echo.
    echo Por favor, instale o Python 3.8 ou superior a partir de:
    echo https://www.python.org/downloads/
    echo.
    echo Certifique-se de marcar a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo Verificando e iniciando a aplicacao...
echo.

REM Executa o script principal da GUI
python executar_app.py

echo.
echo Feche esta janela ou pressione qualquer tecla para sair.
pause > nul
