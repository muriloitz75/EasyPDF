@echo off
echo ===== EXTRATOR DE NOTAS FISCAIS PDF PARA EXCEL =====
echo.

REM Verificar se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python nao encontrado.
    echo.
    echo Por favor, instale o Python 3.8 ou superior:
    echo https://www.python.org/downloads/
    echo.
    echo Apos a instalacao, execute este arquivo novamente.
    echo.
    pause
    exit /b 1
)

REM Verificar se as dependências estão instaladas
python -c "import pandas, pdfplumber, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando dependencias necessarias...
    python -m pip install pandas pdfplumber openpyxl tqdm
)

REM Criar pastas input e output se não existirem
if not exist "input" mkdir input
if not exist "output" mkdir output

REM Executar a aplicação
python executar_app.py %*

REM Se ocorrer um erro, tentar executar diretamente o main.py
if %errorlevel% neq 0 (
    echo.
    echo Tentando executar diretamente o arquivo main.py...
    echo.

    if "%~1"=="" (
        echo Por favor, especifique um arquivo PDF:
        echo.
        echo Exemplo: %~nx0 caminho\para\arquivo.pdf
        echo.
        echo Ou coloque arquivos PDF na pasta "input" e execute novamente.
    ) else (
        python main.py "%~1"
    )
)

echo.
echo Pressione qualquer tecla para sair...
pause > nul
