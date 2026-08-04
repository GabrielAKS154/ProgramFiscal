@echo off
echo ============================================
echo  XML Validator - INSTALACAO (so uma vez!)
echo ============================================
echo.
echo Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado!
    echo Instale em: https://www.python.org/downloads/
    echo Marque a opcao Add Python to PATH durante a instalacao.
    pause
    exit /b 1
)
echo Python encontrado!
echo.
echo Instalando dependencias...
pip install openpyxl
echo.
echo ============================================
echo  Instalacao concluida com sucesso!
echo  Da proxima vez, use apenas: RODAR.bat
echo ============================================
pause
