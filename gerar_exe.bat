@echo off
echo ============================================
echo  Gerando EXE standalone (sem precisar Python)
echo ============================================
echo.
pip install pyinstaller openpyxl
pyinstaller --onefile --windowed --name=XMLValidator xml_validator.py
echo.
echo EXE gerado em: dist\XMLValidator.exe
echo Copie o EXE para qualquer maquina Windows e use!
pause
