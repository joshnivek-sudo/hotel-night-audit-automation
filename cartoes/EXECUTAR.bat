@echo off
chcp 65001 >nul
cd /d "%~dp0"

python preencher_cartoes.py %1

if errorlevel 1 (
    echo.
    echo ============================================================
    echo Deu algum erro ao rodar o script.
    echo Se a mensagem acima for parecida com "python nao e reconhecido",
    echo o Python nao esta instalado ou nao esta no PATH deste PC.
    echo Se for erro de "No module named openpyxl" ou parecido, rode:
    echo     python -m pip install openpyxl xlrd pypdf
    echo ============================================================
    pause
)
