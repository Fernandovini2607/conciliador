@echo off
REM Inicia o Conciliador: atualiza codigo, ativa venv e roda o app.
REM Clique duplo neste arquivo no Explorer para abrir.

cd /d "%~dp0"

echo === Atualizando codigo do GitHub ===
git pull

echo.
echo === Ativando ambiente virtual ===
call .\.venv\Scripts\activate.bat

echo.
echo === Iniciando Conciliador ===
python main.py

REM Se o app cair com erro, mantem a janela aberta para voce ler a mensagem
if errorlevel 1 (
    echo.
    echo [ERRO] O app fechou com codigo de erro. Veja a mensagem acima.
    pause
)
