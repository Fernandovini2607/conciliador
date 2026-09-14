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
echo === Instalando/atualizando dependencias ===
python -m pip install -r requirements.txt --quiet --disable-pip-version-check

REM Se ainda nao existe db_config.json, roda o setup do banco
if not exist "data\db_config.json" (
    echo.
    echo === Primeira execucao: configurando MariaDB ===
    python setup_db.py
    if errorlevel 1 (
        echo.
        echo [ERRO] Setup do banco falhou. Corrija e rode iniciar.bat de novo.
        pause
        exit /b 1
    )
)

echo.
echo === Iniciando Conciliador ===
python main.py

REM Se o app cair com erro, mantem a janela aberta para voce ler a mensagem
if errorlevel 1 (
    echo.
    echo [ERRO] O app fechou com codigo de erro. Veja a mensagem acima.
    pause
)
