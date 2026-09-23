@echo off
REM Roda o Conciliador em modo debug pra uma NF especifica.
REM Gera arquivo debug_dominio.log na pasta com detalhes do que
REM aconteceu com a NF durante a conciliacao.
REM
REM Uso: edite a variavel NF_ALVO abaixo, salve e clique duplo.

REM ============ CONFIGURE A NF QUE QUER DEPURAR ============
set NF_ALVO=276158
REM =========================================================

cd /d "%~dp0"

echo === Atualizando codigo do GitHub ===
git pull

echo.
echo === Ativando ambiente virtual ===
call .\.venv\Scripts\activate.bat

echo.
echo === Ativando modo debug pra NF %NF_ALVO% ===
set DEBUG_NF=%NF_ALVO%

echo.
echo === Iniciando Conciliador em modo debug ===
echo O log sera salvo em debug_dominio.log
python main.py

echo.
echo === Modo debug encerrado ===
echo Verifique o arquivo debug_dominio.log na pasta.
pause
