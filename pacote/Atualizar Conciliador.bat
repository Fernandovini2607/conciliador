@echo off
rem =====================================================================
rem  Atualizar Conciliador
rem
rem  Copia a versao publicada na pasta de rede por cima desta instalacao.
rem    - Preserva "data\dominio_config.json" (suas credenciais do Dominio).
rem    - Apaga arquivos que sairam da versao nova, pra nao sobrar DLL velha.
rem    - Exige o Conciliador FECHADO.
rem
rem  Uso normal: duplo-clique.
rem  Uso avulso: Atualizar Conciliador.bat "\\outro\caminho\atual"
rem
rem  Se a pasta de rede mudar de endereco, edite ORIGEM em
rem  pacote\Atualizar Conciliador.bat no projeto e publique de novo.
rem =====================================================================
setlocal

if /i "%~1"=="--relancado" goto :relancado

set "ORIGEM=\\10.0.1.47\conciliador\atual"
if not "%~1"=="" set "ORIGEM=%~1"

set "DESTINO=%~dp0"
if "%DESTINO:~-1%"=="\" set "DESTINO=%DESTINO:~0,-1%"

rem Este .bat vive dentro da pasta que sera sobrescrita. O cmd le o
rem arquivo conforme executa, entao rodamos uma copia no %TEMP% pra que a
rem atualizacao possa substituir o original sem embaralhar o script.
copy /y "%~f0" "%TEMP%\AtualizarConciliador.bat" >nul 2>&1
if errorlevel 1 goto :executar
call "%TEMP%\AtualizarConciliador.bat" --relancado "%DESTINO%" "%ORIGEM%"
exit /b %errorlevel%

:relancado
set "DESTINO=%~2"
set "ORIGEM=%~3"

:executar
echo.
echo   Atualizando o Conciliador
echo   De ..: %ORIGEM%
echo   Para : %DESTINO%
echo.

tasklist /fi "imagename eq Conciliador.exe" | find /i "Conciliador.exe" >nul
if not errorlevel 1 goto :app_aberto

if not exist "%ORIGEM%\Conciliador.exe" goto :sem_rede

robocopy "%ORIGEM%" "%DESTINO%" /MIR /XF dominio_config.json debug_dominio.log /XD backups /R:2 /W:5 /NP /NFL /NDL
if errorlevel 8 goto :falhou

echo.
echo   [OK] Atualizado.
if exist "%DESTINO%\VERSAO.txt" type "%DESTINO%\VERSAO.txt"
echo.
echo   Pode abrir o Conciliador normalmente.
echo.
pause
exit /b 0

:app_aberto
echo   [ERRO] O Conciliador esta aberto.
echo          Feche o programa e rode este atalho de novo.
echo.
pause
exit /b 1

:sem_rede
echo   [ERRO] Nao consegui ler a pasta de rede acima.
echo          Verifique a conexao com o servidor ou avise o administrador.
echo.
pause
exit /b 1

:falhou
echo.
echo   [ERRO] A copia falhou no meio do caminho.
echo          NAO use o programa: avise o administrador.
echo.
pause
exit /b 1
