@echo off
rem =====================================================================
rem  publicar.bat - gera o pacote de distribuicao do Conciliador.
rem
rem  1. Carimba a versao do dia em versao.py (aparece na tela do app)
rem  2. Roda o PyInstaller (Conciliador.spec)
rem  3. Monta o pacote: LEIA-ME, VERSAO.txt, atualizador e pasta data\
rem     no nivel do .exe (o app le data\ ao lado do executavel)
rem  4. Gera dist\Conciliador_AAAA-MM-DD.zip
rem  5. Publica na pasta de rede, se estiver acessivel
rem
rem  Uso: publicar.bat  [pasta_de_rede]
rem       (sem argumento usa o padrao abaixo)
rem =====================================================================
setlocal

cd /d "%~dp0"

set "DESTINO=\\10.0.1.47\conciliador\atual"
if not "%~1"=="" set "DESTINO=%~1"

set "PKG=dist\Conciliador"

if not exist ".venv\Scripts\pyinstaller.exe" goto :sem_venv

for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd'"`) do set "DATA=%%d"
set "VERSAO=%DATA:-=.%"
set "ZIP=dist\Conciliador_%DATA%.zip"

echo.
echo  === Conciliador v%VERSAO% ===
echo.

echo  [1/5] Carimbando a versao em versao.py...
powershell -NoProfile -Command "$q=[char]34; $p='versao.py'; $t=Get-Content -Raw -Encoding UTF8 $p; $t=[regex]::Replace($t,'(?m)^VERSAO = .+$',('VERSAO = '+$q+'%VERSAO%'+$q)); [IO.File]::WriteAllText((Resolve-Path $p).Path,$t)"
if errorlevel 1 goto :falhou

echo  [2/5] Rodando o PyInstaller...
".venv\Scripts\pyinstaller.exe" Conciliador.spec --noconfirm --clean
if errorlevel 1 goto :falhou

echo  [3/5] Montando o pacote...
copy /y "pacote\LEIA-ME.txt" "%PKG%\" >nul
if errorlevel 1 goto :falhou
copy /y "pacote\Atualizar Conciliador.bat" "%PKG%\" >nul
if errorlevel 1 goto :falhou
if not exist "%PKG%\data" mkdir "%PKG%\data"
rem data\ no nivel do .exe: o spec entrega os datas em _internal\, mas em
rem runtime db.py e parser_dominio.py leem data\ ao lado do executavel.
copy /y "data\db_config.json" "%PKG%\data\" >nul
if errorlevel 1 goto :falhou
copy /y "pacote\dominio_config.json.EXEMPLO" "%PKG%\data\" >nul
if errorlevel 1 goto :falhou
> "%PKG%\VERSAO.txt" echo Conciliador v%VERSAO% - build de %DATA%

echo  [4/5] Gerando %ZIP%...
powershell -NoProfile -Command "$ErrorActionPreference='Stop'; if (Test-Path '%ZIP%') { Remove-Item -LiteralPath '%ZIP%' -Force }; Compress-Archive -Path '%PKG%\*' -DestinationPath '%ZIP%' -CompressionLevel Optimal"
if errorlevel 1 goto :falhou

echo  [5/5] Publicando em %DESTINO% ...
if not exist "%DESTINO%" mkdir "%DESTINO%" 2>nul
if not exist "%DESTINO%" goto :sem_rede
robocopy "%PKG%" "%DESTINO%" /MIR /R:2 /W:5 /NP /NFL /NDL
if errorlevel 8 goto :falha_publicacao
for %%i in ("%DESTINO%") do set "HIST=%%~dpihistorico"
if not exist "%HIST%" mkdir "%HIST%" 2>nul
if exist "%HIST%" copy /y "%ZIP%" "%HIST%\" >nul

echo.
echo  [OK] Versao %VERSAO% publicada.
echo       Pacote ..: %ZIP%
echo       Rede ....: %DESTINO%
echo.
echo  Nas maquinas dos operadores: duplo-clique em
echo  "Atualizar Conciliador.bat" dentro de C:\Conciliador.
echo.
goto :fim

:sem_rede
echo.
echo  [AVISO] Pacote gerado, mas a pasta de rede nao esta acessivel:
echo          %DESTINO%
echo          Publique depois com: publicar.bat "%DESTINO%"
echo          ou distribua o ZIP: %ZIP%
echo.
goto :fim

:falha_publicacao
echo.
echo  [ERRO] O ZIP foi gerado, mas a copia pra rede falhou.
echo         Verifique permissao de escrita em %DESTINO%.
echo.
exit /b 1

:sem_venv
echo  [ERRO] Nao achei .venv\Scripts\pyinstaller.exe
echo         Ative o ambiente do projeto e instale o PyInstaller.
exit /b 1

:falhou
echo.
echo  [ERRO] A geracao falhou no passo acima. Nada foi publicado.
echo.
exit /b 1

:fim
pause
