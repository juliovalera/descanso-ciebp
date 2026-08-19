@echo off
setlocal
chcp 65001 > nul
set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
set "BUILD_DRIVE=Q:"

if exist %BUILD_DRIVE%\NUL (
    echo.
    echo  ERRO: A unidade %BUILD_DRIVE% ja esta em uso.
    echo  Feche ou libere essa unidade e tente novamente.
    pause & exit /b 1
)

subst %BUILD_DRIVE% "%PROJECT_DIR%"
if errorlevel 1 (
    echo.
    echo  ERRO: Nao foi possivel criar a unidade virtual de build.
    pause & exit /b 1
)

pushd %BUILD_DRIVE%\
if errorlevel 1 (
    echo.
    echo  ERRO: Nao foi possivel acessar a unidade virtual de build.
    subst %BUILD_DRIVE% /d > nul 2>nul
    pause & exit /b 1
)

echo.
echo  =========================================================
echo   CIEBP - Gerando executavel do Descanso de Tela
echo  =========================================================
echo.

echo [1/4] Instalando dependencias de build...
pip install --quiet --upgrade pyinstaller pillow pygame-ce pywin32
if errorlevel 1 (
    echo.
    echo  ERRO: Falha ao instalar dependencias.
    pause & exit /b 1
)

echo [2/4] Compilando...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q "CIEBP_Descanso_de_Tela.spec" 2>nul
pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "CIEBP_Descanso_de_Tela" ^
    --hidden-import "win32com.client" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pygame" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageTk" ^
    --hidden-import "PIL.ImageFilter" ^
    --hidden-import "PIL.ImageDraw" ^
    --collect-all "pygame" ^
    main.py

if errorlevel 1 (
    echo.
    echo  ERRO: Falha na compilacao. Verifique as mensagens acima.
    popd
    subst %BUILD_DRIVE% /d > nul 2>nul
    pause & exit /b 1
)

echo [3/4] Copiando assets e config para a pasta dist...
copy /y config.json "dist\CIEBP_Descanso_de_Tela\" > nul
if exist assets xcopy /e /i /y assets "dist\CIEBP_Descanso_de_Tela\assets" > nul

echo [4/4] Limpando arquivos temporarios...
rmdir /s /q build 2>nul
del /q "CIEBP_Descanso_de_Tela.spec" 2>nul
popd
subst %BUILD_DRIVE% /d > nul 2>nul

echo.
echo  =========================================================
echo   Concluido com sucesso!
echo  =========================================================
echo.
echo   Pasta para distribuicao:
echo   dist\CIEBP_Descanso_de_Tela\
echo.
echo   Copie esta pasta inteira (pen drive, Google Drive...)
echo   e execute CIEBP_Descanso_de_Tela.exe no destino.
echo.
pause
exit /b 0


echo.
echo  =========================================================
echo   CIEBP - Gerando pacote de distribuicao
echo  =========================================================
echo.

set DIST=dist\CIEBP_Descanso_de_Tela
set PY_DIR=%DIST%\python
set PY_ZIP=python-embed.zip
set PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
set PIP_URL=https://bootstrap.pypa.io/get-pip.py

:: ── Limpa e cria estrutura de pastas ──────────────
if exist "%DIST%" rmdir /s /q "%DIST%"
mkdir "%DIST%"
mkdir "%PY_DIR%"
mkdir "%DIST%\assets"

:: ── Download do Python embarcado ──────────────────
echo [1/5] Baixando Python 3.11.9 embarcado (portátil)...
powershell -Command "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_ZIP%'" 2>nul
if not exist "%PY_ZIP%" (
    echo.
    echo  ERRO: Nao foi possivel baixar o Python.
    echo  Verifique sua conexao com a internet.
    pause & exit /b 1
)

:: ── Extrai Python ──────────────────────────────────
echo [2/5] Extraindo Python...
powershell -Command "Expand-Archive -Force '%PY_ZIP%' '%PY_DIR%'"
del /q "%PY_ZIP%"

:: Habilita site-packages no Python embarcado
for %%f in ("%PY_DIR%\python3*._pth") do (
    powershell -Command "(Get-Content '%%f') -replace '#import site','import site' | Set-Content '%%f'"
)

:: ── Instala pip ────────────────────────────────────
echo [3/5] Instalando pip...
powershell -Command "Invoke-WebRequest -Uri '%PIP_URL%' -OutFile 'get-pip.py'" 2>nul
"%PY_DIR%\python.exe" get-pip.py --quiet 2>nul
del /q get-pip.py 2>nul

:: ── Instala dependencias do projeto ────────────────
echo [4/5] Instalando dependencias do projeto...
"%PY_DIR%\python.exe" -m pip install --quiet Pillow pygame-ce

if errorlevel 1 (
    echo.
    echo  ERRO: Falha ao instalar as dependencias do projeto.
    pause & exit /b 1
)

:: ── Copia o projeto ────────────────────────────────
echo [5/5] Copiando arquivos do projeto...
copy /y main.py "%DIST%\" > nul
copy /y config.json "%DIST%\" > nul
if exist assets xcopy /e /i /y assets "%DIST%\assets" > nul

:: ── Cria o launcher para o usuário ─────────────────
(
echo @echo off
echo start "" "%%~dp0python\pythonw.exe" "%%~dp0main.py"
) > "%DIST%\Iniciar.bat"

:: Ícone de atalho opcional na área de trabalho
powershell -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut([System.IO.Path]::Combine($ws.SpecialFolders('Desktop'),'CIEBP Descanso de Tela.lnk')); $s.TargetPath=[System.IO.Path]::GetFullPath('%DIST%\Iniciar.bat'); $s.WorkingDirectory=[System.IO.Path]::GetFullPath('%DIST%'); $s.Description='CIEBP Descanso de Tela'; $s.Save()" 2>nul

echo.
echo  =========================================================
echo   Concluido com sucesso!
echo  =========================================================
echo.
echo   Pasta para distribuicao:
echo   %DIST%\
echo.
echo   Como distribuir:
echo   1. Copie a pasta %DIST%\ inteira
echo      (pen drive, Google Drive, rede local).
echo   2. No computador de destino, clique duas vezes em:
echo      Iniciar.bat (ou use o atalho criado na area de trabalho)
echo.
echo   Nenhuma instalacao necessaria no computador de destino.
echo.
pause
