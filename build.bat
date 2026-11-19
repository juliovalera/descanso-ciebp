@echo off
setlocal
chcp 65001 > nul

set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
set "BUILD_DRIVE="
set "BUILD_ENV=.build-env"
set "BUILD_EXIT=0"

rem Escolhe uma letra livre para que uma execucao interrompida nao bloqueie o proximo build.
for %%D in (Q R S T U V W X Y Z) do (
    if not defined BUILD_DRIVE if not exist %%D:\NUL set "BUILD_DRIVE=%%D:"
)

if not defined BUILD_DRIVE (
    echo.
    echo  ERRO: Nao foi encontrada uma letra de unidade livre para a compilacao.
    pause
    exit /b 1
)

subst %BUILD_DRIVE% "%PROJECT_DIR%"
if errorlevel 1 (
    echo.
    echo  ERRO: Nao foi possivel criar a unidade virtual de build.
    pause
    exit /b 1
)

pushd %BUILD_DRIVE%\
if errorlevel 1 (
    echo.
    echo  ERRO: Nao foi possivel acessar a unidade virtual de build.
    subst %BUILD_DRIVE% /d > nul 2>nul
    pause
    exit /b 1
)

echo.
echo  =========================================================
echo   Gerando executavel da Tela de Espera
echo  =========================================================
echo.

echo [1/5] Preparando ambiente isolado de compilacao...
if exist "%BUILD_ENV%" rmdir /s /q "%BUILD_ENV%"
python -m venv "%BUILD_ENV%"
if errorlevel 1 (
    echo.
    echo  ERRO: Falha ao criar o ambiente isolado de compilacao.
    set "BUILD_EXIT=1"
    goto :cleanup
)

echo [2/5] Instalando dependencias de build...
"%BUILD_ENV%\Scripts\python.exe" -m pip install --quiet --disable-pip-version-check --upgrade pyinstaller pillow pygame-ce pywin32
if errorlevel 1 (
    echo.
    echo  ERRO: Falha ao instalar dependencias.
    set "BUILD_EXIT=1"
    goto :cleanup
)

echo [3/5] Compilando...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q "CIEBP_Descanso_de_Tela.spec" 2>nul
"%BUILD_ENV%\Scripts\python.exe" -m PyInstaller ^
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
    set "BUILD_EXIT=1"
    goto :cleanup
)

echo [4/5] Copiando assets e configuracoes...
copy /y config.json "dist\CIEBP_Descanso_de_Tela\" > nul
if exist assets xcopy /e /i /y assets "dist\CIEBP_Descanso_de_Tela\assets" > nul

echo [5/5] Limpando arquivos temporarios...

:cleanup
rmdir /s /q build 2>nul
del /q "CIEBP_Descanso_de_Tela.spec" 2>nul
rmdir /s /q "%BUILD_ENV%" 2>nul
popd
subst %BUILD_DRIVE% /d > nul 2>nul

if "%BUILD_EXIT%"=="1" (
    pause
    exit /b 1
)

echo.
echo  =========================================================
echo   Concluido com sucesso!
echo  =========================================================
echo.
echo   Pasta para distribuicao:
echo   dist\CIEBP_Descanso_de_Tela\
echo.
echo   Copie esta pasta inteira e execute o arquivo .exe no destino.
echo.
pause
exit /b 0
