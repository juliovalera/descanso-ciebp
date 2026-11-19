@echo off
setlocal
chcp 65001 > nul

set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
set "BUILD_DRIVE="

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

echo [1/4] Instalando dependencias de build...
python -m pip install --quiet --upgrade pyinstaller pillow pygame-ce pywin32
if errorlevel 1 (
    echo.
    echo  ERRO: Falha ao instalar dependencias.
    popd
    subst %BUILD_DRIVE% /d > nul 2>nul
    pause
    exit /b 1
)

echo [2/4] Compilando...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q "CIEBP_Descanso_de_Tela.spec" 2>nul
python -m PyInstaller ^
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
    pause
    exit /b 1
)

echo [3/4] Copiando assets e configuracoes...
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
echo   Copie esta pasta inteira e execute o arquivo .exe no destino.
echo.
pause
exit /b 0
