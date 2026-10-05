@echo off
REM ===========================================================================
REM BUILD_WINDOWS_GUI.BAT - Compilador do Executavel Windows (.exe)
REM Projeto: Convert_IMCA (v1.1.0)
REM Gera: dist\Convert_IMCA_GUI.exe (Executavel autonomo com icone, sem console)
REM ===========================================================================

setlocal enabledelayedexpansion

echo ============================================================
echo  [BUILD] Gerando Convert_IMCA_GUI.exe para Windows
echo ============================================================

where python >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Python 3 nao encontrado no PATH do sistema.
    echo Instale o Python em https://www.python.org/downloads/
    exit /b 1
)

echo [1/3] Verificando PyInstaller e PyWebView...
python -m pip install --upgrade pyinstaller Pillow pywebview >nul 2>&1

echo [2/3] Compilando executavel sem janela de console...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "Convert_IMCA_GUI" ^
    --icon "assets\icon.ico" ^
    --add-data "assets;assets" ^
    convert_imca_gui.py

IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha durante a compilacao com PyInstaller.
    exit /b %ERRORLEVEL%
)

echo [3/3] Concluido com sucesso!
echo ============================================================
echo Executavel pronto em: dist\Convert_IMCA_GUI.exe
echo ============================================================
exit /b 0
