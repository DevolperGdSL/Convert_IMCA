@echo off
REM ===========================================================================
REM CONVERT_IMCA - Inicializador Automático para Windows (CMD/Batch)
REM Tenta executar o executável compilado (.exe) ou o script Python com fallback
REM ===========================================================================

setlocal enabledelayedexpansion

IF EXIST "%~dp0target\release\convert_imca.exe" (
    "%~dp0target\release\convert_imca.exe" %*
    exit /b %ERRORLEVEL%
)

IF EXIST "%~dp0convert_imca.exe" (
    "%~dp0convert_imca.exe" %*
    exit /b %ERRORLEVEL%
)

where python >nul 2>nul
IF %ERRORLEVEL% EQU 0 (
    python "%~dp0convert_imca.py" %*
    exit /b %ERRORLEVEL%
)

where py >nul 2>nul
IF %ERRORLEVEL% EQU 0 (
    py "%~dp0convert_imca.py" %*
    exit /b %ERRORLEVEL%
)

echo [ERRO] Nao foi possivel localizar convert_imca.exe ou o interpretador Python na maquina.
echo Por favor, instale o Python 3 ou compile o binario Rust com: cargo build --release
exit /b 1
