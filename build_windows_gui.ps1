<#
===========================================================================
BUILD_WINDOWS_GUI.PS1 - Compilador do Executável Windows (.exe)
Projeto: Convert_IMCA (v1.1.0)
Gera: dist\Convert_IMCA_GUI.exe
===========================================================================
#>

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " 🚀 COMPILANDO CONVERT_IMCA PARA WINDOWS (.EXE)" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python 3 não está instalado ou não está no PATH."
    exit 1
}

Write-Host "📦 [1/3] Verificando dependências de compilação (PyInstaller, Pillow, pywebview)..." -ForegroundColor Yellow
python -m pip install --upgrade pyinstaller Pillow pywebview | Out-Null

Write-Host "⚙️ [2/3] Executando PyInstaller (--onefile --windowed)..." -ForegroundColor Yellow
python -m PyInstaller --noconfirm --onefile --windowed `
    --name "Convert_IMCA_GUI" `
    --icon "assets\icon.ico" `
    --add-data "assets;assets" `
    convert_imca_gui.py

if ($LASTEXITCODE -ne 0) {
    Write-Error "Falha na compilação do executável."
    exit $LASTEXITCODE
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "✨ SUCESSO! Executável Windows gerado com sucesso:" -ForegroundColor Green
Write-Host "   📍 dist\Convert_IMCA_GUI.exe" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
