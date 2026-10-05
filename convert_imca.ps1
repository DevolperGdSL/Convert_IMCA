<#
===========================================================================
CONVERT_IMCA - Inicializador Automático para Windows (PowerShell)
Executa o binário compilado .exe ou o script Python nativo com fallback.
===========================================================================
#>

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ExeTarget = Join-Path $ScriptDir "target\release\convert_imca.exe"
$ExeLocal = Join-Path $ScriptDir "convert_imca.exe"
$PyScript = Join-Path $ScriptDir "convert_imca.py"

if (Test-Path $ExeTarget) {
    & $ExeTarget @Arguments
    exit $LASTEXITCODE
}

if (Test-Path $ExeLocal) {
    & $ExeLocal @Arguments
    exit $LASTEXITCODE
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    & python $PyScript @Arguments
    exit $LASTEXITCODE
}

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py $PyScript @Arguments
    exit $LASTEXITCODE
}

Write-Error "Não foi encontrado nem o executável compilado (.exe) nem o runtime Python."
exit 1
