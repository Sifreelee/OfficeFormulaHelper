$ErrorActionPreference = 'Stop'
try {
    $env:TEMP = Join-Path $PSScriptRoot 'temp'
    $env:TMP = $env:TEMP
    New-Item -ItemType Directory -Path $env:TEMP,(Join-Path $PSScriptRoot 'output') -Force | Out-Null
    Start-Process -FilePath (Join-Path $PSScriptRoot 'OfficeFormulaHelper.exe')
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show($_.Exception.Message,'通用公式助手 v1.0.0') | Out-Null
}