$ErrorActionPreference = 'Stop'
try {
    $env:TEMP = Join-Path $PSScriptRoot 'temp'
    $env:TMP = $env:TEMP
    New-Item -ItemType Directory -Path $env:TEMP,(Join-Path $PSScriptRoot 'output') -Force | Out-Null
    $app = Join-Path $PSScriptRoot 'OfficeFormulaHelper.exe'
    $update = Join-Path $env:TEMP 'OfficeFormulaHelper.update.exe'
    if (Test-Path -LiteralPath $update) {
        $running = Get-Process -Name OfficeFormulaHelper -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $app }
        if ($running) { throw '请先关闭正在运行的旧助手，再点击本启动入口完成更新。' }
        Copy-Item -LiteralPath $update -Destination $app -Force
        if ((Get-FileHash -LiteralPath $app).Hash -ne (Get-FileHash -LiteralPath $update).Hash) { throw '更新校验失败。' }
        Remove-Item -LiteralPath $update -Force
    }
    Start-Process -FilePath $app
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show($_.Exception.Message,'通用公式助手 v1.0.0') | Out-Null
}