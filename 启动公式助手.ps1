$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
try {
    $taskRoot = [IO.Path]::GetFullPath($PSScriptRoot)
    $programPath = [IO.Path]::GetFullPath((Join-Path $taskRoot 'OfficeFormulaHelper.exe'))
    $pendingPath = [IO.Path]::GetFullPath((Join-Path $taskRoot 'OfficeFormulaHelper.next.exe'))
    if ([IO.Path]::GetDirectoryName($programPath) -ne $taskRoot -or [IO.Path]::GetDirectoryName($pendingPath) -ne $taskRoot) { throw 'Invalid application path.' }
    $env:TEMP = Join-Path $taskRoot 'temp'
    $env:TMP = $env:TEMP
    if (Test-Path -LiteralPath $pendingPath) {
        $active = Get-Process -Name OfficeFormulaHelper -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $programPath }
        if ($active) { [Windows.Forms.MessageBox]::Show('请先关闭正在运行的旧版公式助手，再点击本启动入口。编辑框中的内容请先自行保留。','公式助手更新') | Out-Null; exit }
        Copy-Item -LiteralPath $pendingPath -Destination $programPath -Force
        if ((Get-FileHash -LiteralPath $pendingPath).Hash -ne (Get-FileHash -LiteralPath $programPath).Hash) { throw 'Application verification failed.' }
        Remove-Item -LiteralPath $pendingPath -Force
    }
    Start-Process -FilePath $programPath
} catch {
    [Windows.Forms.MessageBox]::Show($_.Exception.Message,'公式助手启动失败') | Out-Null
}
