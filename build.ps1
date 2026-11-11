param([switch]$Test)
$ErrorActionPreference = 'Stop'
$taskRoot = $PSScriptRoot
$env:TEMP = Join-Path $taskRoot 'temp'
$env:TMP = $env:TEMP
New-Item -ItemType Directory -Path $env:TEMP,(Join-Path $taskRoot 'output') -Force | Out-Null
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) { throw 'Missing .NET Framework 4.x compiler.' }
$app = Join-Path $taskRoot 'OfficeFormulaHelper.exe'
$buildArgs = @('/nologo','/target:winexe','/platform:x64','/reference:System.Windows.Forms.dll','/reference:System.Drawing.dll','/reference:System.Web.Extensions.dll','/reference:Microsoft.CSharp.dll','/reference:System.Xml.dll','/reference:System.IO.Compression.dll','/reference:System.IO.Compression.FileSystem.dll',('/out:' + $app))
$buildArgs += @('FormulaHelper.cs','NativeMath.cs','ClipboardLatex.cs','AssemblyInfo.cs') | ForEach-Object { Join-Path $taskRoot $_ }
& $csc @buildArgs
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
if ($Test) {
    foreach ($mode in @('--self-test-extraction','--self-test-wps')) {
        $run = Start-Process -FilePath $app -ArgumentList $mode -WindowStyle Hidden -Wait -PassThru
        if ($run.ExitCode -ne 0) { throw "Verification failed: $mode" }
    }
}
Write-Output 'OfficeFormulaHelper v1.0.0 build complete.'