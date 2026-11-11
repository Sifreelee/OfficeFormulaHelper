param([string]$OutputPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$taskRoot = [IO.Path]::GetFullPath($PSScriptRoot)
$version = ([IO.File]::ReadAllText((Join-Path $taskRoot 'VERSION'))).Trim()
if ($version -notmatch '^v\d+\.\d+\.\d+$') { throw 'Invalid VERSION.' }
$packageRoot = 'Office公式助手_' + $version
if (-not $OutputPath) { $OutputPath = Join-Path (Split-Path $taskRoot -Parent) ($packageRoot + '.zip') }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$mustHave = @('OfficeFormulaHelper.exe','render.cjs','runtime.json','vendor/node/node.exe','vendor/package/js/mathjax.js','vendor/package/js/input/tex.js','vendor/package/js/output/svg.js','vendor/node_modules/sharp/package.json')
foreach ($name in $mustHave) { if (-not (Test-Path -LiteralPath (Join-Path $taskRoot $name))) { throw "Required runtime file missing: $name" } }
$files = Get-ChildItem -LiteralPath $taskRoot -File -Recurse -Force | Where-Object {
    $relative = $_.FullName.Substring($taskRoot.Length + 1)
    $top = ($relative -split '[\\/]')[0]
    # Exclude generated directories only at the application root. vendor/.../output is required.
    $top -notin @('.git','output','temp') -and $_.FullName -ne $OutputPath -and
        $_.Name -notmatch '\.(next|check)\.exe$' -and $_.Extension -notin @('.zip','.rar','.pdb')
}
$stream = [IO.File]::Open($OutputPath,[IO.FileMode]::Create,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
$archive = [IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($taskRoot.Length + 1).Replace('\','/')
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$file.FullName,($packageRoot + '/' + $relative),[IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
    foreach ($name in @('output','temp')) { $archive.CreateEntry($packageRoot + '/' + $name + '/') | Out-Null }
} finally { $archive.Dispose(); $stream.Dispose() }
$archive = [IO.Compression.ZipFile]::OpenRead($OutputPath)
try {
    foreach ($name in $mustHave) { if (-not $archive.GetEntry($packageRoot + '/' + $name)) { throw "Incomplete archive: $name" } }
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($taskRoot.Length + 1).Replace('\','/')
        $entry = $archive.GetEntry($packageRoot + '/' + $relative)
        if (-not $entry -or $entry.Length -ne $file.Length) { throw "Archive file mismatch: $relative" }
    }
} finally { $archive.Dispose() }
Write-Output ('Package verified: ' + $OutputPath)
