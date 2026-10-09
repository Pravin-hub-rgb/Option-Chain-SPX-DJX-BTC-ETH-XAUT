$ErrorActionPreference = 'Stop'

$sourceDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDirectory = Join-Path $sourceDirectory 'dist'
$appDirectory = Join-Path $distDirectory 'OptionChain'
$archivePath = Join-Path $distDirectory 'DeltaOptionChain.zip'
$hashPath = "$archivePath.sha256"

# license.json is intentionally absent. It is the owner's activation file; if it
# ever appears in the release folder the build fails rather than publishing a
# package that would run unlocked for anyone.
$requiredFiles = @(
    'OptionChain.exe',
    'btc_chain.xlsx',
    'eth_chain.xlsx',
    'xaut_chain.xlsx',
    'config.json',
    'README.txt',
    'diagnose_excel_addins.ps1'
)
foreach ($relativePath in $requiredFiles) {
    $path = Join-Path $appDirectory $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required release file is missing: $path"
    }
}

$runtimeDirectory = Join-Path $appDirectory '_internal\pywin32_system32'
foreach ($pattern in @('pythoncom*.dll', 'pywintypes*.dll')) {
    $runtimeDll = Get-ChildItem -LiteralPath $runtimeDirectory -Filter $pattern -File |
        Select-Object -First 1
    if ($null -eq $runtimeDll) {
        throw "Required pywin32 runtime DLL '$pattern' is missing from $runtimeDirectory"
    }
}

$unexpectedAddins = Get-ChildItem -LiteralPath $appDirectory -Recurse -File |
    Where-Object { $_.Extension -ieq '.xll' -or $_.Name -match 'NorenLink' }
if ($unexpectedAddins) {
    $paths = ($unexpectedAddins | ForEach-Object { $_.FullName }) -join "`n"
    throw "Unexpected third-party XLL/NorenLink files found in release:`n$paths"
}

# Hard block on the owner's activation file reaching customers.
$strayLicense = Join-Path $appDirectory 'license.json'
if (Test-Path -LiteralPath $strayLicense) {
    throw "license.json must never ship: it would let anyone run the tool unlocked. Remove it from $strayLicense"
}

# No Python source or licence sidecar files anywhere in the release payload.
$strayFiles = Get-ChildItem -LiteralPath $appDirectory -Recurse -File |
    Where-Object { $_.Extension -ieq '.py' -or $_.Extension -ieq '.pyc' }
if ($strayFiles) {
    $paths = ($strayFiles | ForEach-Object { $_.FullName }) -join "`n"
    throw "Unexpected Python source/bytecode files found in release:`n$paths"
}

# Only config.json is expected at the top level of the client payload.
$strayJson = Get-ChildItem -LiteralPath $appDirectory -File -Filter '*.json' |
    Where-Object { $_.Name -ne 'config.json' }
if ($strayJson) {
    $paths = ($strayJson | ForEach-Object { $_.FullName }) -join "`n"
    throw "Unexpected JSON files found in release root:`n$paths"
}

if (Test-Path -LiteralPath $archivePath) {
    Remove-Item -LiteralPath $archivePath
}
if (Test-Path -LiteralPath $hashPath) {
    Remove-Item -LiteralPath $hashPath
}
Compress-Archive -LiteralPath $appDirectory -DestinationPath $archivePath -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
Set-Content -LiteralPath $hashPath -Value "$hash  DeltaOptionChain.zip" -Encoding ASCII
Write-Output "Release ZIP: $archivePath"
Write-Output "SHA256 file: $hashPath"
