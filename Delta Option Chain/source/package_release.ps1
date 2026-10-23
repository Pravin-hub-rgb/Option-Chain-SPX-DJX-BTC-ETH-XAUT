$ErrorActionPreference = 'Stop'

$sourceDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDirectory = Join-Path $sourceDirectory 'dist'
# The folder the build actually populates. It used to point at
# dist\OptionChain, which nothing writes to: the build produces
# dist\delta_option_chain. So the release ZIP was a month-old build with
# none of the fixes in it, and the customer was sent that. A stale package
# is worse than no package, because it looks like the fix did not work.
$appDirectory = Join-Path $distDirectory 'delta_option_chain'
$archivePath = Join-Path $distDirectory 'DeltaOptionChain.zip'
$hashPath = "$archivePath.sha256"
$productExe = 'delta_option_chain.exe'

# license.json is intentionally absent. It is the owner's activation file; if it
# ever appears in the release folder the build fails rather than publishing a
# package that would run unlocked for anyone.
$requiredFiles = @(
    $productExe,
    'btc_chain.xlsx',
    'eth_chain.xlsx',
    'xaut_chain.xlsx',
    'config.json',
    'README.txt',
    'HOW-TO-USE.txt',
    'CHECK-INSTALL.bat',
    'diagnose_excel_addins.ps1',
    'optionchain.product'
)
foreach ($relativePath in $requiredFiles) {
    $path = Join-Path $appDirectory $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required release file is missing: $path"
    }
}

# A build older than the source is a stale build. Publishing one is how the
# customer ends up installing code that does not have the fix in it and
# reporting that the fix does not work. Fail instead.
$builtExecutable = Get-Item -LiteralPath (Join-Path $appDirectory $productExe)
$newestSource = Get-ChildItem -LiteralPath $sourceDirectory -File -Filter '*.py' |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -ne $newestSource -and $builtExecutable.LastWriteTime -lt $newestSource.LastWriteTime) {
    throw ("The built executable is older than the source ({0} vs {1}). " +
        "Run build_final.bat again before packaging.") -f `
        $builtExecutable.LastWriteTime, $newestSource.LastWriteTime
}

$productId = (Get-Content -LiteralPath (Join-Path $appDirectory 'optionchain.product') -Raw).Trim()
if ($productId -ne 'delta') {
    throw "Unexpected product identity marker in release: '$productId'"
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

# Our own Python source must not sit in the client payload. This is a
# ROOT-LEVEL check only: _internal legitimately carries openpyxl/xlwings
# package sources, because --collect-all ships them. Those are third-party
# library files, not our code, and scanning them recursively fails every build.
$straySource = Get-ChildItem -LiteralPath $appDirectory -File |
    Where-Object { $_.Extension -ieq '.py' -or $_.Extension -ieq '.pyc' }
if ($straySource) {
    $paths = ($straySource | ForEach-Object { $_.FullName }) -join "`n"
    throw "Unexpected Python source/bytecode in release root:`n$paths"
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
