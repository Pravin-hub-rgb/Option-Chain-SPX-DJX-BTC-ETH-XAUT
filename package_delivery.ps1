# Build the two delivery ZIPs from the built packages.
#
#     powershell -NoProfile -ExecutionPolicy Bypass -File package_delivery.ps1
#
# The archives hold the files at the ROOT, with no wrapper folder, so the
# customer can extract straight into a short folder:
#
#     C:\OptionChain\us_option_chain.exe
#     C:\OptionChain\delta_option_chain.exe
#
# That matters. Excel stops opening a workbook somewhere past 128
# characters and never returns, so every folder the customer creates on the
# way to the tool costs them a broken install. No wrapper folder means no
# extra 20 characters they did not ask for.
#
# The script refuses to publish anything that is missing a file the product
# needs, or that carries the owner's activation file or the Databento key.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$delivery = Join-Path $root 'DELIVERY'
if (-not (Test-Path -LiteralPath $delivery)) {
    New-Item -ItemType Directory -Force -Path $delivery | Out-Null
}

$packages = @(
    @{
        Name = 'us-option-chain.zip'
        Source = Join-Path $root 'US Option Chain\client package'
        Required = @(
            'us_option_chain.exe',
            'spx_option_chain.xlsx',
            'djx_option_chain.xlsx',
            'us_stock_option_chain.xlsx',
            'config.json',
            'HOW-TO-USE.txt',
            'CHECK-INSTALL.bat',
            'optionchain.product',
            'build_info.txt'
        )
        Forbidden = @('license.json', 'databento.key', '.env', '.env.example',
                      'instance.lock', 'instance_guard.log', 'doctor-report.txt')
        Product = 'us'
    },
    @{
        Name = 'delta-option-chain.zip'
        Source = Join-Path $root 'Delta Option Chain\source\dist\delta_option_chain'
        Required = @(
            'delta_option_chain.exe',
            'btc_chain.xlsx',
            'eth_chain.xlsx',
            'xaut_chain.xlsx',
            'config.json',
            'HOW-TO-USE.txt',
            'CHECK-INSTALL.bat',
            'optionchain.product',
            'build_info.txt'
        )
        Forbidden = @('license.json', 'databento.key', '.env', '.env.example',
                      'instance.lock', 'instance_guard.log', 'doctor-report.txt',
                      'OptionChain.exe', 'OptionChain.pyd')
        Product = 'delta'
    }
)

foreach ($package in $packages) {
    $source = $package.Source
    $name = $package.Name

    if (-not (Test-Path -LiteralPath $source -PathType Container)) {
        throw "Nothing to package: $source is missing. Build the product first."
    }

    foreach ($relative in $package.Required) {
        $path = Join-Path $source $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Required file is missing from the package: $path"
        }
    }

    # The owner's activation file and the Databento key must never ship.
    foreach ($forbidden in $package.Forbidden) {
        $path = Join-Path $source $forbidden
        if (Test-Path -LiteralPath $path) {
            throw "Refusing to publish: $path must never be in a client package."
        }
    }

    # The product marker tells the instance guard which product a folder
    # holds, and tells the customer which tool they are running.
    $marker = (Get-Content -LiteralPath (Join-Path $source 'optionchain.product') -Raw).Trim()
    if ($marker -ne $package.Product) {
        throw "Product marker is '$marker' but this is the '$($package.Product)' package."
    }

    $archive = Join-Path $delivery $name
    if (Test-Path -LiteralPath $archive) {
        Remove-Item -LiteralPath $archive -Force
    }
    # No .sha256 sidecar. The customer is sent the ZIP and nothing else, and
    # the build id is stamped inside the package (build_info.txt), so support
    # can still tell which build a customer is running from startup_diag.log.

    # The files go in at the root: no wrapper folder.
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $source, $archive, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    $sizeMb = [math]::Round((Get-Item -LiteralPath $archive).Length / 1MB, 2)
    Write-Output ("{0}  {1} MB" -f $name, $sizeMb)
}

Write-Output ''
Write-Output "Delivery archives are in $delivery"
