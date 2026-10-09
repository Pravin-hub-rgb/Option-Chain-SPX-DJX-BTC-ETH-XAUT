param(
    [string]$OutputPath = (Join-Path $env:USERPROFILE 'Desktop\delta_excel_addin_report.txt')
)

$ErrorActionPreference = 'Stop'
$report = [System.Collections.Generic.List[string]]::new()

function Add-ReportLine {
    param([string]$Line)
    $report.Add($Line)
}

function Add-XllFileReport {
    param([string]$Path)

    try {
        $file = Get-Item -LiteralPath $Path -ErrorAction Stop
        $stream = [System.IO.File]::OpenRead($file.FullName)
        try {
            $firstBytes = New-Object byte[] 2
            $bytesRead = $stream.Read($firstBytes, 0, 2)
        }
        finally {
            $stream.Dispose()
        }

        $startsWithMZ = $bytesRead -eq 2 -and
            $firstBytes[0] -eq 0x4D -and $firstBytes[1] -eq 0x5A
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        Add-ReportLine (
            "XLL file: {0} | size_bytes={1} | starts_with_MZ={2} | sha256={3}" -f
            $file.FullName, $file.Length, $startsWithMZ, $hash
        )
    }
    catch {
        Add-ReportLine ("XLL inspection failed: {0} | {1}" -f $Path, $_.Exception.Message)
    }
}

function Add-RegistryOpenEntries {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$HiveName
    )

    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )) {
        $baseKey = $null
        $key = $null
        try {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $view)
            foreach ($version in @('16.0', '15.0', '14.0')) {
                $keyPath = "Software\Microsoft\Office\$version\Excel\Options"
                $key = $baseKey.OpenSubKey($keyPath)
                if ($null -eq $key) {
                    continue
                }

                foreach ($name in $key.GetValueNames()) {
                    if ($name -match '^OPEN\d*$') {
                        $value = [Environment]::ExpandEnvironmentVariables(
                            [string]$key.GetValue($name)
                        )
                        Add-ReportLine (
                            "Registry OPEN: hive={0} | view={1} | version={2} | name={3} | value={4}" -f
                            $HiveName, $view, $version, $name, $value
                        )
                        foreach ($match in [regex]::Matches(
                            $value,
                            '(?i)([A-Z]:\\[^"]*?\.xll)'
                        )) {
                            Add-XllFileReport $match.Groups[1].Value.Trim()
                        }
                    }
                }
                $key.Dispose()
                $key = $null
            }
        }
        catch {
            Add-ReportLine (
                "Registry read error: hive={0} | view={1} | {2}" -f
                $HiveName, $view, $_.Exception.Message
            )
        }
        finally {
            if ($null -ne $key) { $key.Dispose() }
            if ($null -ne $baseKey) { $baseKey.Dispose() }
        }
    }
}

$now = Get-Date
Add-ReportLine "Delta Excel add-in diagnostic (read-only)"
Add-ReportLine ("Created: {0}" -f $now.ToString('o'))
Add-ReportLine ("Windows: {0} | version={1} | build={2}" -f
    [Environment]::OSVersion.VersionString,
    [Environment]::OSVersion.Version,
    (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuild
)
Add-ReportLine "This report reads Excel add-in registrations and common startup folders only. It does not alter add-ins or registry values; only the report file is written."
Add-ReportLine ""
Add-RegistryOpenEntries ([Microsoft.Win32.RegistryHive]::CurrentUser) 'HKCU'
Add-RegistryOpenEntries ([Microsoft.Win32.RegistryHive]::LocalMachine) 'HKLM'

$startupFolders = @(
    (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'),
    (Join-Path $env:APPDATA 'Microsoft\AddIns'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Excel\XLSTART'),
    (Join-Path $env:PROGRAMDATA 'Microsoft\Excel\XLSTART')
)
foreach ($programFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if ($programFiles) {
        $startupFolders += (Join-Path $programFiles 'Microsoft Office\root\Office16\XLSTART')
        $startupFolders += (Join-Path $programFiles 'Microsoft Office\Office16\XLSTART')
        $startupFolders += (Join-Path $programFiles 'Microsoft Office\root\Office16\Library')
    }
}

$seenPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
foreach ($folder in $startupFolders | Select-Object -Unique) {
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        continue
    }

    Add-ReportLine ("Startup folder: {0}" -f $folder)
    try {
        foreach ($file in Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction Stop) {
            Add-ReportLine ("  file: {0} | size_bytes={1}" -f $file.FullName, $file.Length)
            if ($file.Extension -ieq '.xll' -and $seenPaths.Add($file.FullName)) {
                Add-XllFileReport $file.FullName
            }
        }
    }
    catch {
        Add-ReportLine ("  folder read error: {0}" -f $_.Exception.Message)
    }
}

$outputDirectory = Split-Path -Parent $OutputPath
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
$report | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output ("Read-only report saved to: {0}" -f $OutputPath)
Write-Output "No add-ins or registry entries were changed; only the report file was written."
