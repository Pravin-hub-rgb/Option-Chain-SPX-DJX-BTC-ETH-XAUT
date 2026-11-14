$log = $env:TEMP + '\b_delta_final.log'
$content = [System.IO.File]::ReadAllText($log)
$lines = $content -split '\r?\n'
$last = $lines[-1]
Write-Host 'LAST: ' + $last
if ($lines[-1] -match 'wrote .*OptionChain\.exe|completed successfully') {
    Write-Host 'BUILD DONE'
} else {
    Write-Host 'IN PROGRESS'
}
