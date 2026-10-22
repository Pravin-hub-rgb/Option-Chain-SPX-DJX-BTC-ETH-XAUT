$log = Join-Path $env:TEMP 'delta_run2.log'
$maxWait = 60
$wait = 0
while (-not (Test-Path $log) -or (Get-Content $log -Tail 1 -ReadCount 1 | Measure-Object -Line).Lines -eq 0) {
    Start-Sleep -Seconds 2
    $wait += 2
    if ($wait -gt $maxWait) { break }
}
if (Test-Path $log) {
    Write-Host '--- LOG (first 25 lines) ---'
    Get-Content $log -TotalCount 25
} else {
    Write-Host 'log file not found'
}
