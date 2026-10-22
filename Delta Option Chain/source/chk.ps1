$tf = $env:TEMP + '\b_delta_final.log'
if (Test-Path $tf) {
    $content = [System.IO.File]::ReadAllText($tf)
    $lines = $content -split '\r?\n'
    $len = $lines.Length
    $start = [Math]::Max(0, $len - 15)
    for ($i = $start; $i -lt $len; $i++) {
        Write-Host $lines[$i]
    }
} else {
    Write-Host 'no log'
}
