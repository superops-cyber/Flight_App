$source = "PilotApp.html"
$outputFile = "rendered_pilot.html"
$mappingFile = "line_map.txt"

$renderedLines = New-Object System.Collections.Generic.List[string]
$lineMappings = New-Object System.Collections.Generic.List[string]

function Process-File($fileName) {
    $path = $fileName
    if (-not (Test-Path $path)) {
        if ($path -notlike "*.html") { $path = "$path.html" }
    }
    
    if (-not (Test-Path $path)) {
        Write-Warning "File not found: $path"
        return
    }

    $lines = Get-Content $path
    $lineNum = 1
    foreach ($line in $lines) {
        if ($line -match "<\?!= include\('(.+?)'\);? \?>") {
            $includeName = $Matches[1]
            Process-File $includeName
        }
        else {
            $renderedLines.Add($line)
            $lineMappings.Add("$($renderedLines.Count): $path`:$lineNum")
        }
        $lineNum++
    }
}

Process-File $source
$renderedLines | Out-File $outputFile -Encoding utf8
$lineMappings | Out-File $mappingFile -Encoding utf8

$m3841 = $lineMappings | Where-Object { $_ -match "^3841:" }
Write-Output "RESULT_START"
Write-Output "MAPPING_3841: $m3841"
for ($i = 3835; $i -lt 3846; $i++) {
    if ($i -ge 0 -and $i -lt $renderedLines.Count) {
        Write-Output "LINE_$($i+1): $($lineMappings[$i]) | $($renderedLines[$i])"
    }
}
Write-Output "RESULT_END"
