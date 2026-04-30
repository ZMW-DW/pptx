param(
    [Parameter(Mandatory = $true)]
    [string]$Template,

    [Parameter(Mandatory = $true)]
    [string]$Output,

    [string]$SlideSequence = ""
)

$ErrorActionPreference = "Stop"

$templatePath = [System.IO.Path]::GetFullPath($Template)
$outputPath = [System.IO.Path]::GetFullPath($Output)

if (!(Test-Path -LiteralPath $templatePath)) {
    throw "Template not found: $templatePath"
}

$outDir = Split-Path -Parent $outputPath
if ($outDir -and !(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

$pp = $null
$source = $null
$dest = $null

try {
    $pp = New-Object -ComObject PowerPoint.Application

    if ([string]::IsNullOrWhiteSpace($SlideSequence)) {
        Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force
        Write-Output $outputPath
        return
    }

    $indices = @()
    foreach ($part in $SlideSequence.Split(",")) {
        $trimmed = $part.Trim()
        if ($trimmed.Length -gt 0) {
            $indices += [int]$trimmed
        }
    }
    if ($indices.Count -eq 0) {
        throw "SlideSequence did not contain any slide numbers."
    }

    $source = $pp.Presentations.Open($templatePath, $true, $false, $false)
    $slideCount = $source.Slides.Count
    foreach ($idx in $indices) {
        if ($idx -lt 1 -or $idx -gt $slideCount) {
            throw "Slide index $idx is outside template slide range 1..$slideCount"
        }
    }
    $source.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($source) | Out-Null
    $source = $null

    $dest = $pp.Presentations.Add($false)
    foreach ($idx in $indices) {
        [void]$dest.Slides.InsertFromFile($templatePath, $dest.Slides.Count, $idx, $idx)
    }

    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }
    $dest.SaveAs($outputPath, 24)
    Write-Output $outputPath
}
finally {
    if ($source -ne $null) {
        try { $source.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($source) | Out-Null
    }
    if ($dest -ne $null) {
        try { $dest.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($dest) | Out-Null
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}
