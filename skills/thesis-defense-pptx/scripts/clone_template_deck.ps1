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
$dest = $null

try {
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

    Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force

    $pp = New-Object -ComObject PowerPoint.Application
    $dest = $pp.Presentations.Open($outputPath, $false, $false, $false)
    $slideCount = $dest.Slides.Count
    foreach ($idx in $indices) {
        if ($idx -lt 1 -or $idx -gt $slideCount) {
            throw "Slide index $idx is outside template slide range 1..$slideCount"
        }
    }

    $originalSlides = @()
    for ($i = 1; $i -le $slideCount; $i++) {
        $originalSlides += $dest.Slides.Item($i)
    }

    foreach ($idx in $indices) {
        $dupRange = $originalSlides[$idx - 1].Duplicate()
        $newSlide = $dupRange.Item(1)
        $newSlide.MoveTo($dest.Slides.Count)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($dupRange) | Out-Null
    }

    for ($i = $slideCount; $i -ge 1; $i--) {
        $dest.Slides.Item($i).Delete()
    }

    $dest.SaveAs($outputPath, 24)
    Write-Output $outputPath
}
finally {
    if ($dest -ne $null) {
        try { $dest.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($dest) | Out-Null
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}
