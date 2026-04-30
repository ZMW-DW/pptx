param(
    [Parameter(Mandatory = $true)]
    [string]$Pptx,

    [double]$Tolerance = 40,

    [string]$OutputJson = ""
)

$ErrorActionPreference = "Stop"

$pptxPath = [System.IO.Path]::GetFullPath($Pptx)
if (!(Test-Path -LiteralPath $pptxPath)) {
    throw "PPTX not found: $pptxPath"
}

$pp = $null
$pres = $null
$results = @()

try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($pptxPath, $true, $false, $false)

    foreach ($slide in $pres.Slides) {
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $txt = $shape.TextFrame.TextRange.Text.Trim()
                    if ($txt.Length -gt 0) {
                        $bw = $shape.TextFrame2.TextRange.BoundWidth
                        $bh = $shape.TextFrame2.TextRange.BoundHeight
                        $isOverflow = ($bw -gt ($shape.Width + $Tolerance)) -or ($bh -gt ($shape.Height + $Tolerance))
                        if ($isOverflow) {
                            $short = $txt
                            if ($short.Length -gt 180) {
                                $short = $short.Substring(0, 180)
                            }
                            $short = $short -replace "`r|`n", " / "
                            $results += [PSCustomObject]@{
                                slide = $slide.SlideIndex
                                shape = $shape.Name
                                width = [math]::Round($shape.Width, 1)
                                bound_width = [math]::Round($bw, 1)
                                height = [math]::Round($shape.Height, 1)
                                bound_height = [math]::Round($bh, 1)
                                text = $short
                            }
                        }
                    }
                }
            }
            catch {
                # Some shapes expose partial text APIs; skip them.
            }
        }
    }
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}

$payload = [PSCustomObject]@{
    pptx = $pptxPath
    tolerance = $Tolerance
    overflow_count = $results.Count
    overflows = $results
}

$json = $payload | ConvertTo-Json -Depth 6
if (![string]::IsNullOrWhiteSpace($OutputJson)) {
    $jsonPath = [System.IO.Path]::GetFullPath($OutputJson)
    $jsonDir = Split-Path -Parent $jsonPath
    if ($jsonDir -and !(Test-Path -LiteralPath $jsonDir)) {
        New-Item -ItemType Directory -Force -Path $jsonDir | Out-Null
    }
    Set-Content -LiteralPath $jsonPath -Value $json -Encoding UTF8
}
$json
