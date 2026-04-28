$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$pptDir = Join-Path $root "deliverables\ppt"
$checkDir = Join-Path $root "_work\checks\renders"

New-Item -ItemType Directory -Force -Path $checkDir | Out-Null

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = 1

try {
    $results = @()
    Get-ChildItem -LiteralPath $pptDir -Filter "*.pptx" | ForEach-Object {
        $deckPath = $_.FullName
        $deckName = $_.BaseName
        $deckOut = Join-Path $checkDir $deckName
        New-Item -ItemType Directory -Force -Path $deckOut | Out-Null

        $presentation = $powerPoint.Presentations.Open($deckPath, $false, $true, $false)
        $slideCount = $presentation.Slides.Count

        $middle = [Math]::Ceiling($slideCount / 2)
        $targets = @()
        $targets += 1
        if ($middle -ne 1) { $targets += $middle }
        if ($slideCount -ne $middle) { $targets += $slideCount }

        foreach ($slideIndex in ($targets | Select-Object -Unique)) {
            $slideNumber = [int]$slideIndex
            $slide = $presentation.Slides.Item($slideNumber)
            $pngPath = Join-Path $deckOut ("slide-{0}.png" -f $slideNumber)
            $slide.Export($pngPath, "PNG", 1600, 900)
        }

        $results += [PSCustomObject]@{
            Deck = $deckName
            SlideCount = $slideCount
            ExportedSlides = ($targets | Select-Object -Unique) -join ", "
        }

        $presentation.Close()
    }

    $results | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $root "_work\checks\validation-summary.json") -Encoding UTF8
    $results | Format-Table -AutoSize
}
finally {
    $powerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
}
