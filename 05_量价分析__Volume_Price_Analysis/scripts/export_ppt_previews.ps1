param(
    [string]$PptRoot = "g:\OBSIDIAN\Volume-Price Analysis\《量价分析》教学化交付包\01_PPT"
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ComObject)
    }
}

$pptFiles = Get-ChildItem -LiteralPath $PptRoot -Recurse -Filter *.pptx | Sort-Object FullName
$results = @()
$powerPoint = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1

    foreach ($pptFile in $pptFiles) {
        $presentation = $null
        $chapterDir = $pptFile.Directory.FullName
        $previewDir = Join-Path $chapterDir "预览"
        $pngDir = Join-Path $previewDir "PNG"
        New-Item -ItemType Directory -Force -Path $previewDir, $pngDir | Out-Null

        try {
            $presentation = $powerPoint.Presentations.Open($pptFile.FullName, $false, $true, $false)
            $pdfPath = Join-Path $previewDir ($pptFile.BaseName + "_预览.pdf")
            $presentation.SaveAs($pdfPath, 32)
            $presentation.SaveAs($pngDir, 18)

            $results += [PSCustomObject]@{
                file = $pptFile.FullName
                slide_count = $presentation.Slides.Count
                pdf = $pdfPath
                png_dir = $pngDir
                status = "ok"
            }
        }
        catch {
            $results += [PSCustomObject]@{
                file = $pptFile.FullName
                slide_count = $null
                pdf = $null
                png_dir = $null
                status = "error"
                message = $_.Exception.Message
            }
        }
        finally {
            if ($null -ne $presentation) {
                $presentation.Close()
                Release-ComObject $presentation
            }
        }
    }
}
finally {
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
        Release-ComObject $powerPoint
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$resultPath = Join-Path $PptRoot "PPT校验结果.json"
$results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $resultPath -Encoding UTF8
Write-Output "PowerPoint 预览与校验完成：$resultPath"
