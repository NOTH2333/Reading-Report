$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$pptDir = Join-Path $root "03_PPT"
$outDir = Join-Path $root "05_中间素材\导出预览"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$pptFiles = Get-ChildItem -LiteralPath $pptDir -Filter *.pptx | Sort-Object Name
if (-not $pptFiles) {
    throw "未在 $pptDir 找到 PPTX 文件。"
}

$powerPoint = New-Object -ComObject PowerPoint.Application
try {
    foreach ($file in $pptFiles) {
        $pdfPath = Join-Path $outDir ($file.BaseName + ".pdf")
        Write-Host "Exporting $($file.Name) -> $pdfPath"
        $presentation = $powerPoint.Presentations.Open($file.FullName, $false, $false, $false)
        try {
            $presentation.SaveAs($pdfPath, 32)
        }
        finally {
            $presentation.Close()
        }
    }
}
finally {
    $powerPoint.Quit()
}
