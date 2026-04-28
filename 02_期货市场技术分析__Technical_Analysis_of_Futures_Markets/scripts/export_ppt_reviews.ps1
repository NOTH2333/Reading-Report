$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$pptRoot = Join-Path $root "02_章节PPT"

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1

try {
    Get-ChildItem $pptRoot -Directory | ForEach-Object {
        $pptx = Get-ChildItem $_.FullName -Filter *.pptx | Select-Object -First 1
        if (-not $pptx) { return }
        $reviewDir = Join-Path $_.FullName "review"
        if (Test-Path $reviewDir) {
            Get-ChildItem $reviewDir -File | Remove-Item -Force
        }
        else {
            New-Item -ItemType Directory -Path $reviewDir | Out-Null
        }
        $pres = $app.Presentations.Open($pptx.FullName, $false, $true, $false)
        $pres.Export($reviewDir, "PNG", 1280, 720)
        $pres.Close()
        Write-Output "exported $($pptx.Name)"
    }
}
finally {
    $app.Quit()
}
