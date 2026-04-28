param(
    [string]$InputDir = "output\\ppt",
    [string]$OutputDir = "tmp\\rendered"
)

$ErrorActionPreference = "Stop"

$workspace = Resolve-Path (Join-Path $PSScriptRoot "..\\..")
$inputPath = Join-Path $workspace $InputDir
$renderPath = Join-Path $workspace $OutputDir
New-Item -ItemType Directory -Force -Path $renderPath | Out-Null

$powerpoint = $null
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application

    Get-ChildItem -LiteralPath $inputPath -Filter *.pptx | Sort-Object Name | ForEach-Object {
        $deck = $_
        $deckOut = Join-Path $renderPath $deck.BaseName
        if (Test-Path -LiteralPath $deckOut) {
            Remove-Item -LiteralPath $deckOut -Recurse -Force
        }
        New-Item -ItemType Directory -Force -Path $deckOut | Out-Null

        $presentation = $powerpoint.Presentations.Open($deck.FullName, $true, $true, $false)
        try {
            $presentation.SaveAs($deckOut, 18)
        }
        finally {
            $presentation.Close()
        }
        Write-Output "Rendered $($deck.Name) -> $deckOut"
    }
}
finally {
    if ($powerpoint) {
        $powerpoint.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
