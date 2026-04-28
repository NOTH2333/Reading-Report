$ErrorActionPreference = "Stop"

$root = "g:\OBSIDIAN\Volume-Price Analysis"

Push-Location $root
try {
    python -X utf8 .\scripts\build_volume_price_package.py
    node .\scripts\build_chapter_ppts.js
    powershell -ExecutionPolicy Bypass -File .\scripts\export_ppt_previews.ps1
}
finally {
    Pop-Location
}
