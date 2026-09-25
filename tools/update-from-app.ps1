# PEOPLE PIPELINE website - takes the latest page code from the Apps Script project folder.
# The website uses the SAME app/ folder and server/Pipeline.js as the Apps Script project; nothing else is copied.
#
#   powershell -ExecutionPolicy Bypass -File tools\update-from-app.ps1 -AppProject "C:\Users\me\people pipeline"
#
param([Parameter(Mandatory = $true)][string]$AppProject)
$Site = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path -LiteralPath (Join-Path $AppProject "app\index.html"))) { Write-Error "No app\index.html in $AppProject"; exit 1 }
Remove-Item -LiteralPath (Join-Path $Site "app") -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -LiteralPath (Join-Path $AppProject "app") -Destination (Join-Path $Site "app") -Recurse -Force
New-Item -ItemType Directory -Force -Path (Join-Path $Site "server") | Out-Null
Copy-Item -LiteralPath (Join-Path $AppProject "server\Pipeline.js") -Destination (Join-Path $Site "server\Pipeline.js") -Force
Write-Output "Website updated from $AppProject (app\ and server\Pipeline.js)"
