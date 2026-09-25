# PEOPLE PIPELINE website - copies the newest "people pipeline.xlsx" into the website's data folder.
# Run it every hour with Windows Task Scheduler (see README), so the website always shows fresh data.
#
#   powershell -ExecutionPolicy Bypass -File "C:\inetpub\people-pipeline\tools\copy-excel.ps1"
#
# CHANGE THESE TWO LINES:
$Source = "\\server\share\people pipeline.xlsx"          # where the query export saves the Excel file
$Target = "C:\inetpub\people-pipeline\data\people pipeline.xlsx"

if (-not (Test-Path -LiteralPath $Source)) { Write-Error "Excel file not found: $Source"; exit 1 }
$tmp = "$Target.tmp"
Copy-Item -LiteralPath $Source -Destination $tmp -Force          # copy first, then swap: visitors never get half a file
Move-Item -LiteralPath $tmp -Destination $Target -Force
(Get-Item -LiteralPath $Target).LastWriteTime = (Get-Item -LiteralPath $Source).LastWriteTime
Write-Output "Copied $Source -> $Target"
