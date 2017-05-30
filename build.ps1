$script:TextOutput = ''

. .\Recursive.ps1

. .\Set-TextLine.ps1

$PVSdata = Get-Content (Join-Path $PSScriptRoot pvs.json) | ConvertFrom-Json

$PVSData | Convert-ObjToDoc | Set-TextLine

$script:TextOutput | Out-File c:\jimm\FinalText.txt