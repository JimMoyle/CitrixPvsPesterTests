$script:TextOutput = ''

. .\Recursive.ps1

. .\Set-TextLine.ps1

$PVSdata = Get-Content (Join-Path $PSScriptRoot pvs.json) | ConvertFrom-Json

$PVSData | Convert-ObjToDoc | set-content c:\jimm\temp.txt #Set-TextLine

#$script:TextOutput | Out-File c:\jimm\FinalText.txt