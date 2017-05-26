function Set-Text {
    [CmdletBinding()]
    param (

        [Parameter(
            Position = 0,
            ValuefromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Mandatory = $true
        )]
        [System.Object]$Data,

        [Parameter(
            Position = 1,
            ValuefromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Mandatory = $false
        )]
        [int]$Depth = 1
    )

    BEGIN {
    }

    PROCESS {

        $properties = $Data | get-member -type NoteProperty | Select-Object -ExpandProperty Name

        foreach ($property in $properties) {

            if ($property -eq 'Sites') {
                Write-Output 'Site'
            }
            if ($Data.$property.GetType().Name -ne 'PSCustomObject' -and $Data.$property.GetType().BaseType.ToString() -ne 'System.Array') {
                Write-Host "Table Entry $property"
            }
            else {
                Write-Host "Heading Entry $property $Depth"
                $Data.$property | Set-Text -Depth ($Depth + 1)
            }
        }
    }


    END {

    }

}

$PVSdata = Get-Content "C:\Users\Jim\Dropbox (Personal)\ScriptScratch\VSCodeGit\CitrixPvsPesterTests\pvs.json" | ConvertFrom-Json

$PVSData | Set-Text