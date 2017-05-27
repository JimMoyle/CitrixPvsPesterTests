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
            try {

                switch ($true) {
                    ($null -eq $Data.$property) {
                        Write-Host "Null Table Entry $property"
                        break
                     }
                    ($Data.$property.GetType().Name -ne 'PSCustomObject' -and $Data.$property.GetType().BaseType.ToString() -ne 'System.Array') {
                        Write-Host "Match Table Entry $($property.ToString())"
                        break
                    }
                    Default {
                        Write-Host "Default Heading Entry $property $Depth"
                        $Data.$property | Set-Text -Depth ($Depth + 1)
                    }
                }
            }
            catch {
                Write-Host 'bug'
            }

        }
    }


    END {

    }

}

$PVSdata = Get-Content "C:\Users\Jim\Dropbox (Personal)\ScriptScratch\VSCodeGit\CitrixPvsPesterTests\pvs.json" | ConvertFrom-Json

$PVSData | Set-Text