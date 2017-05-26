function Set-Text {
    [CmdletBinding()]
    param (

        [Parameter(
            Position = 0,
            ValuefromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Mandatory = $true
        )]
        [System.Object]$Data
    )

    BEGIN {
    }

    PROCESS {
        if ($Data.GetType().Name -ne 'PSCustomObject') {
            write-Host "Table Entry $($Data.ToString())"
        }
        else {
            $Data | ForEach-Object {
                Write-Host "Heading Entry $($_.ToString())"
                $property | Set-Text
            }
        }
    }


    END {

    }

}

        $PVSdata = Get-Content "C:\Users\Jim\Dropbox (Personal)\ScriptScratch\VSCodeGit\CitrixPvsPesterTests\pvs.json" | ConvertFrom-Json

        $PVSData | Set-Text