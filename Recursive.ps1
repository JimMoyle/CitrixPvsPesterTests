
function Set-Text {
    <#
    .SYNOPSIS
    Walks an object with recursion

    .DESCRIPTION
    If the item in the objec has children the item is a heading and the depth into the object is recorded.  If the item is at the end of the branch and has no children, that item is a table entry.

    .PARAMETER Data
    The initial object you want the function to walk

    .PARAMETER Depth
    Mostly used to keep track of depth during recursion set to 1 by default.

    .EXAMPLE
    $PVSData | Set-Text

    This will walk the $PVSData object and output according to Heading and Table

    .NOTES
    General notes
    #>
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

        #$properties = $Data | get-member -type NoteProperty | Select-Object -ExpandProperty Name
        $properties = $Data.psobject.Properties | Select-Object -ExpandProperty Name

        foreach ($property in $properties) {
            try {

                switch ($true) {
                    ($null -eq $Data.$property) {
                        #Write-Output "Null Table Entry $property $Depth"
                        $output = [PSCustomObject]@{
                            Line = 'Table'
                            Property = $property.ToString()
                            Value = $null
                            Depth = $Depth
                        }
                        break
                    }
                    ($Data.$property.GetType().Name -ne 'PSCustomObject' -and $Data.$property.GetType().BaseType.ToString() -ne 'System.Array') {
                        #Write-Output "Match Table Entry $($property.ToString())"
                        $output = [PSCustomObject]@{
                            Line = 'Table'
                            Property = $property.ToString()
                            Value = $Data.$property.ToString()
                            Depth = $Depth
                        }
                        break
                    }
                    Default {
                        #Write-Output "Default Heading Entry $property $Depth"
                        $output = [PSCustomObject]@{
                            Line = 'Heading'
                            Property = $property.ToString()
                            Value = $null
                            Depth = $Depth
                        }
                        $Data.$property | Set-Text -Depth ($Depth + 1)
                    }
                }
            }
            catch {
                Write-Host 'bug'
            }
            Write-Output $output
        }
    }


    END {

    }

}

$PVSdata = Get-Content (Join-Path $PSScriptRoot pvs.json) | ConvertFrom-Json

$PVSData | Set-Text #| Add-Content c:\jimm\pvsresult.txt