function Convert-ObjToDoc {
    <#
    .SYNOPSIS
    Walks an object with recursion.

    .DESCRIPTION
    If the item in the object has children the item is a heading and the depth into the object is recorded.  If the item is at the end of the branch and has no children, that item is a table entry.

    .PARAMETER Data
    The initial object you want the function to walk

    .PARAMETER Depth
    Mostly used to keep track of depth during recursion set to 1 by default.

    .EXAMPLE
    $PVSData | Set-Text

    This will walk the $PVSData object and output according to Heading and Table

    LineType    Property                      Value                                    Depth
    --------    --------                      -----                                    -----
    Table       Version                       7.13.0.13008                                 2
    Table       FarmName                      MyFarm                                       3
    Table       Description                                                                3
    Heading     General                                                                    2
    Table       AuthGroupName                 JimMoyle.local/Builtin/Administrators        3
    Table       AuthGroupName                 JimMoyle.local/Users/Domain Admins           3
    Heading     Security                                                                   2
    Table       AuthGroupName                 JimMoyle.local/Builtin/Administrators        3
    Table       AuthGroupName                 JimMoyle.local/Users/Domain Admins           3
    Table       AuthGroupName                 JimMoyle.local/Users/PVSLondonSiteAdmins     3

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
        [int]$Depth = 0
    )

    BEGIN {
    }

    PROCESS {

        #Get object properties so that we can pass them all through the recursion
        $properties = $Data.psobject.Properties | Select-Object -ExpandProperty Name

        foreach ($property in $properties) {
            try {
                #If the data is Null or a Custom Object or an Array, then make it a heading else default to table entry.
                #Both custom objects and Arrays should be created by the script.
                switch ($true) {
                    ($null -eq $Data.$property) {
                        #Write-Output "Null Table Entry $property $Depth"
                        $output = [PSCustomObject]@{
                            LineType = 'Table'
                            Property = $property.ToString()
                            Value    = $null
                            Depth    = $Depth
                        }
                        break
                    }
                    ($Data.$property.GetType().Name -ne 'PSCustomObject' -and $Data.$property.GetType().BaseType.ToString() -ne 'System.Array') {
                        #Write-Output "Match Table Entry $($property.ToString())"
                        $output = [PSCustomObject]@{
                            LineType = 'Table'
                            Property = $property.ToString()
                            Value    = $Data.$property.ToString()
                            Depth    = $Depth
                        }
                        break
                    }
                    Default {
                        #Write-Output "Default Heading Entry $property $Depth"
                        $output = [PSCustomObject]@{
                            LineType = 'Heading'
                            Property = $property.ToString()
                            Value    = $null
                            Depth    = $Depth
                        }
                        $Data.$property | Set-Text -Depth ($Depth + 1)
                    }
                }
            }
            catch {
                Write-Host 'bug'
            }
            #Output is from loop rather than gathering up so that if data is pipelined out of the function it can be worked on quicker
            Write-Output $output
        }
    }
    END {
    }
}