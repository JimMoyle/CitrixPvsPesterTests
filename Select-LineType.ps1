Function Select-WordLineType {

    [CmdletBinding()]
    Param(
        [Parameter(
            Position = 0,
            ValuefromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Mandatory = $true
        )]
        [String]$Property,

        [Parameter(
            Position = 1,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [String]$Value,

        [Parameter(
            Position = 2,
            ValuefromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Mandatory = $false
        )]
        [Int]$Depth = 1,

        [Parameter(
            Position = 3,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [String]$LineType
    )

    BEGIN {}
    PROCESS {
        #Switch block to correctly pass headings and tables to the correct word functions
        switch ($true) {
            ($LineType -eq 'Heading' -and $script:previousLineType -ne 'Table') {
                #Heading where you don't need to do anything with a previous table
                Set-WordHeadingLine -Property $Property -Depth $Depth -LineType $LineType
                break
            }
            ($LineType -eq 'Table') {
                #table Entry, save values into a hash table to feed to the create table function.
                $Script:tableLines += @{$Property = $Value}
                break
            }
            ($LineType -eq 'Heading' -and $script:previousLineType -eq 'Table') {
                #Write table out
                Write-Output $Script:tableLines # TODO add table function
                #Write heading ater table
                Set-WordHeadingLine -Property $Property -Depth $Depth -LineType $LineType
                break
            }
            Default {
                Write-Error "Could not determine $Property LineType. Current line type is $LineType, Previous line type was $script:previousLineType"
            }
        } #switch

        #Put Current line type into variable to be read on next loop
        $script:previousLineType = $LineType
    }
    END {}
}

$data = [pscustomobject]@{LineType = 'Heading'; Property = 'General'; Value = $null; Depth = 2}

$data | Set-WordHeadingLine