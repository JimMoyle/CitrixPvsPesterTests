function Set-TextLine {
    #function created by Michael B. Smith, Exchange MVP
    #@essentialexchange on Twitter
    #http://TheEssentialExchange.com
    #for creating the formatted text report
    #created March 2011
    #updated March 2015
    #updated May 2017 by Jim Moyle
    [CmdletBinding()]
    param (

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
        [String]$LineType = 'Heading'
    ) #Param

    BEGIN {}
    PROCESS {
        #Standard amount of characters from the start of the property in a table
        $firstTableColumeWidth = 30
        #If this is a heading straight after a table add an extra newline
        if ($script:previousLineType -eq 'Table' -and $lineType -eq 'Heading') {
            $script:TextOutput += "`r`n"
        }
        #Get the property the correct number of tabs in.
        While ( $Depth -gt 0 ) {
            $script:TextOutput += "`t"
            $Depth--
        }
        #Always put the property in the file
        $script:TextOutput += $Property

        #If the line is a table we need to put the value on the same line
        if ($LineType -eq 'Table') {
            #Standard amount of characters from the start of the property is $firstTableColumeWidth
            if ($Property.Length -lt $firstTableColumeWidth) {
                $spaces = $firstTableColumeWidth - $Property.Length
                While ( $spaces -gt 0 ) {
                    #Adding spaces rather than tabs due to differences in display in windows text readers
                    $script:TextOutput += ' '
                    $spaces--
                }
                $script:TextOutput += ": $Value"
            }
            else {
                #If the property contains more than 40 characters, just add a tab at the end.
                $script:TextOutput += "`t: $Value"
            }
        }
        #Add return at the end of the line
        $script:TextOutput += "`r`n"
        #Set previous line type so we can add a space after each table before a heading
        $script:previousLineType = $LineType
    }
    END {}
}