Function Set-WordHeadingLine {
    #Function created by Ryan Revord
    #@rsrevord on Twitter
    #Function created to make output to Word easy in this script
    #updated 27-Mar-2015 to include font name, font size, italics and bold options
    #Updated May 2017 by Jim Moyle to standarise Parameters with the other outputs and create advanced function whiwill take piipeline input.
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
        [String]$LineType = 'Heading',

        [Parameter(
            Position = 4,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [int]$Style = 0,

        [Parameter(
            Position = 5,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [string]$FontName = $Null,

        [Parameter(
            Position = 6,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [int]$fontSize = 0,

        [Parameter(
            Position = 7,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [bool]$Italics = $False,

        [Parameter(
            Position = 8,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [bool]$boldface = $False,

        [Parameter(
            Position = 9,
            ValuefromPipelineByPropertyName = $true,
            Mandatory = $false
        )]
        [Switch]$nonewline
    )

    BEGIN {}
    PROCESS {
        #Build output style
        [string]$output = ""

        #Max heading depth is 9 for word.
        if ($Depth -lt 10) {
            $Script:Selection.Style = $Script:MyHash.Word_Heading($Depth + 1)
        }
        else {
            $Script:Selection.Style = $Script:MyHash.Word_NoSpacing
        }

        <#Switch ($Depth) {

            0 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
            1 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
            2 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
            3 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
            4 {$Script:Selection.Style = $Script:MyHash.Word_Heading5; Break}
            5 {$Script:Selection.Style = $Script:MyHash.Word_Heading6; Break}
            6 {$Script:Selection.Style = $Script:MyHash.Word_Heading7; Break}
            7 {$Script:Selection.Style = $Script:MyHash.Word_Heading8; Break}
            8 {$Script:Selection.Style = $Script:MyHash.Word_Heading9; Break}
            Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
        }#>

        <#    #build # of tabs
            While ($tabs -gt 0) {
            $output += "`t"; $tabs--;
            }
        #>

        If (![String]::IsNullOrEmpty($fontName)) {
            $Script:Selection.Font.name = $fontName
        }

        If ($fontSize -ne 0) {
            $Script:Selection.Font.size = $fontSize
        }

        If ($italics -eq $True) {
            $Script:Selection.Font.Italic = $True
        }

        If ($boldface -eq $True) {
            $Script:Selection.Font.Bold = $True
        }

        #output the rest of the parameters.
        $output += $Property #+ $value
        $Script:Selection.TypeText($output)

        #test for new WriteWordLine 0.
        If ($nonewline) {
            # Do nothing.
        }
        Else {
            $Script:Selection.TypeParagraph()
        }
    }
    END{}
}