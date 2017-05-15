function Param-Test
{
    [CmdletBinding(DefaultParameterSetName = 'None') ]

    Param(
        [ValidateSet('MSWord', 'PDF', 'Text', 'HTML')]
        [String]$Output='MSWord',

        [Switch]$AddDateTime,
        
        [Switch]$Hardware,

        [string]$ComputerName='LocalHost',
        
        [string]$Folder,
        
        [Alias("CN")]
        [string]$CompanyName,
        
        [Alias("CP")]
        [string]$CoverPage='Sideline', 

        [Alias("UN")]
        [string]$UserName=$env:username,

        [parameter(ParameterSetName="SMTP",Mandatory=$True)]
        [string]$SmtpServer,

        [parameter(ParameterSetName="SMTP",Mandatory=$False)]
        [int]$SmtpPort=25,

        [parameter(ParameterSetName="SMTP",Mandatory=$False)]
        [switch]$UseSSL,

        [parameter(ParameterSetName="SMTP",Mandatory=$True)]
        [string]$From,

        [parameter(ParameterSetName="SMTP",Mandatory=$True)]
        [string]$To,

        [Switch]$Dev,
        
        [Alias("SI")]
        [Switch]$ScriptInfo
	
	)

    switch ($Output){
        MSword {$MSword = $true; break}
        PDF {$PDF = $true; break}
        HTML {$HTML = $true
            if ($CompanyName -ne ""){
                Write-Warning 'The CompanyName parameter is not used when specifying HTML as the Output'
            }
            if ($CoverPage -ne "Sideline"){
                Write-Warning 'The CoverPage parameter is not used when specifying HTML as the Output'
            }
            if ($UserName -ne $env:username){
                Write-Warning 'The UserName parameter is not used when specifying HTML as the Output'
            }
            break}
        Text {$Text = $true
            if ($CompanyName -ne ""){
                Write-Warning 'The CompanyName parameter is not used when specifying Text as the Output'
            }
            if ($CoverPage -ne "Sideline"){
                Write-Warning 'The CoverPage parameter is not used when specifying Text as the Output'
            }
            if ($UserName -ne $env:username){
                Write-Warning 'The UserName parameter is not used when specifying Text as the Output'
            }
        break}
    }
    Write-Output $PSBoundParameters
}