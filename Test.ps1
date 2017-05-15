$ErrorActionPreference = 'Continue'
$PSDefaultParameterValues = $null

$params = @{
    Output = 'MSWord'
    CompanyName = 'JimCorp'
    AddDateTime = $true
    ComputerName = 'localhost'
    CoverPage =  'Whisp'
    Folder = 'C:\JimM'
    UserName = 'Jim Moyle'
    Verbose = $true
    #to = "Jim <jim@atlantiscomputing.com>"
    #From = "James <jim@atlantiscomputing.com>"
    #UseSSL = $true
    #SmtpServer = 'smtp.office365.com'
}

& 'E:\JimM\Dropbox\Dropbox (Personal)\ScriptScratch\CarlDocScriptTemplate\ScriptTemplateJim.ps1' @params