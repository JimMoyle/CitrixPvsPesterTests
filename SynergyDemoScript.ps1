#This is an as yet unofficial change to Carls scripts.
Start-Process http://carlwebster.com/category/citrix-pvs/

#New Object parameter

$params = @{
    Output       = 'Object' #I did think about using -PassThru, but settled on Object
    ComputerName = 'localhost'
    Verbose      = $true
}

#Run the script with verbose
$PvsObject = & '.\PvsDocumentationScriptWithObject.ps1' @params

#Show what is in the Object
$PvsObject | Format-List

#Show contents of the farm information Object
$PvsObject.PVSFarmInformation | Format-List

#Show the same information in the PVS console.

#Output the Object to a JSON file
$PvsObject | ConvertTo-Json | Set-Content C:\JimM\pvs.json

#Show contents of JSON File in Code

#Show Pester Script

#Use Pester to compare current config from JSON to actual config (Should all pass)
Invoke-Pester

#Change stuff in the PVS Console

#Rerun Pester (Some tests should fail)
Invoke-Pester

#This Output is written to the host.
#We all know that everytime you use Write-Host a puppy dies so what can we do?
$pesterObject = Invoke-Pester -Quiet -PassThru

#Show object contents
$pesterObject

#Show Failed Coumt
$pesterObject.FailedCount

#Do something with a failed test
If ($pesterObject.FailedCount -gt 0) {
    Write-Output 'Do Something Here'
}

