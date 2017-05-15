
ps | where Name -eq winword | Stop-Process
Add-Type -AssemblyName Office
Add-Type -AssemblyName Microsoft.Office.Interop.Word
$p = 'Title'
$au = 'Abby'

#[array]$AryProperties = "Title"
#[array]$newValue = "Jim Title Text"
$application = New-Object -ComObject word.application
$application.Visible = $false
$binding = "System.Reflection.BindingFlags" -as [type]
$doc = Get-childitem -path "C:\JimM\jim_Template_Script_2017-02-16_0942.docx"
$document = $application.documents.open($doc.fullname)
$BuiltinProperties = $document.BuiltInDocumentProperties 
$pn = [System.__ComObject].invokemember("item",$binding::GetProperty,$null,$BuiltinProperties,$p)

$result = [System.__ComObject].invokemember("value",$binding::SetProperty,$null,$pn,$au)

$pn

Write-Output 'End'

#$builtinPropertiesType = $builtinProperties.GetType()
#$BuiltInProperty = $builtinPropertiesType.invokemember("item",$binding::GetProperty,$null,$BuiltinProperties,$AryProperties) 
#$BuiltInPropertyType = $BuiltInProperty.GetType()
#$BuiltInPropertyType.invokemember("value",$binding::SetProperty,$null,$BuiltInProperty,$newValue)
