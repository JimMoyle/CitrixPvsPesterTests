# -----------------------------------------------------------------------------
# Script: Set-SpecificDocumentProperties.ps1
# Author: ed wilson, msft
# Date: 08/02/2012 17:45:27
# Keywords: Office, Microsoft Word
# comments: 
# HSG-8-3-2012
# -----------------------------------------------------------------------------
Param(
  $path = "C:\jimm", [array]$include = @("Jim*.doc*","WES*.doc*"))
[array]$AryProperties = "Comments"
[array]$newValue = "Scripting Guy blog"
ps | where Name -eq winword | Stop-Process
$application = New-Object -ComObject word.application
$application.Visible = $false
$binding = "System.Reflection.BindingFlags" -as [type]
$docs = Get-childitem -path $Path -Recurse -Include $include 
 
Foreach($doc in $docs)
 {
  $document = $application.documents.open($doc.fullname)
  $BuiltinProperties = $document.BuiltInDocumentProperties 
  $builtinPropertiesType = $builtinProperties.GetType() 
    Try 
     { 
      $BuiltInProperty = $builtinPropertiesType.invokemember("item",$binding::GetProperty,$null,$BuiltinProperties,$AryProperties) 
      $BuiltInPropertyType = $BuiltInProperty.GetType()
      $BuiltInPropertyType.invokemember("value",$binding::SetProperty,$null,$BuiltInProperty,$newValue)}
    Catch [system.exception]
      { write-host -foreground blue "Unable to set value for $AryProperties" } 
   $document.close() 
   [System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
   [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
   Remove-Variable -Name document, BuiltinProperties
   }

$application.quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
Remove-Variable -Name application
[gc]::collect()
[gc]::WaitForPendingFinalizers()