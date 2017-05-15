# -----------------------------------------------------------------------------
# Script: Get-SpecificDocumentProperties.ps1
# Author: ed wilson, msft
# Date: 08/01/2012 17:45:27
# Keywords: Microsoft Office, Word
# comments: 
# HSG-8-2-2012
# -----------------------------------------------------------------------------
Param(
  $path = "C:\jimm",
  [array]$include = @("jim*.docx","WES*.docx"))
  ps | where Name -eq winword | Stop-Process
$AryProperties = "Title","Author","Keywords", "Number of words", "Number of pages"
$application = New-Object -ComObject word.application
$application.Visible = $false
$binding = "System.Reflection.BindingFlags" -as [type]
[ref]$SaveOption = "microsoft.office.interop.word.WdSaveOptions" -as [type]
$docs = Get-childitem -path $Path -Recurse -Include $include 
Foreach($doc in $docs)
 {
  $document = $application.documents.open($doc.fullname)
  $BuiltinProperties = $document.BuiltInDocumentProperties
  $objHash = @{"Path"=$doc.FullName}
   foreach($p in $AryProperties)
    {Try 
     { 
      $pn = [System.__ComObject].invokemember("item",$binding::GetProperty,$null,$BuiltinProperties,$p) 
      $value = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$pn,$null)
      $objHash.Add($p,$value) }
     Catch [system.exception]
      { write-host -foreground blue "Value not found for $p" } }
   $docProperties = New-Object psobject -Property $objHash
   $docProperties 
   $document.close([ref]$saveOption::wdDoNotSaveChanges) 
   [System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
   [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
   Remove-Variable -Name document, BuiltinProperties
   }

$application.quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
Remove-Variable -Name application
[gc]::collect()
[gc]::WaitForPendingFinalizers()