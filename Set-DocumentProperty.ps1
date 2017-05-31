function Set-DocumentProperty  {
	param (
		[object]$Document,
		[String]$DocProperty,
		[string]$Value
	)
	try{
		$binding = "System.Reflection.BindingFlags" -as [type]
		$builtInProperties = $Document.BuiltInDocumentProperties
		$property = [System.__ComObject].invokemember("item",$binding::GetProperty,$null,$BuiltinProperties,$DocProperty)
		[System.__ComObject].invokemember("value",$binding::SetProperty,$null,$property,$Value)
	}
	catch [system.exception]{
		Write-Warning "Failed to set $DocProperty to $Value"
	}
}