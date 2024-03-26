$shares = gwmi Win32_Share | Select-Object Name, Path, Description | Sort-Object Name
$shares_dict = @{}
foreach ($shr in $shares) {
	$shares_dict[$shr.Name] = @($shr.Path, $shr.Description)	
}
return $shares_dict | ConvertTo-Json


