$pass_politics = net accounts
$pass_dict = @{} # A dictionary to store setting values
$i = 0
foreach ($row in $pass_politics[0 .. 7]) # Iterating through the lines of obtained settings
{
	if ($row -match ':') # If the line is correct
	{
		$pass_dict[[string]$i] = $row.split(':')[1].Trim() # Retrieving the value of a specific setting
	}
	else
	{
		$pass_dict[[string]$i] = '' # Adding an empty value
	}
	$i += 1 # Incrementing the parameter number
}
return $pass_dict | ConvertTo-Json