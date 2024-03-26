$autorun_software = @{} # Dictionary for storing information about software in autostart
$OSArchitecture = (gwmi win32_operatingSystem).OSArchitecture # Checking architecture (64 or 86)
$list1 = gcim win32_startupcommand | Select-Object Caption, Command # List of software from the Autostart module in the OS
foreach ($app in $list1) # Iterating through all the software
{
	if ($app -notin $autorun_software) # If the software is not in the dictionary
	{
		$autorun_software[$app.Caption] = $app.Command # Add name and path
	}
	else
	{
		# If the software is missing in the dictionary
		$autorun_software[$app.Caption + '_'] = $app.Command # Add new name and path 
	}
}
$registry_paths = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"

if ($OSArchitecture -like '*64*')
{
	# If the OS is 64-bit
	$registry_paths += "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce"
}

foreach ($reg in $registry_paths)
{
	$apps = Get-Item -Path $reg | Where-Object ValueCount -ne 0 # Autostart data from the registry
	foreach ($app in $apps.Property)
	{
		if ($app -notin $autorun_software) # If the software is not in the dictionary
		{
			$autorun_software[$app] = (Get-ItemProperty -Path $reg).$app # Add name and path
		}
		else
		{
			# If the software is in the dictionary
			$autorun_software[$app + '_'] = (Get-ItemProperty -Path $reg).$app # Add new name and path 
		}
	}
}
return $autorun_software | ConvertTo-Json