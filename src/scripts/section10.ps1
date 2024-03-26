$OSArchitecture = (gwmi win32_operatingSystem).OSArchitecture # Checking architecture (64 or 86)
if ($OSArchitecture -like '*64*')
{
	if (Test-Path -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
	{
		$list1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
	}
	if (Test-Path -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\')
	{
		$list2 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
	}
	if (Test-Path -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
	{
		$list3 = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion
	}
				
	$std_software = $list1 + $list2 + $list3 | Where-Object { $_.DisplayName -notlike '*Update for*' } | Sort-Object DisplayName | Get-Unique -AsString # Removed empty lines and duplicates
}
			
else
{
	if (Test-Path -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
	{
		$list1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
	}
	if (Test-Path -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
	{
		$list2 = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
	}
	$std_software = $list1 + $list2 | Where-Object { $_.DisplayName -notlike '*Update for*' } | Sort-Object DisplayName | Get-Unique -AsString # Removed empty lines and duplicates
}

$product_type = (gwmi Win32_OperatingSystem).ProductType
if ($product_type -eq 1)
{
	$sys_software = Get-AppxPackage -allusers | Select-Object Name, Version | Sort-Object Name | Get-Unique -AsString # System software (metro applications)
	
	return New-Object -TypeName PSCustomObject -Property @{
		sys_software_dict = $sys_software | ConvertTo-Json
		std_software_dict = $std_software | ConvertTo-Json
	}
}
else
{
	$roles = Get-WindowsFeature | Select-Object Name, InstallState, Description | Where-Object { [string]$_.InstallState -in 'Installed', 'Removed' } | Sort-Object Name
	return New-Object -TypeName PSCustomObject -Property @{
		sys_software_dict = $roles | ConvertTo-Csv | ConvertFrom-Csv | ConvertTo-Json
		std_software_dict = $std_software | ConvertTo-Json
	}
}