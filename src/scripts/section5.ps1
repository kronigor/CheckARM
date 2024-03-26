$wsus_service = gwmi win32_service | Where-Object { $_.name -like '*wuauserv*' } # Searching for the Windows Update service
$usosvc_service = gwmi win32_service | Where-Object { $_.name -like '*UsoSvc*' } # Searching for the Update Orchestrator service
if ((($wsus_service.StartMode -eq 'Manual') -or ($wsus_service.StartMode -eq 'Auto')) -and ($usosvc_service.State -eq 'Running'))
{
	# If the Windows Update service is set to Manual, and the Orchestrator service is running
	$state = "Служба работает`n" + "(Статус: '$($wsus_service.State)')" # Status of the WSUS service
}
else
{
	$state = 'Служба отключена' # Status of the WSUS service
}

if (Test-Path -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate')
{
	# Checking if the registry path exists
	$wsusServer = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate').WUServer # WSUS address
	if (!$wsusServer)
	{
		# If there is no WSUS address
		$wsusip = 'None'
	}
	else
	{
		# If a WSUS address exists
		$wsusServer = $wsusServer.Trim('http://').replace(':8530', '') # Retain only the address
		$wsusip = [System.Net.Dns]::GetHostByName($wsusServer).AddressList.IPAddressToString # WSUS IP
	}
				
}
else
{
	# If the registry path does not exist
	$wsusip = 'None'
}

$hotfixes = Get-Hotfix | Select-Object HotFixID, InstalledOn, Description
$count = ($hotfixes | Measure-Object).count # Number of OS updates
$hotfixes = $hotfixes | Select-Object -Last 5  # Last 5 OS updates
$hotfixes_dict = @{} # Dictionary for creating the final output file
foreach ($hotfix in $hotfixes)
{
	if ($hotfix.InstalledOn)
	{
		$hotfix.InstalledOn = $hotfix.InstalledOn.ToShortDateString() # Removing time from update dates
	}
	else
	{
		$hotfix.InstalledOn = Get-Date -Format 'dd.MM.yyyy' # If installation date is missing, add current date
	}
	if (!$hotfix.Description)
	{
		$hotfix.Description = '' # If there is no description, add an empty string
	}
	$hotfixes_dict += @{ $hotfix.HotFixID = @($hotfix.InstalledOn, $hotfix.Description) }
}

$hotfixes_dict += @{ 'Other' = @($state, $wsusip, $count) }
return $hotfixes_dict | ConvertTo-Json