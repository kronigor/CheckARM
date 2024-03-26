$OSArchitecture = (gwmi win32_operatingSystem).OSArchitecture # Checking architecture (64 or 86)
if ($OSArchitecture -like '*64*')
{
	if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\Data')
	{
		$name_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\Data').LocalizedProductName # AVZ name
		$server_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\Data').ServerName # AVZ server name
	}
	else 
	{
		if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES.21.8\Data')
		{
			$name_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES.21.8\Data').LocalizedProductName # AVZ name
			$server_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES.21.8\Data').ServerName # AVZ server name
		}	
	} 
	if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\environment')
	{
		$vers_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\environment').Ins_ProductVersion # AVZ version
	}
	if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\watchdog\BasesInfo')
	{
		$base_date = Get-Date([datetime]([datetime]::FromFileTime((Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\watchdog\BasesInfo').Date))) -Format "dd/MM/yyyy" # Date and time of AVZ database update
	}
	else 
	{
		if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES.21.8\watchdog\BasesInfo')
		{
			$base_date = Get-Date([datetime]([datetime]::FromFileTime((Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES.21.8\watchdog\BasesInfo').Date))) -Format "dd/MM/yyyy" # Date and time of AVZ database update
		}
	} 
}
else
{
	if (Test-Path -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\Data')
	{
		$name_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\Data').LocalizedProductName # AVZ name
		$server_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\Data').ServerName # AVZ server name 
	}
	else 
	{
		if (Test-Path -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES.21.8\Data')
		{
			$name_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES.21.8\Data').LocalizedProductName # AVZ name
			$server_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES.21.8\Data').ServerName # AVZ server name
		}	
	} 
	if (Test-Path -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\environment')
	{
		$vers_avz = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\environment').Ins_ProductVersion # Версия АВЗ
	}
	if (Test-Path -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES\watchdog\BasesInfo')
	{
		$base_date = Get-Date([datetime]([datetime]::FromFileTime((Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\KasperskyLab\protected\KES\watchdog\BasesInfo').Date))) -Format "dd/MM/yyyy" # Date and time of AVZ database update
	}
	else 
	{
		if (Test-Path -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES.21.8\watchdog\BasesInfo')
		{
			$base_date = Get-Date([datetime]([datetime]::FromFileTime((Get-ItemProperty -Path 'HKLM:\SOFTWARE\KasperskyLab\protected\KES.21.8\watchdog\BasesInfo').Date))) -Format "dd/MM/yyyy" # Date and time of AVZ database update
		}
	} 
}
$avz_info = @()
if (!$name_avz) {
	$name_avz_list = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object -Unique displayName
	foreach ($name_avz in $name_avz_list) {
		$avz_info += @{
		"Name"    = $name_avz.displayName;
		"Version" = $vers_avz;
		"Server"  = $server_avz;
		"Base"    = $base_date
		}
	}
}
else {
	$avz_info += @{
		"Name"    = $name_avz;
		"Version" = $vers_avz;
		"Server"  = $server_avz;
		"Base"    = $base_date
	}
}
return $avz_info | ConvertTo-Json