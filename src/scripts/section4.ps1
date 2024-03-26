function ConvertTo-Encoding ([string]$From, [string]$To)
{
	# Converting to cp866 encoding
	Begin
	{
		$encFrom = [System.Text.Encoding]::GetEncoding($from)
		$encTo = [System.Text.Encoding]::GetEncoding($to)
	}
	Process
	{
		$bytes = $encTo.GetBytes($_)
		$bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)
		$encTo.GetString($bytes)
	}
}

$time_service = gwmi win32_service | Where-Object { $_.name -like '*w32time*' } # Searching for time service
if ($time_service.State -eq 'Running')
{
	# If the service is running
	$state = 'Служба работает' # Service status
	$ntp = w32tm /query /status 
	if ($ntp) # If NTP server name is obtained
	{
		$ntp_source = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters").NtpServer.Split(',')[0]
		try
		{
			$ntpip = [System.Net.Dns]::GetHostAddresses($ntp_source.trim()).IPAddressToString  # IP ntp
		}
		catch
		{
			$ntpip = $ntp_source
			$Error.clear()
		}
	}
			
	$last_time = $ntp -match 'Время последней успешной синхронизации:' # Date of the last synchronization
	if (!$last_time) # If the last successful synchronization time is obtained
	{
		$last_time = $ntp | ConvertTo-Encoding -From cp866 -To windows-1251  | Where-Object { $_ -match 'Время последней успешной синхронизации:' }
	}
    try
	{
        $last_time = $last_time.Replace('Время последней успешной синхронизации: ', '')
    }
    catch
	{ 
		$last_time = 'None'
		$Error.clear()
	}
}
else
{
	$state = 'Служба отключена' # Service status
}
		
$time_zone = gwmi win32_timeZone | Select-Object caption # Time zone		
		
return new-object psobject -property @{
	State	  = $state
	NtpSource = $ntpip
	TimeZone  = $time_zone.caption
	Last_time = $last_time
} | ConvertTo-Json
