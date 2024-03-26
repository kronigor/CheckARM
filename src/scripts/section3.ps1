$netadapters_info = @{} # Dictionary for network adapters
$netadapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } # All WORKING network adapters
foreach ($netadapter in $netadapters)
{
	$netadapter_settings = Get-NetIPConfiguration -InterfaceAlias $netadapter.Name -Detailed # Getting detailed settings for a specific adapter
	foreach ($setting in $netadapter_settings)
	{
		# Iterating through adapter settings    
		$settings = @() # List with settings
		$settings += @{ 'ip' = $setting.IPv4Address.IPAddress } # ip address 
		if ($null -ne $setting.IPv4DefaultGateway.NextHop)
		{
			$settings += @{ 'gateway' = $setting.IPv4DefaultGateway.NextHop } # gateway
		}
		if ($null -ne $setting.DNSServer.ServerAddresses)
		{
			# dns
			$settings += @{ 'dns' = $setting.DNSServer.ServerAddresses }
		}
		$settings += @{ 'mac' = $setting.NetAdapter.LinkLayerAddress } # mac
	}
	$netadapters_info.add($netadapter.Name, $settings)
}

return $netadapters_info | ConvertTo-Json