$computer_name = [ADSI]"WinNT://$env:COMPUTERNAME"
		
$list = @{} # Dictionary with groups that have at least one user

		
$computer_name.psbase.children | Where-Object { $_.psbase.schemaClassName -eq 'group' } | ForEach-Object {
	$group = [ADSI]$_.psbase.Path
	$group.psbase.Invoke("Members") | ForEach-Object {
		$us = $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
		$group_name = $group.Name.replace('{', '').replace('}', '')
		if ($group_name -notin $list.Keys)
		{
			$list[$group_name] = @($us) # If the group is not in the dictionary, add it along with the user's name
		}
		else
		{
			$list[$group_name] += $us # If the group exists in the dictionary, add the user's name to it
		}
	}
}

$groups = gwmi Win32_Group -Filter "LocalAccount='$True'" # Local groups of the computer
$groups_dict = @{} # Dictionary with data about groups
foreach ($group in $groups.Name)
{
	if ($group -in $list.Keys)
	{
		$groups_dict[$group] = $list[$group] # If there are users in the group, add the users to the dictionary
	}
	else { $groups_dict[$group] = '' } # Add an empty string 
}
$product_type = (gwmi Win32_OperatingSystem).ProductType

if ($product_type -eq 1) {		
	$local_accounts = gwmi win32_useraccount -Filter "LocalAccount='$True'" | Select-Object -Property Name, Disabled, Description # Local users of the computer
}
else {
	$local_accounts = gwmi win32_useraccount | Select-Object -Property Name, Disabled, Description # Local users 
}
$local_accounts_dict = @{ }
foreach ($account in $local_accounts)
{
	:outer
	foreach ($key in $groups_dict.Keys)
	{
		if ($account.Name -in $groups_dict[$key])
		{
			# If the user account is part of any groups, add the groups to a new property named Group
			$account | Add-Member -NotePropertyName Group -NotePropertyValue $key # Add a new property with the name of the Group
			break :outer # Break
		}
	}
}
foreach ($account in $local_accounts)
{
	# Creating the output file
	$account_name = (Get-Culture).TextInfo.ToTitleCase($account.Name) # Capitalize the first letter of the user account's name
	$local_accounts_dict += @{ $account_name = @($account.Disabled, $account.Group, $account.Description) } # Add the user account and its properties to the dictionary
}
return New-Object -TypeName PSCustomObject -Property @{groups_dict=$groups_dict | ConvertTo-Json
                                                       local_accounts_dict=$local_accounts_dict | ConvertTo-Json
                                                       }