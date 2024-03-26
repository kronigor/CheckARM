### Information blocks ($local = 1 - local check without WMI) ###

$Section1 = ###1.General Information###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section1.ps1"
	
	try {
		if (!$local) { $comp_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $comp_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section1]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($comp_info) { $comp_info > "$($path)\src\jsons\$($computer)\1_comp.json" }
	} catch { }
}

$Section2 = ###2.Operating System###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section2.ps1"
	
	try {
		if (!$local) { $os_info = Invoke-Command -ComputerName $computer -FilePath $script } 
        else { $os_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section2]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
        if ($os_info) { $os_info > "$($path)\src\jsons\$($computer)\2_os.json" } 
		
	} catch { }
}

$Section3 = ###3.Network Configuration###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section3.ps1"
	
	try {
		if (!$local) { $netadapters_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else {
			$netadapters_info = Invoke-Expression -Command $script
			$Error.clear()
		}
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section3]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($netadapters_info) { $netadapters_info > "$($path)\src\jsons\$($computer)\3_net.json" }
	} catch { }
}

$Section4 = ###4.Time Service###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section4.ps1"
	
	try {
		if (!$local) { $measure_time_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $measure_time_info = Invoke-Expression -Command $script }
		if ($Error) {
			"$((Get-Date).ToString())`n[Section4]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($measure_time_info) { $measure_time_info > "$($path)\src\jsons\$($computer)\4_time.json" }
	} catch { }
}

$Section5 = ###5.OS Updates###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section5.ps1"
	
	try {
		if (!$local) { $hotfixes_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $hotfixes_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section5]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($hotfixes_info) { $hotfixes_info > "$($path)\src\jsons\$($computer)\5_hotfixes.json" }

	} catch { }
}

$Section6 = ###6. Accounts - 7. Groups###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section6.ps1"
	
	try {
		if (!$local) { $accounts_and_users_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $accounts_and_users_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section6]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
        if ($accounts_and_users_info.groups_dict) { $accounts_and_users_info.groups_dict  > "$($path)\src\jsons\$($computer)\7_groups.json" }
        if ($accounts_and_users_info.local_accounts_dict) { $accounts_and_users_info.local_accounts_dict > "$($path)\src\jsons\$($computer)\6_accounts.json" }
	} catch { }
}

$Section8 = ###8. Shared Folders###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section8.ps1"
	
	try {
		if (!$local) { $shares_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $shares_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section8]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($shares_info) { $shares_info > "$($path)\src\jsons\$($computer)\8_shares.json" }
	} catch { }
}

$Section9 = ###9. AVZ###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section9.ps1"
	
	try
	{
		if (!$local) { $avz_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $avz_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section9]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($avz_info) { $avz_info > "$($path)\src\jsons\$($computer)\9_avz.json" }

	} catch { }
}

$Section10 = ###10. Software###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section10.ps1"
	try
	{
		if (!$local) { $software_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $software_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section10]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		} 
        if ($software_info.std_software_dict) { $software_info.std_software_dict > "$($path)\src\jsons\$($computer)\10_1_std_software.json" }
		if ($software_info.sys_software_dict)	{ $software_info.sys_software_dict > "$($path)\src\jsons\$($computer)\10_2_sys_software.json" }
	} catch { }
}

$Section11 = ###11. Autostart###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section11.ps1"
	
	try {
		if (!$local) { $autorun_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $autorun_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section11]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
        if ($autorun_info) { $autorun_info > "$($path)\src\jsons\$($computer)\11_autorun.json" }
	} catch { }
}

$Section12 = ###12. Task Scheduler###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section12.ps1"
	
	try {
		if (!$local) { $tasks_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $tasks_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section12]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($tasks_info) { $tasks_info > "$($path)\src\jsons\$($computer)\12_tasks.json" }

	} catch { }
}

$Section13 = ###13. OS Event Logs###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section13.ps1"
	
	try {
		if (!$local) { $logs_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $logs_info = Invoke-Expression -Command $script }
		if ($Error -and $logs_info.security_EventLog -and $logs_info.system_EventLog) {
            "$((Get-Date).ToString())`n[Section13]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
        }
		$Error.clear()
        if ($logs_info.system_EventLog) {$logs_info.system_EventLog > "$($path)\src\jsons\$($computer)\13_1_systemlog.json" }
		if ($logs_info.security_EventLog)   { $logs_info.security_EventLog > "$($path)\src\jsons\$($computer)\13_2_securitylog.json" }
	} catch { }
}

$Section14 = ###14. OS Services###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section14.ps1"
	
	try {
		if (!$local) { $services_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $services_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section14]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		} 
        if ($services_info) { $services_info > "$($path)\src\jsons\$($computer)\14_services.json" }
	} catch { }
}

$Section16 = ###16. Password Policy###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Int]$local
	)
	
	$ErrorActionPreference = "SilentlyContinue"
	$script = "$($path)\src\scripts\section16.ps1"
	
	try {
		if (!$local) { $pass_politics_info = Invoke-Command -ComputerName $computer -FilePath $script }
		else { $pass_politics_info = Invoke-Expression -Command $script }
		
		if ($Error) {
			"$((Get-Date).ToString())`n[Section16]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
			$Error.clear()
		}
		if ($pass_politics_info) { $pass_politics_info > "$($path)\src\jsons\$($computer)\16_pass.json" }

	} catch { }
}

$Section15 = ### Report Creation (Python) ###
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
        [Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[String]$report_version,
		[Parameter(Mandatory = $true)]
		[String]$build_number
	)
	# Running a Python program to generate a Word report
    & "$path\src\reporting\report_create.exe" -c $computer -p $path -rv $report_version -ov $build_number -r 90| Tee-Object -Variable out_path
	if ($out_path -eq 'Error') {
        "$((Get-Date).ToString())`n[Section 15]`nPlease check the error log at 'errors\python_errors.log' for details." >> "$($path)\errors\$($computer).txt"
    }
    $Error.clear()
	return $out_path # Returning the name of the saved report
}

### WINRM Service Start ###
Function EnableWinRM 
{
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer
	)
	$ErrorActionPreference = "SilentlyContinue"
    $Error.clear()
	try
	{
		$winrmStatus = Get-Service -Name WinRM -ComputerName $computer | Select-Object Status, Name, StartType, DisplayName
		if ($winrmStatus.Status -ne "Running") {
			Get-Service -Name WinRM -ComputerName $computer | Start-Service # Starting the winrm service
		}
		if ($Error)
		{
			"$((Get-Date).ToString())`n[EnableWinRM]`n$($Error[0])" >> "$path\errors\$($computer).txt"
			$Error.clear()
            return
		}
	} catch {$Error.clear()}
	return
}

### Checking the Existence of a Dump Directory for a Specific ARM and Clearing ###

Function CheckDirectory {

	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
        [Parameter(Mandatory = $true)]
		[Boolean]$report_only
	)
	# Checking for the existence of a folder to dump files

    if ($report_only) {
        if (!(Test-Path -Path "$path\src\jsons\$($computer)\")) {
            Throw "Missing directory with JSON files."
        }
    }
    else {
	    if (!(Test-Path -Path "$path\src\jsons\$($computer)\")) {
            # If the folder does not exist, create it
		    New-Item -Path "$path\src\jsons\$($computer)\" -ItemType Directory
        } 
    }
	return
}
### Checking ARM Availability ###
Function TestConn {

    param(
        [Parameter(Mandatory = $true)]
		[String]$computer
    )
    # Ping 1 packet
    try { $test = Test-Connection -Count 1 -ComputerName $computer -Quiet } catch {
        # If the IP cannot be obtained by the hostname
        "$((Get-Date).ToString())`n[TestConn]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
        return 0, $computer
    } 
    if (!$test) { 
        # If the host is not accessible
        "$((Get-Date).ToString())`n[TestConn]`nArm is unavailable!" >> "$($path)\errors\$($computer).txt"
        return 0, $computer
    }
    # If an IP address is specified
    $pattern = "(?:25[0-5]|2[0-4][0-9]|[1][0-9]?[0-9]?)\.((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)"
    if ($computer -match "^$($pattern)$") {
        $computer_old = $computer
        try {
            # Getting the ARM's name by connecting to it
            $computer = (gwmi win32_computersystem -ComputerName $computer -Property Name).Name
        } catch {
            "$((Get-Date).ToString())`n[TestConn]`n$($Error[0])" >> "$($path)\errors\$($computer_old).txt"
            return 0, $computer_old
        }
        if ($Error) {
            "$((Get-Date).ToString())`n[TestConn]`n$($Error[0])" >> "$($path)\errors\$($computer_old).txt"
		    $Error.clear()
            return 0, $computer_old
        }
    }
    return 1, $computer
}
### Checking winrm on ARM ###
Function TestInvoke {
    
    param(
        [Parameter(Mandatory = $true)]
		[String]$computer
    )

    try {
        # Enabling the winrm service
        EnableWinRM -computer $computer
        # Checking if we can execute commands remotely on the ARM and getting the OS type (client/server)
		Invoke-Command -ComputerName $computer -ScriptBlock {} 
    }
    catch
    {
        "$((Get-Date).ToString())`n[TestInvoke]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
	    return 0
    } 
    if ($Error) {
        "$((Get-Date).ToString())`n[TestInvoke]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
	    $Error.clear()
	    return 0
    }
    return 1
}
### Checking for Standard Directories (Creation if Missing). Clearing JSON ###
Function DefaultDirectories {
    
    param(
		[Parameter(Mandatory = $true)]
		[String]$computer,
        [Parameter(Mandatory = $true)]
		[String]$path,
        [Parameter(Mandatory = $true)]
		[Boolean]$report_only
    )

    try
	{
        # Checking for the existence of a folder to dump files
		if (!(Test-Path -Path "$path\src\jsons\")) {
			New-Item -Path "$path\src\jsons\" -ItemType Directory # If the folder does not exist, create it
		}
        if (!$report_only) {
            # If the folder exist, delete it
		    Remove-Item "$path\src\jsons\*" -Force -Recurse -Confirm:$false 
		}
        # Deleting all files in the JSON folder
        if (!(Test-Path -Path "$path\reports\")) {
			New-Item -Path "$path\reports\" -ItemType Directory # If the folder does not exist, create it
		}
        # Checking for the existence of a folder to dump files
		if (!(Test-Path -Path "$path\errors\")) {
			New-Item -Path "$path\errors\" -ItemType Directory # If the folder does not exist, create it
		}
		
	} catch {
		"$((Get-Date).ToString())`n[DefaultDirectories]`n$($Error[0])" >> "$path\errors\$($computer).txt"
		return 0
	}

}
### Checking a Single ARM ###
Function Single {

    # If the name of the ARM on which the program is run matches the user-entered name, set local mode
    if ($env:computername -eq $computer) { $local = 1 } else { 
    # If the Name does not match, perform the ARM availability check
    $local = 0
    # Checking ARM availability
    $testconn, $computer = TestConn -computer $computer
    if ($testconn) {
        $testinvoke = TestInvoke -computer $computer
        # If the ARM is pingable, but there are no rights for WMI
        if (!$testinvoke) { 
            write-host -ForegroundColor Red "`n[ERROR] '$($computer)' is unavailable!"
            if ($mode) { break main } else {return}
        } 
    } else {
        write-host -ForegroundColor Red "`n[ERROR] '$($computer)' is unavailable!"
        if ($mode) { break main } else {return}
    }
	}
		
	# Import functions
    . $path\src\functions\single.ps1
     
    # Starting the check *\functions\single.ps1*
    $result = $GetSettingsSingle.Invoke($computer, $report_version, $build_number, $script_threads, $local, $measure_time, $json_only, $report_only)
	if ($result[-1])
	{
		return $result[-1]
	} else { $result[-2]  }
}

### Checking Multiple ARMs from a File ###
Function Multiple {
    # Checking for the presence of a file with names of unchecked ARMs
    if (Test-Path -Path "$path\unchecked_arm.txt") {
        # Deleting the file if it exists
        Remove-Item "$path\unchecked_arm.txt" -Force -Recurse -Confirm:$false 
    }
    if (!$report_only) {
        # If an IP address range is specified
        if ($range -match "^$($pattern)-$($pattern)$") {
            # Array for IP addresses
            $file = @()
            # Starting and ending IP
            $range = $range -split "-"
            # The first three octets of the starting IP
            $range[0] -match "\d+\.\d+\.\d+" | Out-Null
            $ip_0 = $matches[0]
			The first three octets of the ending IP
            $range[1] -match "\d+\.\d+\.\d+" | Out-Null
            $ip_1 = $matches[0]
            $range[0] -match "\d+$" | Out-Null
            # The last octet of the starting IP
            $ip_start = $matches[0]
            $range[1] -match "\d+$" | Out-Null
            # The last octet of the ending IP
            $ip_end = $matches[0]

            if (($ip_0) -ne ($ip_1)) {
			    write-host -ForegroundColor Red "`n[ERROR] An IP address range is wrong!"
			    return
            } else {
                foreach ($_ in $ip_start..$ip_end) {
                    $file += "$($ip_0).$($_)"
                }
            }
        }
        # If a text file is specified
        elseif ($range.EndsWith(".txt")) {
            # Loading the list of ARMs into a variable
            try { $file = Get-Content -Path $range } catch {
                write-host -ForegroundColor Red "`n[ERROR] Unable to open the file!"
                if ($mode) { break multiple } else {return}
            }
        }
        # If a CSV file is specified
        else { 
            # Loading the list of ARMs into a variable
	        try { $file = (Import-Csv -Path $range -Header Header -Encoding UTF8).Header } catch {
                write-host -ForegroundColor Red "`n[ERROR] Unable to open the file!"
                if ($mode) { break multiple } else {return}
            }
        }
    }
    # Import functions
    . $path\src\functions\multi.ps1 

   # Starting the check *\functions\multi.ps1**       
    $result =  $CheckMultiple.Invoke($file, $arms_threads, $report_version, $build_number, $script_threads, $path, $measure_time, $json_only, $report_only)
    return $result[-1]
}
### Checking Multiple ARMs from OU Active Directory ###
Function Ad {
    # Name of the PS cmdlet
    $cmdName = "Get-AdComputer"
    # Checking for the presence of the cmdlet
    if (!(Get-Command $cmdName -errorAction SilentlyContinue)) {
        write-host -ForegroundColor Red "`n[ERROR] Commandlet 'Get-AdComputer' doesn't exist! You need to install RSAT."
        if ($mode) { break main } else {return}
    }
    # Getting the names of ARMs from the specified OU
    try {
        $file = (Get-ADComputer -SearchBase $organizational_unit -Filter * -properties Name | Select-Object Name | Sort-Object Name).Name
    } catch {
        write-host -ForegroundColor Red "`n[ERROR] OU is wrong or empty!."
        if ($mode) { break main } else {return}
    }
    # Import functions
    . $path\src\functions\multi.ps1 

    # Starting the check *\functions\multi.ps1*
    $result =  $CheckMultiple.Invoke($file, $arms_threads, $report_version, $build_number, $script_threads, $path, $measure_time, $json_only, $report_only)
    return $result[-1]
}