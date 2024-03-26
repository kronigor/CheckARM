<#
.SYNOPSIS
    The program checks the computer confuguration.
.DESCRIPTION
    The program checks the computer confuguration.
    The following information is available:
    The short report:
        1. General Information;
        2. Operating System;
        3. Network Configuration
        4. Network Time Protocol
        5. Windows Updates
        6. Local Accounts
        7. Local Groups
        8. Shared Folders
        9. Antivirus Software
        10. Third-Party Software
        11. Password Policy
    The full report (include the short report):
        12. Autorun
        13. Task Sheduler
        14. System Software (embedded)
        15. Services
        16. System Event Log
        17. Security Event Log    
.PARAMETER Path
    The path to the .
.PARAMETER LiteralPath
    Specifies a path to one or more locations. Unlike Path, the value of
    LiteralPath is used exactly as it is typed. No characters are interpreted
    as wildcards. If the path includes escape characters, enclose it in single
    quotation marks. Single quotation marks tell Windows PowerShell not to
    interpret any characters as escape sequences.
.EXAMPLE
    <Display a help message>
        CheckARM_console.exe -h
    
    <Create a report only>
        CheckARM_console.exe -ro
.EXAMPLE
    <Single Mode>
    
    <Use the default options>
        CheckARM_console.exe -c PCNAME
        CheckARM_console.exe -c 192.168.0.1
    
    <Use the custom options>
        CheckARM_console.exe -c PCNAME -rv short -bn 19000 -st 4
   
    <Measure the execution time>
        CheckARM_console.exe -c 192.168.0.1 -mt
    
    <Create JSON files without a report>
        CheckARM_console.exe -c PCNAME -jo
.EXAMPLE
    <Multiple Mode>

    <Use the default options>
        CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt
        CheckARM_console.exe -f C:\Users\User\Desktop\arms.csv
        CheckARM_console.exe -i 192.168.0.1-192.168.0.100

    <Use the custom options>
        CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt -rv short -bn 19000 -st 5 -at 10

    <Measure the execution time>
        CheckARM_console.exe -i 192.168.0.1-192.168.0.100 -st 3 -mt

    <Create JSON files without a report>
        CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt -jo
.EXAMPLE
    <Active Directory Mode>

    <Use the default options>
        CheckARM_console.exe -o OU=testou,DC=testdc

    <Use the custom options>
        CheckARM_console.exe -o OU=testou,DC=testdc -rv short -bn 19000 -at 2

    <Measure the execution time>
        CheckARM_console.exe -o OU=testou,DC=testdc -mt

    <Create JSON files without a report>
        CheckARM_console.exe -o OU=testou,DC=testdc -jo
.NOTES
    Author: kronigor
    Date:   March, 2024
#>

[CmdletBinding(DefaultParametersetName = "single_mode", PositionalBinding = $false)]
param (
	[Alias("c")]
	[Parameter(Mandatory = $False, ParameterSetName = "single_mode", HelpMessage = "Enter a computer name or IP address.")]
	[ValidateScript({
			if (($_ -match "^(?:25[0-5]|2[0-4][0-9]|[1][0-9]?[0-9]?)\.((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$") -or
				($_ -match "^[A-Za-z0-9][A-Za-z0-9\-_]{1,14}$")) { $true }
			else { Throw "The computer name or IP address is empty or wrong!" }
		})]
	[String]$computer = $env:computername,
	[Alias("f")]
	[Parameter(Mandatory = $False, ParameterSetName = "multiple_mode", HelpMessage = "Enter a file that contains the computer names.")]
	[ValidateScript({
			if (($_.ToLower().Endswith(".txt") -or $_.ToLower().Endswith(".csv")) -and (Test-Path $_ -PathType leaf)) { $true }
			else { Throw "Path doesn't exist!" }
		})]
	[String]$filename,
	[Alias("i")]
	[Parameter(Mandatory = $False, ParameterSetName = "iprange_mode", HelpMessage = "Enter an IP address range.")]
	[ValidateScript({
			if ($_ -match "^(?:25[0-5]|2[0-4][0-9]|[1][0-9]?[0-9]?)\.((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)-(?:25[0-5]|2[0-4][0-9]|[1][0-9]?[0-9]?)\.((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$") { $true }
			else { Throw "An IP adress range is wrong or empty (192.168.0.1-192.168.0.100)" }
		})]
	[String]$iprange,
	[Alias("ou")]
	[Parameter(Mandatory = $False, ParameterSetName = "ad_mode", HelpMessage = "Enter an Organizational Unit.")]
	[ValidatePattern("^(OU|DC)=\w+[\s,\.\w=]*$")]
	[String]$organizational_unit,
	[Alias("rv")]
	[Parameter(Mandatory = $False, HelpMessage = "Enter a version report.")]
	[ValidateSet("full", "short")]
	[String]$report_version,
	[Alias("st")]
	[Parameter(Mandatory = $False, HelpMessage = "Enter a number of running scriptblocks in parallel (min - 1, max - 30).")]
	[ValidateRange(1, 30)]
	[Int32]$script_threads,
	[Alias("bn")]
	[Parameter(Mandatory = $False, HelpMessage = "Enter an Operating System version.")]
	[ValidatePattern("^[0-9]+$")]
	[String]$build_number,
	[Alias("at")]
	[Parameter(Mandatory = $False, ParameterSetName = "multiple_mode", HelpMessage = "Enter a number of computers checked at the same time (in parallel).")]
	[Parameter(Mandatory = $False, ParameterSetName = "iprange_mode", HelpMessage = "Enter a number of computers checked at the same time (in parallel).")]
	[Parameter(Mandatory = $False, ParameterSetName = "ad_mode", HelpMessage = "Enter a number of computers checked at the same time (in parallel).")]
	[ValidateScript({
			if ([Int32]$_ -ge 1) { $true }
			else { Throw "The number of computers must be [Int32] and -ge 1" }
		})]
	[Int32]$arms_threads,
	[Alias("h")]
	[Parameter(Mandatory = $False, ParameterSetName = "help", HelpMessage = "Show help message.")]
	[Switch]$help = $false,
	[Alias("mt")]
	[Parameter(Mandatory = $False, HelpMessage = "Measure the program's execution time.")]
	[Switch]$measure_time,
	[Alias("jo")]
	[Parameter(Mandatory = $False, ParameterSetName = "single_mode", HelpMessage = "Create JSON files without a report.")]
	[Parameter(Mandatory = $False, ParameterSetName = "multiple_mode", HelpMessage = "Create JSON files without a report.")]
	[Parameter(Mandatory = $False, ParameterSetName = "iprange_mode", HelpMessage = "Create JSON files without a report.")]
	[Parameter(Mandatory = $False, ParameterSetName = "ad_mode", HelpMessage = "Create JSON files without a report.")]
	[Parameter(Mandatory = $False, ParameterSetName = "json_only", HelpMessage = "Create JSON files without a report.")]
	[Switch]$json_only,
	[Alias("ro")]
	[Parameter(Mandatory = $False, ParameterSetName = "report_only", HelpMessage = "Create a report only.")]
	[Switch]$report_only
)

#NonInteractive mode 
Function NonInteractiveMode
{
	
	param (
		[Alias("arg")]
		[Parameter(Mandatory = $true)]
		$arguments
	)
	$ErrorActionPreference = "Stop"
	try
	{
		# Import functions
		. $path\src\functions\main.ps1
		. $path\src\functions\scriptblocks.ps1

		# Check that the script is run as admin
		if (! (IsAdmin))
		{
			write-host -ForegroundColor Red "`n[ERROR] Only admin may check the ARMs!"
			write-host "`nEnter any key to exit"
			write-host -ForegroundColor Yellow "-->: " -NoNewLine
			$choice = read-host
			return
		}
		else
		{
			# If report version is specified
			if ($arguments.ContainsKey('report_version'))
			{
				$report_version = $arguments.report_version
			}
			
			 # If OS build number is specified
			if ($arguments.ContainsKey('build_number'))
			{
				$build_number = $arguments.build_number
			}
			
			 # If the number of simultaneously executed blocks is specified
			if ($arguments.ContainsKey('script_threads'))
			{
				$script_threads = $arguments.script_threads
			}
			
			 # If the number of ARMs to be checked at the same time is specified
			if ($arguments.ContainsKey('arms_threads'))
			{
				$arms_threads = $arguments.arms_threads
			}
			
			# If the parameter for time calculation is specified
			if ($arguments.ContainsKey('measure_time'))
			{
				$measure_time = $true
			}
			 # If the parameter for exporting json without report is specified
			if ($arguments.ContainsKey('json_only'))
			{
				$json_only = $true
			}
			# If the parameter for generating report without exporting json is specified
			if ($arguments.ContainsKey('report_only'))
			{
				$report_only = $true
				$file = (Get-ChildItem -Path "$($path)\src\jsons\" -Directory).Name
				if ($file.Count -eq 0)
				{
					write-host -ForegroundColor Red "`n[ERROR] '\src\jsons\' folder is empty!"
					return
				}
				
				Multiple | Out-Null
				
			}
			# If the help marker is specified
			elseif ($arguments.ContainsKey('help'))
			{
				ShowHelp
			}
			# If the file with names is specified
			elseif ($arguments.ContainsKey('filename'))
			{
				$range = $filename
				Multiple | Out-Null
			}
			# If the IP range is specified
			elseif ($arguments.ContainsKey('iprange'))
			{
				$range = $iprange
				Multiple | Out-Null
			}
			# If OU is specified
			elseif ($arguments.ContainsKey('ou'))
			{
				Ad | Out-Null
			}
			# Checking single ARM
			else
			{
				Single | Out-Null
			}
		}
	}
	catch
	{
		Write-Host -ForegroundColor Red "[ERROR] Please check the error log at 'errors\' for details."
		"$((Get-Date).ToString())`n[NonInteractiveMode]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
	}
	return
}

#Interactive mode     
Function InteractiveMode
{
	$ErrorActionPreference = "SilentlyContinue"
	try
	{
		# Import functions
		. $path\src\functions\main.ps1
		. $path\src\functions\scriptblocks.ps1
		# Displaying the logo
		Invoke-Command $logo
		# The default checking mode - single ARM
		$mode = "single"
		# Check that the script is run as admin
		if (! (IsAdmin))
		{
			write-host -ForegroundColor Red "`n[ERROR] Only admin may check the ARMs!"
			write-host "`nEnter any key to exit"
			write-host -ForegroundColor Yellow "-->: " -NoNewLine
			$choice = read-host
			return
		}
		else
		{
			while ($true)
			{
				# Displaying the menu
				Invoke-Command $main_menu
				:main while ($true)
				{
					# Reading user input
					write-host -ForegroundColor Yellow "`n-->: " -NoNewLine
					$choice = read-host
					Switch ($choice)
					{
						1 {
							 # Main menu status
							$menu = $false
							$ErrorActionPreference = "Stop"
							# ARM name
							$computer = $env:computername
							write-host -ForegroundColor Yellow "`n[1] Single mode (one ARM)"
							
							:single while ($true)
							{
								# Reading user input
								write-host "`nThe computer name (IP address) is: " -NoNewLine
								write-host -ForegroundColor DarkCyan "'$($computer)'"
								write-host "Do you want to change the computer name (IP address)? [y/N]"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								$choice = read-host
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") -or ([String]::IsNullOrWhiteSpace($_)) } { break single }
									
									{ ($_.ToLower() -eq "y") } {
										# Validating the entered ARM name *functions\main.ps1*
										$computer = ValidateComputer -computer $computer
										 # Exiting to the main menu
										if ($computer -eq "m")
										{
											$computer = $env:computername
											break main
										}
										write-host "`nThe computer name is: " -NoNewLine
										write-host -ForegroundColor DarkCyan "'$($computer)'"
										break single
									}
									 # Exiting to the main menu
									m { break main }
									
									Default {
										write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' or 'm' - back to Main menu."
									}
								}
							}
							# Additional options block
							:options while ($true)
							{
								# Reading user input
								write-host "Do you want to enter the advanced options? [y/N]"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								$choice_ = read-host
								Switch ($choice_)
								{
									{ ($_.ToLower() -eq "n") -or ([String]::IsNullOrWhiteSpace($_)) } { break options }
									{ ($_.ToLower() -eq "y") } {
										# Entering additional options and validation functions\main.ps1
										$report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only = Invoke-Command $advanced_params
										break options
									}
									# Exit to the main menu
									m { break main }
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' or 'm' - back to Main menu." }
								}
							}
							while ($true)
							{
								# Checking the computer
								# Reading user input
								write-host -ForegroundColor Yellow "Start checking the ARM '$($computer)'? " -NonewLine
								$choice = read-host "[Y/n] ('o' - show current options)"
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") } { break main }
									{ (($_.ToLower() -eq "y") -or ([String]::IsNullOrWhiteSpace($_))) } {
										# Starting the check functions\scriptblocks.ps1
										$result = Single
										CheckResult -result $result -path $path -mode $mode
										break main
									}
									# Show additional options functions\main.ps1
									o { ShowOptions -mode $mode -menu $false }
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' - back to Main menu." }
								}
							}
							
						}
						2 {
							# Main menu status
							$menu = $false
							# Checking mode - multiple computers from a file
							$mode = "multiple"
							write-host -ForegroundColor Yellow "`n[2] Multiple mode (several ARMs)`n"
							# Reading user input
							:multiple while ($true)
							{
								write-host "Enter a file that contains the computer names or an IP address range ('m' - back to Main menu)"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								# Validation
								try
								{
									[ValidateScript({
											if ((($_.ToLower().Endswith(".txt") -or $_.ToLower().Endswith(".csv")) -and (Test-Path $_ -PathType leaf)) -or ($_ -match "^$($pattern)-$($pattern)$") -or ($_ -eq "m")) { $true }
											else { Throw "Path doesn't exist or an IP address range is wrong!" }
										})][String]$range = Read-Host
									# Exit to the menu
									if ($range -eq "m") { break main }
									break multiple
								}
								catch
								{
									$range = ""
									write-host -ForegroundColor Red "`n[ERROR] Path doesn't exist or an IP address range is wrong! (need an valid IP range or a real path, *.txt or *.csv)`n"
								}
							}
							# Additional options block
							:options while ($true)
							{
								# Reading user input
								if ($range -match "^$($pattern)-$($pattern)$") { write-host "`nAn IP address range: " -NoNewLine }
								else { write-host "`nThe filename is: " -NoNewLine }
								write-host -ForegroundColor DarkCyan "'$($range)'"
								write-host "Do you want to enter the advanced options? [y/N]"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								$choice = read-host
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") -or ([String]::IsNullOrWhiteSpace($_)) } { break options }
									{ ($_.ToLower() -eq "y") } {
										# Entering additional options and validation functions\main.ps1
										$report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only = Invoke-Command $advanced_params
										break options
									}
									# Exit to the main menu
									m { break main }
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' or 'm' - back to Main menu." }
								}
							}
							while ($true)
							{
								# Checking the computer
								# Reading user input
								write-host -ForegroundColor Yellow "Start checking the ARMs? " -NonewLine
								$choice = read-host "[Y/n] ('o' - show current options)"
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") } { break main }
									{ (($_.ToLower() -eq "y") -or ([String]::IsNullOrWhiteSpace($_))) } {
										# Starting the check functions\scriptblocks.ps1
										$result = Multiple
										CheckResult -result $result -path $path -mode $mode
										break main
									}
									# Show additional options functions\main.ps1
									o { ShowOptions -mode $mode -menu $false }
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' - back to Main menu." }
								}
							}
						}
						
						3 {
							# Main menu status
							$menu = $false
							# Checking mode - multiple computers from OU in Active Directory
							$mode = "ad"
							write-host -ForegroundColor Yellow "`n[3] Active Directory mode (scan OUs)`n"
							
							:ad while ($true)
							{
								# Reading user input
								write-host "Enter an Organizational Unit ('m' - back to Main menu)"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								# Validation
								try
								{
									[ValidatePattern("^m$|^(OU|DC)=\w+[\s,\.\w=]*$")][String]$organizational_unit = Read-Host
									if ($organizational_unit -eq "m") { break main }
									break ad
								}
								catch
								{
									$organizational_unit = ""
									write-host -ForegroundColor Red "`n[ERROR] The Organizational Unit is empty or wrong! ('^OU|^DC)=\w')`n"
								}
							}
							# Additional options block
							:options while ($true)
							{
								# Reading user input
								write-host "`nThe Organizational Unit is: " -NoNewLine
								write-host -ForegroundColor DarkCyan "'$($organizational_unit)'"
								write-host "Do you want to enter the advanced options? [y/N]"
								write-host -ForegroundColor Yellow "-->: " -NoNewLine
								$choice = read-host
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") -or ([String]::IsNullOrWhiteSpace($_)) } { break options }
									{ ($_.ToLower() -eq "y") } {
										# Entering additional options and validation functions\main.ps1
										$report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only = Invoke-Command $advanced_params
										break options
									}
									# Exit to the main menu
									m { break main }
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' or 'm' - back to Main menu." }
								}
							}
							while ($true)
							{
								# Checking the computer
								# Reading user input
								write-host -ForegroundColor Yellow "Start checking the ARMs? " -NonewLine
								$choice = read-host "[Y/n] ('o' - show current options)"
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") } { break main }
									{ (($_.ToLower() -eq "y") -or ([String]::IsNullOrWhiteSpace($_))) } {
										# Starting the check functions\scriptblocks.ps1
										$result = Ad
										CheckResult -result $result -path $path -mode $mode
										break main
									}
									# Show additional options functions\main.ps1
									o { ShowOptions -mode $mode -menu $false }
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' - back to Main menu." }
								}
							}
						}
						
						4 {
							$report_only = $true
							$json_only = $false
							# Main menu status
							$menu = $false
							# Report generation mode - multiple computers
							$mode = "multiple"
							# Number of computers in the \src\jsons\ directory
							$file = (Get-ChildItem -Path "$($path)\src\jsons\" -Directory).Name
							if ($file.Count -eq 0)
							{
								write-host -ForegroundColor Red "`n[ERROR] '\src\jsons\' folder is empty!"
								break main
							}
							write-host "`nFound $($file.Count) ARM(s)"
							while ($true)
							{
								write-host -ForegroundColor Yellow "Create reports? " -NonewLine
								$choice = read-host "[Y/n] ('n' - back to main menu)"
								Switch ($choice)
								{
									{ ($_.ToLower() -eq "n") } { break main }
									{ (($_.ToLower() -eq "y") -or ([String]::IsNullOrWhiteSpace($_))) } {
										# Starting the check functions\scriptblocks.ps1
										$result = Multiple
										CheckResult -result $result -path $path -mode $mode
										$report_only = $false
										break main
									}
									
									Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' - back to Main menu." }
								}
							}
						}
						# Show help block functions\main.ps1
						h { ShowHelp; Write-Host; Invoke-Command $main_menu }
						# Show additional options functions\main.ps1
						o {
							$report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only = ShowOptions -menu $true
							Invoke-Command $main_menu
						}
						# Exit the program
						e { return }
						Default { write-host -ForegroundColor Red "`n`n[ERROR] Wrong input!"; Invoke-Command $main_menu }
					}
				}
			}
		}
	}
	catch { }
	if ($Error)
	{
		Write-Host -ForegroundColor Red "`n[ERROR] Please check the error log at 'errors\' for details."
		"$((Get-Date).ToString())`n[InteractiveMode]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
		$Error.clear()
	}
	return
}

# Default script directory
try
{
	$path = $PSCommandPath | Split-Path -Parent
}
catch { $path = $PSScriptRoot }
# IP address pattern for regular expression
$pattern = "(?:25[0-5]|2[0-4][0-9]|[1][0-9]?[0-9]?)\.((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)"

try
{
	# Reading additional parameters
	$advanced_options = Get-Content "$path\src\config.json" | ConvertFrom-Json
	$report_version = $advanced_options.report_version
	if ($report_version -notin ("full", "short"))
	{
		Throw "('full or 'short')"
	}
	$build_number = $advanced_options.build_number
	if ($build_number -notmatch "^[0-9]+$")
	{
		Throw "(Only digits ('17045')"
	}
	$script_threads = [int32]($advanced_options.script_threads)
	if ($script_threads -notin 1 .. 30)
	{
		Throw "(min - 1, max - 30)"
	}
	$arms_threads = [Int32]($advanced_options.arms_threads)
	if ($arms_threads -lt 1)
	{
		Throw "The number of computers must be [Int] and -ge 1"
	}
	$measure_time = [System.Convert]::ToBoolean($advanced_options.measure_time)
	$json_only = [System.Convert]::ToBoolean($advanced_options.json_only)
	$report_only = [System.Convert]::ToBoolean($advanced_options.report_only)
}
catch
{
	Write-Host -ForegroundColor Red "[ERROR] Please check the error log at 'errors\' for details.`n$Error[0]"
	"$((Get-Date).ToString())`n[InteractiveMode]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
	write-host -ForegroundColor Yellow "-->: " -NoNewLine
	$pause = read-host
	return
}

# If arguments are passed to the program, start in non-interactive mode; otherwise, in interactive mode
if ($PSBoundParameters.Count -ge 1)
{
	NonInteractiveMode -arg $PSBoundParameters
}
else { InteractiveMode }
    