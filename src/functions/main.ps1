# Function to determine user privileges
Function IsAdmin {
    if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        return 1
    }
    return 0
}
# Viewing help
Function ShowHelp
{
	# Array of all possible strings
	$lines = @{
		'0'  = @("PARAMETER", "DESCRIPTION", "DEFAULT VALUE");
		'1'  = @("Required options (only one):", "", "");
		'2'  = @("-c (-computer)", "The computer name", "'$($env:computername)'");
		'3'  = @("-f (-filename)", "The file name or an IP address range", "'none'");
		'4'  = @("-i (-iprange)", "An IP address range", "'none'");
		'5'  = @("-ou (-organizational_unit)", "The OUs", "'none'");
		'6'  = @("Advanced options:", "", "");
		'7'  = @("-rv (-report_version)", "The version of the report", "'$($advanced_options.report_version)'");
		'8'  = @("-bn (-build_number)", "The Operating System build number", "'$($advanced_options.build_number)'");
		'9'  = @("-st (-script_threads)", "The max number (threads) of the ScriptBlocks", "'$($advanced_options.script_threads)'");
		'10' = @("-at (-arms_threads)", "The max number (threads) of the ARMs", "'$($advanced_options.arms_threads)'");
		'11' = @("-mt (-measure_time)", "Measure the program's execution time", "'$(([String]$advanced_options.measure_time).ToLower())'");
        '12' = @("-jo (-json_only)", "Create JSON files without a report", "'$(([String]$advanced_options.json_only).ToLower())'");
        '13' = @("-ro (-report_only)", "Create a report only", "'$(([String]$advanced_options.report_only).ToLower())'");
	}
	
    # Array of column boundaries
    $columns = @{
        '0' = "| ";
        '1' = " | ";
        '2' = " | ";
    }

	# Maximum length for each column
	$max_line_0 = 0
	$max_line_1 = 0
	$max_line_2 = 0
	
	# Determining the maximum length of strings for each column
	foreach ($line in $lines.Values)
	{
		foreach ($i in 0 .. 2)
		{
			New-Variable -Name "len_$i" -Value $line["$i"].length -Force
			if ((Get-Variable -Name "len_$i" -ValueOnly) -gt (Get-Variable -Name "max_line_$i" -ValueOnly))
			{
				New-Variable -Name "max_line_$i" -Value (Get-Variable -Name "len_$i" -ValueOnly) -Force
			}
		}
	}
    # Border (lower/upper)
	$border = $("+" + "-" * ($max_line_0 + 2) + "+" + "-" * ($max_line_1 + 2) + "+" + "-" * ($max_line_2 + 2) + "+")
    # Color
    $border_color = "DarkCyan"
    $text_color = "White"
    # Number of columns
    $columns_count = $columns.Count - 1

	write-host -ForegroundColor Yellow "`nPARAMETERS:"

    # Displaying the table
	foreach ($i in 0 .. 13)
	{
        # Determining text color and displaying the top border
		if ($i -in 0, 1, 6) { 
            $text_color = "DarkCyan"
            if ($i -ne 1) { write-host -ForegroundColor $border_color $border }
        } else { $text_color = "White" }

        # Filling the table
        foreach ($j in 0..$columns_count) {
		    write-host -ForegroundColor $border_color $columns["$j"] -NoNewLine
		    write-host -ForegroundColor $text_color -NoNewLine $($lines["$($i)"][$j] + " " * ((Get-Variable -Name "max_line_$j" -ValueOnly) - $lines["$($i)"][$j].length))
        }
        # Displaying the right and bottom borders
        write-host -ForegroundColor $border_color " |"
        if ($i -in 0, 1, 6) { write-host -ForegroundColor $border_color $border }
	}
	write-host -ForegroundColor DarkCyan $border
    
    # Examples CLI
	write-host -ForegroundColor Yellow "`nEXAMPLES:"
    
    Write-Host -ForegroundColor $border_color "`nDisplay a help message:"
    Write-Host "CheckARM_console.exe -h"

    Write-Host -ForegroundColor $border_color "`nCreate a report only:"
    Write-Host "CheckARM_console.exe -ro"


    Write-host -ForegroundColor $border_color "`n1. Single Mode:"
    Write-Host -ForegroundColor DarkGray "Use the default options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -c PCNAME"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -c 192.168.0.1"
    Write-Host -ForegroundColor DarkGray "`nUse the custom options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -c PCNAME -rv short -bn 19000 -st 4"
    Write-Host -ForegroundColor DarkGray  "`nMeasure the execution time:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -c 192.168.0.1 -mt"
    Write-Host -ForegroundColor DarkGray  "`nCreate JSON files without a report:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -c PCNAME -jo"


    Write-host -ForegroundColor $border_color "`n2. Multiple Mode:"
    Write-Host -ForegroundColor DarkGray "Use the default options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -f C:\Users\User\Desktop\arms.csv"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -i 192.168.0.1-192.168.0.100"
    Write-Host -ForegroundColor DarkGray "`nUse the custom options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt -rv short -bn 19000 -st 5 -at 10"
    Write-Host -ForegroundColor DarkGray  "`nMeasure the execution time:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -i 192.168.0.1-192.168.0.100 -st 3 -mt"
    Write-Host -ForegroundColor DarkGray  "`nCreate JSON files without a report:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -f C:\Users\User\Desktop\arms.txt -jo"

    Write-host -ForegroundColor $border_color "`n3. Active Directory Mode:"
    Write-Host -ForegroundColor DarkGray "Use the default options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -o OU=testou,DC=testdc"
    Write-Host -ForegroundColor DarkGray  "`nUse the custom options:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -o OU=testou,DC=testdc -rv short -bn 19000 -at 2"
    Write-Host -ForegroundColor DarkGray  "`nMeasure the execution time:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -o OU=testou,DC=testdc -mt"
    Write-Host -ForegroundColor DarkGray  "`nCreate JSON files without a report:"
    Write-Host -ForegroundColor $text_color "CheckARM_console.exe -o OU=testou,DC=testdc -jo"
}

# Viewing additional parameters
Function ShowOptions
{
	
	param (
		[Parameter(Mandatory = $false)]
		[ValidateSet("single", "multiple", "ad")]
		[String]$mode,
		
		[Parameter(Mandatory = $false)]
		[ValidateSet($false, $true)]
		[Boolean]$menu
	)
	# Array of all possible strings
	$lines = @{
		'0' = @("DESCRIPTION (* - required)", "VALUE");
		'1' = @("[*] The computer name (default - '$($env:computername)')", "'$($computer)'");
		'2' = @("[*] The file name or an IP address range (default - 'none')", "'$(if (!$range) { 'none' }
				elseif ($range -match '\\') { ($range -split '\\')[-1] }
				else { $range })'");
		'3' = @("[*] The OUs (default - 'none')", "'$(if (!$organizational_unit) { 'none' }
				else { ($organizational_unit -split ',')[0] })'");
		'4' = @("[ ] The version of the report (default - '$($advanced_options.report_version)')", "'$($report_version)'");
		'5' = @("[ ] The Operating System build number (default - '$($advanced_options.build_number)')", "'$($build_number)'");
		'6' = @("[ ] The max number (threads) of the ScriptBlocks (default - '$($advanced_options.script_threads)')", "'$($script_threads)'");
		'7' = @("[ ] The max number (threads) of the ARMs  (default - '$($advanced_options.arms_threads)')", "'$($arms_threads)'");
        '8' = @("[ ] Create JSON files without a report (default - '$($advanced_options.json_only)')", "'$($json_only)'");
		'9' = @("[ ] Measure the program's execution time (default - '$($advanced_options.measure_time)')", "'$($measure_time)'");
	}

    # Array of column boundaries
    $columns = @{
        '0' = "| ";
        '1' = " | ";
    }

	# Maximum length for each column
	$max_line_0 = 0
	$max_line_1 = 0
	
	# Determining the maximum length of strings for each column
	foreach ($line in $lines.Values)
	{
		$len_1 = $line[0].length
		$len_2 = $line[1].length
		if ($len_1 -gt $max_line_0) { $max_line_0 = $len_1 }
		if ($len_2 -gt $max_line_1) { $max_line_1 = $len_2 }
	}
	# Border (lower/upper)
    $border = $("+" + "-" * ($max_line_0 + 2) + "+" + "-" * ($max_line_1 + 2) + "+")
    # Color
    $border_color = "DarkCyan"
    $text_color = "White"

	write-host -ForegroundColor Yellow "`nCurrent options:"
	
	# Arrays of strings depending on the check mode and the place of call

	if ($menu) {
		# If the function is called from the main menu
		$lines_mode = 0, 4, 5, 6, 7, 8, 9
	} else {
		$lines_mode = @{ "single" = 0, 1, 4, 5, 6, 8, 9; "multiple" = 0, 2, 4, 5, 6, 7, 8, 9; "ad" = 0, 3, 4, 5, 6, 7, 8, 9 }[$mode]
	}
	# Number of columns
	$columns_count = $columns.Count - 1
    foreach ($i in $lines_mode) {
		if ($i -eq 0) {
			write-host -ForegroundColor $border_color $border
			$text_color = "DarkCyan"
		} else {
			$text_color = "White"
		}
        foreach ($j in 0..$columns_count) {
            write-host -ForegroundColor $border_color $columns["$j"] -NoNewLine
		    write-host -ForegroundColor $text_color -NoNewLine $($lines["$($i)"][$j] + " " * ((Get-Variable -Name "max_line_$j" -ValueOnly) - $lines["$($i)"][$j].length))
        }
        write-host -ForegroundColor $border_color " |"
        if ($i -in 0, 9) { write-host -ForegroundColor $border_color $border }
	}
	if ($menu) {
		# Block of additional parameters
		:options while ($true)
		{
			# Reading user input
			write-host  "Do you want to change the advanced options? [y/N]"
			write-host -ForegroundColor Yellow "-->: " -NoNewLine
			$choice_ = read-host
			Switch ($choice_)
			{
				{ ($_.ToLower() -eq "n") -or ([String]::IsNullOrWhiteSpace($_)) } {
					return $report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only
				}
				{ ($_.ToLower() -eq "y") } {
					# Entering additional parameters and validation functions\main.ps1
					return Invoke-Command $advanced_params
				}
				# Exit to the main menu
				m { break main }
				Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n' or 'm' - back to Main menu." }
			}
		}
	}
}
# Function for validating the name of the ARM (IP address)
Function ValidateComputer {

    param(
        [Parameter(Mandatory = $true)]
		[String]$computer
    )
    # Remembering the name for return in case of errors
    $std_computername = $computer

    while ($true) {
        try {
            # Reading user input and checking
            write-host -ForegroundColor Yellow "Enter the computer name or IP address " -NoNewLine
            [ValidateScript({if (($_ -match "^$($pattern)$") -or ($_ -match "^[A-Za-z0-9][A-Za-z0-9\-_]{1,14}$") -or ($_ -eq 'm')) {$true} 
            else {Throw "The computer name/IP address is empty or wrong!"}
            })][String]$computer = Read-Host "('m' - back to Main menu)"
            break 
        } catch {
            # Setting default value
            $computer = $std_computername
            write-host -ForegroundColor Red "`n[ERROR] The computer name/IP address is empty or wrong!`n"
        }
    }
    return $computer
}

# Logo
$logo = {
    write-host -ForegroundColor DarkCyan "╔═╗┬ ┬┌─┐┌─┐┬┌─  ╔═╗╦═╗╔╦╗"
    write-host -ForegroundColor DarkCyan "║  ├─┤├┤ │  ├┴┐  ╠═╣╠╦╝║║║"
    write-host -ForegroundColor DarkCyan "╚═╝┴ ┴└─┘└─┘┴ ┴  ╩ ╩╩╚═╩ ╩"
    write-host -ForegroundColor DarkCyan "console version (@kronigor)"
}

# Main menu 
$main_menu = {
    Write-host -ForegroundColor DarkCyan "------------------------------------------------" 
    Write-host -ForegroundColor Yellow "Please, select the operation mode:`n"
    Write-host "[1] Single mode (one ARM)"
    Write-host "[2] Multiple mode (several ARMs)"
	Write-host "[3] Active Directory mode (scan OUs)"
    Write-host "[4] Create a report (from \src\jsons\)"
	Write-host "[o] Change options"
	Write-host "[h] Show help"
    Write-host -ForegroundColor DarkCyan "------------------------------------------------"
    Write-host "[e] Exit"
}

# Additional parameters (for single mode)
$advanced_params = {
         
    while ($true) {
        try {
            # Reading user input and checking
            write-host -ForegroundColor Yellow "Enter the version report " -NonewLine
            [ValidateScript({if (([String]::IsNullOrWhiteSpace($_)) -or ($_.ToLower() -in  ("full", "short", "m"))) {$true}
                                else {Throw "('full, 'short', 'm' - back to Main menu)"}
            })][String]$report_inpt = Read-Host "('m' - back to Main menu)"
            # Setting default value and exiting to the menu
			if ($report_inpt.ToLower() -eq "m") { break main }
			elseif ([String]::IsNullOrWhiteSpace($report_inpt)) { break }
			else { $report_version = $report_inpt; break }
		}
		catch
		{
			write-host -ForegroundColor Red "`n[ERROR] The version report is empty or wrong! ('full', 'short', 'm' - back to Main menu)`n"
		}
	}
	
	while ($true)
	{
		try
		{
			# Reading user input and checking
			write-host -ForegroundColor Yellow "Enter an Operating System build number " -NonewLine
			[ValidateScript({
					if (($_.ToLower() -eq "m") -or ($_ -match "^[0-9]+$|^m$") -or ([String]::IsNullOrWhiteSpace($_))) { $true }
					else { Throw "(Only digits ('17045') ('m' - back to Main menu)" }
				})][String]$build_number_inpt = Read-Host "('m' - back to Main menu)"
			
			# Setting default value and exiting to the menu
			if ($build_number_inpt.ToLower() -eq "m") { break main }
			elseif ([String]::IsNullOrWhiteSpace($build_number_inpt)) { break }
			else { $build_number = $build_number_inpt; break }
		}
		catch
		{
			write-host -ForegroundColor Red "`n[ERROR]  The Operating System build number is empty or wrong! ('17045')`n"
		}
	}
	
	while ($true)
	{
		try
		{
			# Reading user input and checking
			write-host -ForegroundColor Yellow "Enter a number of running scriptblocks in parallel " -NonewLine
			[ValidateScript({
					if (($_.ToLower() -eq "m") -or ($_ -in 1 .. 30) -or ([String]::IsNullOrWhiteSpace($_))) { $true }
					else { Throw "(min - 1, max - 30, 'm' - back to Main menu)" }
				})][String]$script_threads_inpt = Read-Host "(min - 1, max - 30, 'm' - back to Main menu)"
			# Setting default value and exiting to the menu
			if ($script_threads_inpt.ToLower() -eq "m") { break main }
			elseif ([String]::IsNullOrWhiteSpace($script_threads_inpt)) { break }
			else { $script_threads = $script_threads_inpt; break }
		}
		catch
		{
			write-host -ForegroundColor Red "`n[ERROR]  The number of running scriptblocks is empty or wrong (min - 1, max - 30, 'm' - back to Main menu)!`n"
		}
	}
	
	if ($mode -ne "single" -or $menu) {
        while ($true) {
            try {
                # Reading user input and checking
                write-host -ForegroundColor Yellow "Enter a number of computers checked at the same time (in parallel) " -NonewLine
                [ValidateScript({if (($_.ToLower() -eq "m") -or ([String]::IsNullOrWhiteSpace($_)) -or ([Int32]$_ -ge 1)) {$true}
                                else {Throw "The number of computers must be [Int] and -ge 1"}
                                })][String]$arms_threads_inpt = Read-Host "('m' - back to Main menu)"
                # Setting default value and exiting to the menu                        
				if ($arms_threads_inpt.ToLower() -eq "m") { break main }
				elseif ([String]::IsNullOrWhiteSpace($arms_threads_inpt)) { break }
				else { $arms_threads = $arms_threads_inpt; break }
			}
			catch
			{
				write-host -ForegroundColor Red "`n[ERROR]  The number of computers is empty or wrong (must be [Int] and -ge 1)!`n"
			}
		}
	}

    while ($true)
	{
		try
		{
			# Reading user input and checking
			write-host -ForegroundColor Yellow "Do you want to create JSON files without a report? " -NonewLine
			[ValidateScript({
					if ($_.ToLower() -in @("y", "n", "m") -or [String]::IsNullOrWhiteSpace($_) -or $_ -in @($false, $true)) { $true }
				})][String]$json_only_inpt = Read-Host "[y/n] ('m' - back to Main menu)"
			
			# Setting default value and exiting to the menu
			if ($json_only_inpt.ToLower() -eq "m") { break main}
			elseif ([String]::IsNullOrWhiteSpace($json_only_inpt)) { break }
			elseif ($json_only_inpt.ToLower() -eq "n" -or $json_only_inpt.ToLower() -eq "false") { $json_only = $false; break }
			elseif ($json_only_inpt.ToLower() -eq "y" -or $json_only_inpt.ToLower() -eq "true") { $json_only = $true;  break }
			else {
				write-host -ForegroundColor Red "`n[ERROR] Input is wrong! (Must be 'y', 'n', 'm', 'false', 'true')!`n"
			}
			
		}
		catch
		{
			write-host -ForegroundColor Red "`n[ERROR] Input is wrong! ('y', 'n', 'm', 'false', 'true')!`n"
		}
	}
	
	while ($true)
	{
		try
		{
			# Reading user input and checking
			write-host -ForegroundColor Yellow "Do you want to measure execution time of the program? " -NonewLine
			[ValidateScript({
					if ($_.ToLower() -in @("y", "n", "m") -or [String]::IsNullOrWhiteSpace($_) -or $_ -in @($false, $true)) { $true }
				})][String]$measure_time_inpt = Read-Host "[y/n] ('m' - back to Main menu)"
			
			# Setting default value and exiting to the menu
			if ($measure_time_inpt.ToLower() -eq "m") { break main}
			elseif ([String]::IsNullOrWhiteSpace($measure_time_inpt)) { break }
			elseif ($measure_time_inpt.ToLower() -eq "n" -or $measure_time_inpt.ToLower() -eq "false") { $measure_time = $false; break }
			elseif ($measure_time_inpt.ToLower() -eq "y" -or $measure_time_inpt.ToLower() -eq "true") { $measure_time = $true;  break }
			else {
				write-host -ForegroundColor Red "`n[ERROR] Input is wrong! (Must be 'y', 'n', 'm', 'false', 'true')!`n"
			}
			
		}
		catch
		{
			write-host -ForegroundColor Red "`n[ERROR] Input is wrong! ('y', 'n', 'm', 'false', 'true')!`n"
		}
	}
	
	return $report_version, $build_number, $script_threads, $arms_threads, $measure_time, $json_only
}

# Checking results and displaying the report/directory
Function CheckResult {

    param(
        [Parameter(Mandatory = $true)]
		[String]$result,
        [Parameter(Mandatory = $true)]
		[String]$path,
        [Parameter(Mandatory = $true)]
		[String]$mode
    )
    while ($true) {
        if ($mode -eq "single") {
            if($result -in "Errors","jsons") { write-host -ForegroundColor Yellow "`nOpen a '$($result)' folder? " -NonewLine } 
            else { write-host -ForegroundColor Yellow "`nOpen a report? " -NonewLine }
            
            $choice =  read-host "[y/N]"
            Switch ($choice) {
                {($_.ToLower() -eq "n") -or  ([String]::IsNullOrWhiteSpace($_))} { return }
                {($_.ToLower() -eq "y")} {
                    # Entering additional parameters and validation functions\main.ps1
                    if ($result -eq "Errors") { Invoke-Item "$($path)\errors\" } 
                    elseif ($result -eq "json") { Invoke-Item "$($path)\src\jsons\$($computer)" }
                    else { Start-Process $result }
                    return
            }
                Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n'."}
            } 
        } else {
            write-host -ForegroundColor Yellow "`nOpen a '$($result)' folder? " -NonewLine
            $choice =  read-host "[y/N]"
            Switch ($choice) {
                {($_.ToLower() -eq "n") -or  ([String]::IsNullOrWhiteSpace($_))} { return }
                {($_.ToLower() -eq "y")} {
                    # Entering additional parameters and validation functions\main.ps1
                    if ($result -eq "reports") { Invoke-Item "$($path)\reports\" } 
                    elseif ($result -eq "jsons") { Invoke-Item "$($path)\src\jsons\" }
                    else { Invoke-Item "$($path)\errors\" }
                    return
            }
                Default { write-host -ForegroundColor Red "`n[ERROR] Wrong input! Please, enter 'y' or 'n'."}
            } 
        }
    }
}