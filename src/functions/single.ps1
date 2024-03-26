# Function to obtain ARM settings (Single)
$GetSettingsSingle = {
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
        [Parameter(Mandatory = $true)]
		[String]$report_version,
        [Parameter(Mandatory = $true)]
		[String]$build_number,
		[Parameter(Mandatory = $true)]
		[Int]$script_threads,
		[Parameter(Mandatory = $true)]
		[Int]$local,
		[Parameter(Mandatory = $true)]
		[Boolean]$measure_time,
        [Parameter(Mandatory = $true)]
		[Boolean]$json_only,
        [Parameter(Mandatory = $true)]
		[Boolean]$report_only
	)
    
    try {
		# Timer
		$watch = [System.Diagnostics.Stopwatch]::StartNew()
		$watch.Start()
		$ErrorActionPreference = "SilentlyContinue"
        
        # Number of blocks exported depending on the report version
        if ($report_version -eq 'full') { $sections = @(1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 16)
        } else { $sections = @(1, 2, 3, 4, 5, 6, 8, 9, 10, 16) }	
		# Array of blocks
	    $ScriptBlocks = @{
		    1 = $Section1; 2 = $Section2; 3 = $Section3;
		    4 = $Section4; 5 = $Section5; 6 = $Section6;
		    8 = $Section8; 9 = $Section9; 10 = $Section10;
		    11 = $Section11; 12 = $Section12; 13 = $Section13;
		    14 = $Section14; 16 = $Section16
	    }
        # Checking for standard directories
	    try {
            DefaultDirectories -computer $computer -path $path -report_only $report_only -ErrorAction Stop
        } catch {
            write-host -ForegroundColor Red "`n[ERROR] An error occurred [DefaultDirectories]!"
            return "Errors" 
        }
        # Checking the JSON directory
	    try {
		    CheckDirectory -computer $computer -report_only $report_only -ErrorAction Stop
	    } catch {
            Write-Host -ForegroundColor Red "`n[ERROR] Please check the error log at 'errors\' for details."
            "$((Get-Date).ToString())`n[CheckDirectory]`n$($Error[0])" >> "$($path)\errors\$($computer).txt" 
            return "Errors" 
        }
        if (!$report_only) {
	        # Number of items for the progress bar
	        $sectionsCount = $sections.Count + [Int](!$json_only)
            # Pool for tasks (multithreading)
	        $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $script_threads) 
	        $RunspacePool.Open()
            # Array of tasks
	        $Jobs = @()
	    
            # Task generator
	        foreach ($i in $sections)
	        {
		        $PowerShell = [powershell]::Create()
		        $PowerShell.RunspacePool = $RunspacePool
                # Adding a script block and passing arguments
		        $PowerShell.AddScript($ScriptBlocks[$i]).AddArgument($computer).AddArgument($path).AddArgument($local)
		        $Jobs += $PowerShell.BeginInvoke()
		    }
		    # Disabling the cursor in the console (correct display of the progress bar)
		    if (!(Test-Path variable:global:psISE)) { [Console]::CursorVisible = $false }
		    Write-Host -ForegroundColor Yellow "`nChecking '$($computer)':"
            # Monitoring the status of tasks and displaying progress
	        while ($Jobs.IsCompleted -contains $false)
	        {
		        Start-Sleep -Milliseconds 10
		        $count = ($Jobs.IsCompleted | Group-Object | Where-Object { $_.Name -eq $true }).count
                $percent = [int]($count * 100 / $sectionsCount)
                # Different progress bars for ISE and console
                if (Test-Path variable:global:psISE){
                    Write-Progress -Activity "Checking '$($computer)'" -PercentComplete $percent -Status "Collecting information from the ARM:"
                } else {
				    write-host -BackgroundColor DarkCyan -ForegroundColor Black "`rProgress: [$(' ' * (3 - $percent.tostring().length))$percent%]" -NoNewline
                    write-host -NoNewLine " [$('#' * ($percent / 2) + "." * ((100 - $percent) / 2))]"
                }
            
	        }
        
	        # Closing the Pool
	        $RunspacePool.Close() 
        }
        if (!$json_only) {
            # Function for generating the report and recording the result in a variable
	        try
			{
				$path_report = $Section15.Invoke($computer, $path, $report_version, $build_number)
		        $Error.clear()
		        # If an error occurred while executing the PYTHON script
	            if ($path_report -eq 'Error') {
                    write-host
                    write-host -ForegroundColor Red "[ERROR] An error occurred while generating the report!"
                    return "Errors" 
	            }
				
				$name_report = ($path_report -split "\\")[-1]
                $percent = 100
                if (Test-Path variable:global:psISE){
                    Write-Progress -Activity "Checking '$($computer)'" -PercentComplete $percent -Status "Creating a report:"
                    Write-Progress -Activity "Checking '$($computer)'" -Completed
                } else {
				    write-host -BackgroundColor DarkCyan -ForegroundColor Black "`rProgress: [$(' ' * (3 - $percent.tostring().length))$percent%]" -NoNewline
                    write-host " [$('#' * ($percent / 2))]"
                }
                write-host -ForegroundColor Green "[SUCCESS]" -NoNewline
                write-host " See the report: \reports\$($name_report)"
		    } catch {}
        }
        else {
            write-host -ForegroundColor Green "`n`n[SUCCESS]" -NoNewline
            write-host " See JSON files: \src\jsons\$($computer)"
            $path_report = "jsons"
        }
		# Enabling the cursor in the console
		if (!(Test-Path variable:global:psISE)) { [Console]::CursorVisible = $true }
		# Stopping the timer
		$watch.Stop()
		if ($measure_time) { Write-Host -ForegroundColor DarkGray "[TIME] $($watch.Elapsed)" }
		return $path_report
    } catch {}
    if ($Error) {
        write-host -ForegroundColor Red "[ERROR] Please check the error log at 'errors\' for details."
        "$((Get-Date).ToString())`n[GetSettingsSingle]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
        $Error.clear()
        return "Errors" 
    }
}