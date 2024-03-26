# Function to retrieve ARM settings (Multiple)
$GetSettingsMultiple = {
	Param (
		[Parameter(Mandatory = $true)]
		[String]$computer,
		[Parameter(Mandatory = $true)]
		[Array]$sections,
        [Parameter(Mandatory = $true)]
		[String]$build_number,
		[Parameter(Mandatory = $true)]
		[Int]$script_threads,
        [Parameter(Mandatory = $true)]
		[String]$path,
        [Parameter(Mandatory = $true)]
		[Boolean]$report_only
        
	)
    try {
		$ErrorActionPreference = "SilentlyContinue"
        # Connecting functions
	    . $path\src\functions\scriptblocks.ps1
        # Operation mode - winrm
        $local = 0
        # Array of blocks
	    $ScriptBlocks = @{
		    1 = $Section1; 2 = $Section2; 3 = $Section3;
		    4 = $Section4; 5 = $Section5; 6 = $Section6;
		    8 = $Section8; 9 = $Section9; 10 = $Section10;
		    11 = $Section11; 12 = $Section12; 13 = $Section13;
		    14 = $Section14; 16 = $Section16
	    }
	    # Checking ARM availability
        $testconn, $computer = TestConn -computer $computer
        if ($testconn) {
            # If the ARM is pingable, but there are no rights for wmi 
            $testinvoke = TestInvoke -computer $computer
            if (!$testinvoke) {
                # Array is needed to determine errors in multithreading  
                $out = @('ErrorCheck', $computer)
		        return $out
            }
        } else {
            $out = @('ErrorCheck', $computer)
		    return $out
        } 

        # Checking the directory for dumping
        try { CheckDirectory -computer $computer -report_only $report_only -ErrorAction Stop  } catch { 
            $out = @('ErrorCheck', $computer)
		    "$((Get-Date).ToString())`n[CheckDirectory]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
		    $Error.clear()
		    return $out
		}
				
		# Task pool (multithreading)
	    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $script_threads) 
	    $RunspacePool.Open() 
        # Array of tasks   	    
        $Jobs = @()
	    # Task generator
	    foreach ($i in $sections) {
		    $PowerShell = [powershell]::Create()
		    $PowerShell.RunspacePool = $RunspacePool
            # Adding a scriptblock and passing arguments
		    $PowerShell.AddScript($ScriptBlocks[$i]).AddArgument($computer).AddArgument($path).AddArgument($local)
		    $Jobs += $PowerShell.BeginInvoke()
	    }
        # Monitoring the status of tasks and displaying progress
	    while ($Jobs.IsCompleted -contains $false) { Start-Sleep -Milliseconds 100 }
	    # Closing the Pool
	    $RunspacePool.Close() 
	    return
    } catch {}
    if ($Error) {
        "$((Get-Date).ToString())`n[GetSettingsMultiple]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
		$Error.clear()
		$out = @('ErrorCheck', $computer)
		return $out
    }
}

# Function for checking multiple ARMs (Multiple/AD)
$CheckMultiple = {
	
    Param (
		[Parameter(Mandatory = $true)]
		[Array]$file,
        [Parameter(Mandatory = $true)]
        [Int32]$arms_threads,
        [Parameter(Mandatory = $true)]
		[String]$report_version,
        [Parameter(Mandatory = $true)]
		[String]$build_number,
		[Parameter(Mandatory = $true)]
		[Int32]$script_threads,
        [Parameter(Mandatory = $true)]
		[String]$path,
		[Parameter(Mandatory = $true)]
		[Boolean]$measure_time,
        [Parameter(Mandatory = $true)]
		[Boolean]$json_only,
        [Parameter(Mandatory = $true)]
		[Boolean]$report_only

	)
	# Timer
	$watch = [System.Diagnostics.Stopwatch]::StartNew()
	$watch.Start()
	$ErrorActionPreference = "SilentlyContinue"
    # Checking for standard directories
    try {
        DefaultDirectories -computer $computer -path $path -report_only $report_only -ErrorAction Stop
    } catch {
        write-host -ForegroundColor Red "[ERROR] An error occured [DefaultDirectories]!"
        return 
    }
    # Number of exported blocks depending on the report version
    if ($report_version -eq 'full') { $sections = @(1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 16)
    } else { $sections = @(1, 2, 3, 4, 5, 6, 8, 9, 10, 16) }
	
	try {
		
		# Disabling the console cursor (for correct progress bar display)
		if (!(Test-Path variable:global:psISE)) { [Console]::CursorVisible = $false }
		# Task pool (multithreading)
		$RunspacePoolGroup = [runspacefactory]::CreateRunspacePool(1, $arms_threads)
		$RunspacePoolGroup.Open()
        # Array for storing ARMs with errors
		$Results = @()
        # Two phases: 1. Dumping JSON from ARMs; 2. Generating reports
        
        if ($json_only) { $arr = 0 }
		elseif ($report_only) { Write-Host ""; $arr = 1 }
		else { $arr = @(0, 1) }
		
		foreach ($j in $arr)
		{
            # Removing empty lines
            $file = $file.Where({![string]::IsNullOrWhiteSpace($_)})
            # Progress of phase completion
            $percent = 0
            # Number of completed ARMs
            $count = 0
            # Array of tasks
			$AllJobs = New-Object System.Collections.ArrayList
            # If it's the 1st phase
			if (!$j) { 
                $status = "Collecting an information from the arms:"
                write-host -ForegroundColor Yellow "`n$($status)" 
                $count_length = ($file | Measure-Object).Count.tostring().length 
            }
            # If it's the 2nd phase
            else {
                # Getting names of ARMs from which JSONs were dumped
				$file = (Get-ChildItem -Directory "$path\src\jsons\").Name
                if ($file) { 
                    $status = "Creating the reports:"
					write-host -ForegroundColor Yellow $status
                    if (!$count_length){
                        $count_length = 0
                    }
                }
			}
            # Total number of ARMs (at each stage)
            $count_arms = ($file | Measure-Object).Count
            # Different progress bars for ISE and console
            if (Test-Path variable:global:psISE) {
                Write-Progress -Activity "Progress" -PercentComplete $percent -Status "$($status) [$($count)/$($count_arms)]"
            } else {
                # If there are accessible ARMs
                if ($count_arms) {
                    $spaces_1 = $(' ' * ((@($count_length, $count_arms.tostring().length) | Measure-Object -Maximum).Maximum - $count.tostring().length))
                    $spaces_2 = $(' ' * ((@($count_length, $count_arms.tostring().length) | Measure-Object -Maximum).Maximum - $count_arms.tostring().length))
					write-host -BackgroundColor DarkCyan -ForegroundColor Black "Progress: [$($spaces_1)$($count)/$($count_arms)$($spaces_2)]" -NoNewline
                    write-host -NoNewLine " [$('#' * ($percent / 2) + "." * ((100 - $percent) / 2))]"
                }
            }
            # Going through all ARMs
			foreach ($computer in $file) {
				# Removing spaces
				$computer = $computer.Trim()
                # Creating a task
				$asyncJob = [powershell]::Create()
                # Adding to the Pool
				$asyncJob.RunspacePool = $RunspacePoolGroup

				if (!$j) {
                    # Adding a scriptblock and arguments to the task (1st phase)
					$asyncJob.AddScript($GetSettingsMultiple).AddArgument($computer).AddArgument($sections).AddArgument($build_number).AddArgument($script_threads).AddArgument($path).AddArgument($report_only)
                } else {
                    # Checking the JSON directory
                    try { CheckDirectory -computer $computer -report_only $report_only -ErrorAction Stop } catch { 
                        $Results += $computer
		                "$((Get-Date).ToString())`n[CheckDirectory]`n$($Error[0])" >> "$($path)\errors\$($computer).txt"
		                $Error.clear()
                        continue
		            }
                    # Adding a scriptblock and arguments to the task (2nd phase)
					$asyncJob.AddScript($Section15).AddArgument($computer).AddArgument($path).AddArgument($report_version).AddArgument($build_number)
				}
                # Creating an object
				$asyncJobObj = @{
					JobHandle   = $asyncJob;
					AsyncHandle = $asyncJob.BeginInvoke()
				}
                # Adding the object to the array
				$AllJobs.Add($asyncJobObj) | Out-Null	
            }
            # Execution status
			$ProcessingJobs = $true
			Do
			{
                # Getting completed tasks   
				$CompletedJobs = $AllJobs | Where-Object { $_.AsyncHandle.IsCompleted }
				if ($null -ne $CompletedJobs) {
                    # Going through tasks
					foreach ($job in $CompletedJobs) {
						$result = $job.JobHandle.EndInvoke($job.AsyncHandle)
                        # If a task is completed with an error, add the ARM's name to the array
						if ($result -like 'ErrorCheck') { $Results += $result[-1] }
						$job.JobHandle.Dispose()
                        # Removing the task from the array
						$AllJobs.Remove($job)
						$count += 1
                        # Displaying progress
                        $percent = [int]($count * 100 / $count_arms)
                        # Different progress bars for ISE and console
                        if (Test-Path variable:global:psISE){
                            Write-Progress -Activity "Progress" -PercentComplete $percent -Status "$($status) [$($count)/$($count_arms)]"
                        } else {
                            $spaces_1 = $(' ' * ((@($count_length, $count_arms.tostring().length) | Measure-Object -Maximum).Maximum - $count.tostring().length))
                            $spaces_2 = $(' ' * ((@($count_length, $count_arms.tostring().length) | Measure-Object -Maximum).Maximum - $count_arms.tostring().length))
					        write-host -BackgroundColor DarkCyan -ForegroundColor Black "`rProgress: [$($spaces_1)$($count)/$($count_arms)$($spaces_2)]" -NoNewline
                            write-host -NoNewLine " [$('#' * ($percent / 2) + "." * ((100 - $percent) / 2))]"
                        }
					}
				} else {
                    # If the array is empty, changing the execution status (completed)
					if ($AllJobs.Count -eq 0) { $ProcessingJobs = $false }
					else { Start-Sleep -Milliseconds 100 }
				}
			}
			While ($ProcessingJobs)
            # If using the console
            if ($count_arms -and !(Test-Path variable:global:psISE)){ write-host }
            # If using ISE
            if (Test-Path variable:global:psISE) {
                Write-Progress -Activity "Progress" -PercentComplete 100 -Completed
                if ($count_arms -notin @($Results.Count, 0)) { Write-host -ForegroundColor DarkCyan 'DONE!' }
            }
		}
        # Closing the Pool
		$RunspacePoolGroup.Close()
		$RunspacePoolGroup.Dispose()
        # Dumping ARMs with errors to a file
		$Results >> "$path\unchecked_arm.txt"
		Start-Sleep -Milliseconds 500
		# Enabling the console cursor
		if (!(Test-Path variable:global:psISE)) { [Console]::CursorVisible = $true }
		# Stopping the timer
		$watch.Stop()
        if ($count_arms) {
		    write-host -ForegroundColor Green "`n[SUCCESS]" -NoNewline
            if ($json_only) {
                write-host  " Check 'src\jsons\' for exported JSON data."
                $out = 'jsons'
            } else { 
                write-host  " Check 'reports\' for the generated report." 
                $out = 'reports'
            }
			if ($measure_time) { Write-Host -ForegroundColor DarkGray "[TIME] $($watch.Elapsed)" }
			return $out
        } else {
			write-host -ForegroundColor Red "`n[ERROR] The computers are unavailable. Check 'errors\'"
			if ($measure_time) { Write-Host -ForegroundColor DarkGray "[TIME] $($watch.Elapsed)" }
			return "Errors" 
        }
	
	} catch {
        write-host -ForegroundColor Red "`n[ERROR] Please check the error log at 'errors\' for details."
		"$((Get-Date).ToString())`n[CheckMultiple]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
		return "Errors" 
	}
    if ($Error) {
        write-host -ForegroundColor Red "`n[ERROR] Please check the error log at 'errors\' for details."
        "$((Get-Date).ToString())`n[CheckMultiple]`n$($Error[0])" >> "$path\errors\MAIN_ERRORS.txt"
        $Error.clear()
        return "Errors" 
    }
}