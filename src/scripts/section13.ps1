$system_EventLog = @{}
$start_day = (Get-Date).AddDays(-60)
$end_day = Get-Date # Retrieving the current date
try { 
    $system_logs = Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=$start_day; EndTime=$end_day; Level=1, 2, 3} | Select-Object Id, TimeCreated, ProviderName, LevelDisplayName, RecordId, Message
} catch {
    continue
}
# Iterating through all errors
foreach ($log in $system_logs)
{
    # If the error is not in the dictionary
	if ([string]$log.Id -notin $system_EventLog)
	{
		# If the message is longer than 151 characters
		if ($log.Message) {
	        if ($log.Message.Length -gt 151) {
		        # Truncating the message
		        $log.Message = ($log.Message).Substring(0, 150) + '...' 
	        }
        } else { $log.Message = "\n"  }
        if (!$log.LevelDisplayName) {
            $log.LevelDisplayName = 'none'
        }
        try { 
		    $system_EventLog[[string]$log.Id] = @{ 'Time' = ($log.TimeCreated).ToString(); 'ProviderName' = $log.ProviderName; 'Level' = ($log.LevelDisplayName).ToString(); 'RecordId' = $log.RecordId; 'Message' = $log.Message.Replace('\n', ' ') } # Adding the error and parameters with a description
        } catch { continue }
    }

} 

$security_EventLog = @{}

$EventId = 4609, 4625, 4697, 4698, 4719, 4720, 4722, 4723, 4724, 4725, 4726, 4728, 4729, 4731, 4732, 4733, 4734, 4735, 4738, 4740, 4776
try { 
    $security_logs = Get-WinEvent -FilterHashtable @{LogName='Security'; StartTime=$start_day; EndTime=$end_day; ID=$EventId} | Select-Object RecordId, TimeCreated, Id, Message
} catch {
    continue
}
# Iterating through all errors
foreach ($log in $security_logs)
{
    # If the message contains a process name
    if ($log.Message -match "Имя процесса:\t\t(.+)\r") {
        $log | Add-Member -NotePropertyName Process -NotePropertyValue $matches[1]
    } else {
        $log | Add-Member -NotePropertyName Process -NotePropertyValue ""
    }
    # If the message is longer than 200 characters
    if ($log.Message) {
	    if ($log.Message.Length -gt 201) {
		    # Truncating the message
		    $log.Message = ($log.Message).Substring(0, 200) + '...' 
	    }
    } else { $log.Message = "\n"  }
    try {  		
	    $security_EventLog[[string]$log.RecordId] = @{'Time' = ($log.TimeCreated).ToString(); 'Id' = $log.Id; 'Message' = $log.Message.Replace('\n', ' ')}
    } catch { continue }
}

if (!$system_logs) {
    return New-Object -TypeName PSCustomObject -Property @{security_EventLog=$security_EventLog | ConvertTo-Json}
}
if (!$security_logs) {return New-Object -TypeName PSCustomObject -Property @{system_EventLog=$system_EventLog  | ConvertTo-Json}}

return New-Object -TypeName PSCustomObject -Property @{system_EventLog=$system_EventLog  | ConvertTo-Json
                                                       security_EventLog=$security_EventLog | ConvertTo-Json
                                                       }