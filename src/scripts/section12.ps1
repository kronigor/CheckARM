$tasks_dict = @{ } # Dictionary for collecting tasks from the Scheduler
$tasks = Get-ScheduledTask | Where-Object { $_.TaskPath -notmatch 'Microsoft' } | Select-Object TaskName, State, Author # Retrieving data from the scheduler
foreach ($task in $tasks) # Iterating through all tasks
{
	$tasks_dict[$task.TaskName] = @{ 'State' = [string]$task.State; 'Author' = $task.Author } # Adding name, status, and author to the dictionary
}

return $tasks_dict | ConvertTo-Json