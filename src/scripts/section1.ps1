return Get-WmiObject -class win32_computerSystem | Select-Object Name, Domain, UserName | ConvertTo-Json  
