$services = gwmi win32_service | Select-Object Name, Startmode, State, DisplayName, PathName | Sort-Object -Property DisplayName   # Exporting the list of services 
foreach ($service in $services) {
    if ($service.Name -match "(\w+)_") {
        $service.Name = $matches[1]
    }
}
return $services | ConvertTo-Json