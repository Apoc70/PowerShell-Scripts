# Ermitteln der aktuellen Größe des Arbeitsspeichers
$TotalMemory = (Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory
$TotalMemoryMB = [math]::Round($TotalMemory / 1MB)

# Berechnen der Größe der Auslagerungsdatei (25% des Arbeitsspeichers)
$PageFileSizeMB = [math]::Round($TotalMemoryMB * 0.25)

# Ausgabe der ermittelten Werte
Write-Output "Gesamter Arbeitsspeicher: $TotalMemoryMB MB"
Write-Output "Größe der Auslagerungsdatei (25% des Arbeitsspeichers): $PageFileSizeMB MB"

# Deaktivieren der automatischen Verwaltung der Auslagerungsdatei
$computersys = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
$computersys.AutomaticManagedPagefile = $False
$computersys.Put()

# Festlegen der Größe der Auslagerungsdatei auf 25% des Arbeitsspeichers
Set-CimInstance -Query "SELECT * FROM Win32_PageFileSetting" -Property @{
    InitialSize = $PageFileSizeMB
    MaximumSize = $PageFileSizeMB
}

# Bestätigung der neuen Einstellungen
Write-Output "Die Auslagerungsdatei wurde erfolgreich auf $PageFileSizeMB MB konfiguriert."
Write-Output "Bitte starte den Computer neu, um die Änderungen zu übernehmen."
