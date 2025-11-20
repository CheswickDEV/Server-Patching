# Server-Patching

Dieses Repository stellt ein PowerShell-Skript bereit, das Windows Server in einer Domäne automatisiert patcht.

## Skript: `Patch-WindowsServers.ps1`

### Funktionsumfang
- Anstoßen des Patchings mehrerer Server von einem zentralen System aus
- Automatisches Ermitteln und Installieren der aktuell lokal/verfügbaren Windows-Updates (inkl. kumulativer KBs)
- Nutzung der integrierten Windows Update Agent API über einen temporären SYSTEM-Scheduled-Task (kein externes PowerShell-Modul erforderlich)
- Neustartkontrolle mit Rückmeldung, ob Server nach dem Patchen wieder online sind
- Vorher/Nachher-Vergleich von OS-Version und Build
- Direktes Live-Reporting im Terminal sowie Export als CSV-Report
- Entspricht funktional dem manuellen Klick auf „Check for updates“ bzw. „Install now“ in den Windows-Update-Einstellungen

### Voraussetzungen
- PowerShell 5.1 (oder neuer) auf dem Management-Server
- Aktiviertes PowerShell Remoting (WinRM) auf allen Zielsystemen
- Domain- oder lokales Konto mit Administratorrechten auf den Zielservern
- Kein Internetzugriff notwendig – Updates können vorab über WSUS oder manuell auf die Zielsysteme gebracht werden
- Die Zielserver müssen das Anlegen und Ausführen geplanter Aufgaben (Aufgabenplanung) erlauben; das Skript erstellt temporär einen Task unter `C:\ProgramData\Remote-Patching`
- Unterstützte Zielsysteme: Windows Server 2016, 2019 und 2022

### Beispielaufruf
```powershell
# Liste der zu patchenden Server definieren
$servers = "APP01","DB01"

# Skript mit administrativen Rechten ausführen
.
\Patch-WindowsServers.ps1 -Servers $servers
```

Nach der Ausführung befindet sich der Report als CSV-Datei im aktuellen Verzeichnis.
