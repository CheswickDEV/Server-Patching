# Server-Patching

Dieses Repository enthält Hilfsskripte für Server-Patching-Aufgaben. Verwenden Sie `install-dotnet48.ps1`, um die .NET Framework 4.8-Installation auf einem Windows-Server automatisiert auszuführen.

## .NET Framework 4.8 per PowerShell installieren

1. Öffnen Sie eine administrative PowerShell-Sitzung auf dem Zielserver.
2. Laden Sie das Skript herunter oder kopieren Sie es auf den Server.
3. Stellen Sie sicher, dass der Pfad zu Ihrer vorhandenen Offline-Installer-EXE (z. B. ein UNC-Pfad zu einem Share) bekannt ist.
4. Führen Sie das Skript aus:

   ```powershell
   .\install-dotnet48.ps1 -InstallerPath \\fileserver\pfad\zu\ndp48-x86-x64-allos-enu.exe
   ```

   Das Skript erwartet einen vorhandenen Offline-Installer (kein Download). Es führt ihn im Quiet-Mode aus (`/q /norestart`) und schreibt ein Log nach `%TEMP%\dotnet48-install.log`.

### Optionale Parameter

- `-InstallerPath <Pfad>` (erforderlich): Pfad zu einer vorhandenen Offline-Installer-EXE, z. B. von einem Netzwerk-Share.
- `-LogPath <Pfad>`: Ändert den Speicherort des Installationslogs.

### Ablauf des Skripts

- Prüft, ob die PowerShell-Sitzung mit Administratorrechten läuft.
- Überspringt die Installation, wenn .NET Framework 4.8 (Release-Key ≥ 528040) bereits vorhanden ist.
- Führt den bereitgestellten Offline-Installer mit `Start-Process -Wait` aus.
- Behandelt die Rückgabecodes `0` (erfolgreich) und `3010` (Neustart erforderlich) als erfolgreiche Installation, vermerkt den erforderlichen Neustart im Log, führt ihn jedoch **nicht** durch.
- Validiert nach Abschluss erneut die installierte .NET-Version.
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
