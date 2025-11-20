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
