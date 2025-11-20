<#
.SYNOPSIS
    Automatisiert das Einspielen von Windows Updates auf einer Liste von Servern.

.DESCRIPTION
    Dieses Skript wird von einem zentralen Administrationsserver gestartet. Es stellt eine Verbindung
    zu jedem angegebenen Zielserver her, ermittelt über Windows Update die verfügbaren Patches,
    installiert diese automatisiert und löst bei Bedarf einen Neustart aus. Im Anschluss werden die
    Versionsstände erneut abgefragt und ein ausführlicher Report ausgegeben.

    Für die Installation wird die integrierte Windows Update Agent COM-Schnittstelle genutzt. Die
    eigentliche Suche und Installation läuft auf dem Zielserver über einen temporären Scheduled Task,
    der unter dem SYSTEM-Konto ausgeführt wird. So können auch Server ohne Internetzugang gepatcht
    werden, sofern die Updates bereits über WSUS oder manuell auf dem System bereitstehen. Das Skript
    erzeugt am Ende eine CSV-Ausgabe mit den wichtigsten Ergebnissen.

.NOTES
    Autor:       Automatisiert von ChatGPT (gpt-5-codex)
    Kompatibel:  Windows Server 2016, 2019 und 2022 mit aktiviertem PowerShell Remoting
    Voraussetzungen:
        - WinRM/PowerShell-Remoting muss aktiviert sein
        - Der ausführende Benutzer benötigt lokale Administratorrechte auf den Zielsystemen
        - Updates müssen lokal im Windows Update Cache oder über WSUS bereitstehen
        - Aufgabenplanung muss aktiv sein; das Skript legt temporär Dateien unter %ProgramData%\Remote-Patching an

.EXAMPLE
    # Serverselektion anpassen und Skript mit Admin-Rechten starten
    .\Patch-WindowsServers.ps1 -Servers "APP01","DB01"

    # Report prüfen
    Import-Csv .\PatchReport-20230101-120000.csv | Format-Table
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]
    $Servers,

    [Parameter()]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter()]
    [string]
    $ReportPath = "./PatchReport-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'),

    [Parameter()]
    [int]
    $RebootTimeoutMinutes = 30,

    [Parameter()]
    [int]
    $PollIntervalSeconds = 30
)

if (-not $Credential) {
    $Credential = Get-Credential -Message "Bitte Domänenkonto mit lokalen Adminrechten angeben"
}

$script:PatchWorkerScript = @'
param(
    [string]$LogPath
)

$script:state = [ordered]@{
    Status  = "Running"
    Messages = @()
    Result  = [ordered]@{
        Available      = @()
        InstallLog     = @()
        Installed      = @()
        Failed         = @()
        RebootRequired = $false
        Error          = $null
    }
}

function Save-State {
    $json = $script:state | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($LogPath, $json, [System.Text.Encoding]::UTF8)
}

function Add-Message {
    param(
        [string]$Level,
        [string]$Message
    )

    $script:state.Messages += ,([pscustomobject]@{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Level     = $Level
        Message   = $Message
    })

    Save-State
}

function Update-Result {
    param(
        [hashtable]$Data
    )

    foreach ($key in $Data.Keys) {
        $script:state.Result[$key] = $Data[$key]
    }

    Save-State
}

Save-State
Add-Message -Level "INFO" -Message "Initialisiere Windows Update Agent"

try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $searcher = $updateSession.CreateUpdateSearcher()

    try { $searcher.ServerSelection = 1 } catch { }
    try { $searcher.Online = $false } catch { }

    Add-Message -Level "INFO" -Message "Suche nach bereitgestellten Updates"
    $searchResult = $searcher.Search("IsInstalled=0 and IsHidden=0")

    $result = [ordered]@{
        Available      = @()
        InstallLog     = @()
        Installed      = @()
        Failed         = @()
        RebootRequired = $false
        Error          = $null
    }

    if ($searchResult.Updates.Count -eq 0) {
        Add-Message -Level "SUCCESS" -Message "Keine Updates gefunden"
        Update-Result -Data $result
        $script:state.Status = "Completed"
        Save-State
        return
    }

    for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
        $update = $searchResult.Updates.Item($i)
        $kb = if ($update.KBArticleIDs -and $update.KBArticleIDs.Count -gt 0) { $update.KBArticleIDs -join ',' } else { '' }
        $result.Available += ,([pscustomobject]@{
            KB           = $kb
            Title        = $update.Title
            IsDownloaded = [bool]$update.IsDownloaded
            NeedsReboot  = [int]$update.InstallationBehavior.RebootBehavior -ge 1
        })
    }

    $downloadCollection = New-Object -ComObject Microsoft.Update.UpdateColl
    $installCollection = New-Object -ComObject Microsoft.Update.UpdateColl

    for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
        $update = $searchResult.Updates.Item($i)
        if (-not $update.EulaAccepted) {
            $update.AcceptEula()
        }

        $installCollection.Add($update) | Out-Null

        if (-not $update.IsDownloaded) {
            $downloadCollection.Add($update) | Out-Null
        }
    }

    if ($downloadCollection.Count -gt 0) {
        Add-Message -Level "INFO" -Message "Lade fehlende Updatepakete herunter"
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $downloadCollection
        $downloadResult = $downloader.Download()
        $result.InstallLog += ,([pscustomobject]@{
            Stage   = "Download"
            KB      = ""
            Title   = "Alle Updates"
            Status  = $downloadResult.ResultCode.ToString()
            HResult = ('{0:X8}' -f $downloadResult.HResult)
        })
        if ($downloadResult.ResultCode -ne 2) {
            $result.Error = "Download fehlgeschlagen"
            Update-Result -Data $result
            $script:state.Status = "Failed"
            Add-Message -Level "ERROR" -Message "Download der Updates fehlgeschlagen"
            return
        }
    }

    Add-Message -Level "INFO" -Message "Starte Installation der Updates"
    $installer = $updateSession.CreateUpdateInstaller()
    $installer.Updates = $installCollection
    $installationResult = $installer.Install()
    $result.RebootRequired = [bool]$installationResult.RebootRequired

    for ($i = 0; $i -lt $installCollection.Count; $i++) {
        $update = $installCollection.Item($i)
        $kb = if ($update.KBArticleIDs -and $update.KBArticleIDs.Count -gt 0) { $update.KBArticleIDs -join ',' } else { '' }
        $updateResult = $installationResult.GetUpdateResult($i)
        $status = switch ([int]$updateResult.ResultCode) {
            2 { "Installed" }
            3 { "SucceededWithErrors" }
            4 { "Failed" }
            5 { "Aborted" }
            default { $updateResult.ResultCode.ToString() }
        }

        $entry = [pscustomobject]@{
            Stage   = "Install"
            KB      = $kb
            Title   = $update.Title
            Status  = $status
            HResult = ('{0:X8}' -f $updateResult.HResult)
        }

        $result.InstallLog += ,$entry

        $kbLabel = if ($kb) { $kb } else { "KB-unbekannt" }
        Add-Message -Level (if ($status -eq "Installed") { "SUCCESS" } elseif ($status -eq "Failed" -or $status -eq "Aborted") { "WARN" } else { "INFO" }) -Message ("{0}: {1} -> {2}" -f $kbLabel, $update.Title, $status)

        if ($status -eq "Installed" -or $status -eq "SucceededWithErrors") {
            if ($kb) { $result.Installed += ,$kb }
        }
        elseif ($status -eq "Failed" -or $status -eq "Aborted") {
            if ($kb) { $result.Failed += ,$kb }
        }
    }

    if ($installationResult.ResultCode -ne 2 -and $installationResult.ResultCode -ne 5) {
        $result.Error = "Installation meldete Fehlercode $($installationResult.ResultCode)"
    }

    Update-Result -Data $result
    $script:state.Status = if ($result.Error) { "Failed" } else { "Completed" }
    if ($result.Error) {
        Add-Message -Level "WARN" -Message $result.Error
    }
    else {
        if ($result.Installed.Count -gt 0) {
            Add-Message -Level "SUCCESS" -Message ("Installierte KBs: {0}" -f ($result.Installed -join ", "))
        }
        else {
            Add-Message -Level "INFO" -Message "Keine Rückmeldung zu installierten KBs erhalten"
        }
    }
    Save-State
}
catch {
    $script:state.Result.Error = $_.Exception.Message
    $script:state.Status = "Failed"
    Add-Message -Level "ERROR" -Message ("Fehler: {0}" -f $_.Exception.Message)
    Save-State
}
'@

function Write-Stage {
    param(
        [string]$Server,
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
        default { 'Gray' }
    }

    Write-Host "[$timestamp][$Server][$Level] $Message" -ForegroundColor $color
}

function Invoke-ServerPatch {
    param(
        [string]$Server,
        [System.Management.Automation.PSCredential]$Credential,
        [int]$RebootTimeoutMinutes,
        [int]$PollIntervalSeconds
    )

    $result = [ordered]@{
        Server            = $Server
        Status            = 'NotStarted'
        Error             = $null
        AvailableUpdates  = $null
        Installed         = $null
        OsVersionBefore   = $null
        OsVersionAfter    = $null
        OsBuildBefore     = $null
        OsBuildAfter      = $null
        RebootTriggered   = $false
        RebootCompleted   = $false
        DurationMinutes   = 0
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        Write-Stage -Server $Server -Message 'Starte Verbindung zum Server'
        $session = New-PSSession -ComputerName $Server -Credential $Credential -ErrorAction Stop

        try {
            Write-Stage -Server $Server -Message 'Ermittle aktuellen Patch-Stand'
            $preState = Invoke-Command -Session $session -ScriptBlock {
                $info = Get-ComputerInfo | Select-Object -First 1 OsName, OsVersion, WindowsVersion, OsBuildNumber
                $hotfixes = (Get-HotFix | Select-Object -ExpandProperty HotFixID)
                [pscustomobject]@{
                    OsName      = $info.OsName
                    OsVersion   = $info.OsVersion
                    WindowsVersion = $info.WindowsVersion
                    OsBuild     = $info.OsBuildNumber
                    HotFixes    = $hotfixes
                }
            }

            $result.OsVersionBefore = $preState.OsVersion
            $result.OsBuildBefore = $preState.OsBuild

            Write-Stage -Server $Server -Message 'Stoße Windows Update Suche nach lokalen Updates an'

            $runId = [guid]::NewGuid().ToString()
            $taskName = "Remote-Patching-$runId"

            $setupInfo = Invoke-Command -Session $session -ScriptBlock {
                param($WorkerContent, $TaskName)

                Import-Module ScheduledTasks -ErrorAction Stop | Out-Null

                $root = Join-Path $env:ProgramData 'Remote-Patching'
                if (-not (Test-Path $root)) {
                    New-Item -ItemType Directory -Path $root -Force | Out-Null
                }

                $scriptPath = Join-Path $root ("{0}.ps1" -f $TaskName)
                $logPath = Join-Path $root ("{0}.json" -f $TaskName)

                [System.IO.File]::WriteAllText($scriptPath, $WorkerContent, [System.Text.Encoding]::UTF8)

                $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
                $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -LogPath `"$logPath`""
                $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

                Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
                Start-ScheduledTask -TaskName $TaskName

                [pscustomobject]@{
                    ScriptPath = $scriptPath
                    LogPath    = $logPath
                    TaskName   = $TaskName
                }
            } -ArgumentList $script:PatchWorkerScript, $taskName

            if (-not $setupInfo) {
                throw 'Konnte geplanten Updateauftrag nicht erstellen'
            }

            Write-Stage -Server $Server -Message 'Lokalen SYSTEM-Task für Updates gestartet'

            $pollDeadline = (Get-Date).AddHours(2)
            $lastMessageCount = 0
            $workerStatus = 'Running'
            $wuOutcome = $null

            do {
                Start-Sleep -Seconds ([Math]::Max(5, [int][Math]::Min(30, $PollIntervalSeconds)))

                $logContent = Invoke-Command -Session $session -ScriptBlock {
                    param($Path)
                    if (Test-Path $Path) {
                        Get-Content -Path $Path -Raw
                    }
                } -ArgumentList $setupInfo.LogPath

                if ($logContent) {
                    try {
                        $logObject = $logContent | ConvertFrom-Json -ErrorAction Stop
                    }
                    catch {
                        $logObject = $null
                    }

                    if ($logObject) {
                        if ($logObject.Messages) {
                            $total = $logObject.Messages.Count
                            for ($idx = $lastMessageCount; $idx -lt $total; $idx++) {
                                $entry = $logObject.Messages[$idx]
                                if (-not $entry) { continue }
                                $level = if ($entry.Level) { $entry.Level } else { 'INFO' }
                                $message = if ($entry.Message) { $entry.Message } else { '' }
                                if ($message) {
                                    Write-Stage -Server $Server -Message $message -Level $level
                                }
                            }
                            $lastMessageCount = $total
                        }

                        if ($logObject.Status) {
                            $workerStatus = $logObject.Status
                        }

                        if ($logObject.Status -in @('Completed', 'Failed')) {
                            if ($logObject.Result) {
                                $wuOutcome = $logObject.Result
                            }
                            break
                        }
                    }
                }

                $taskState = Invoke-Command -Session $session -ScriptBlock {
                    param($TaskName)
                    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
                    if ($task) {
                        $task.State
                    } else {
                        $null
                    }
                } -ArgumentList $setupInfo.TaskName

                if (-not $taskState -and -not $logContent) {
                    break
                }
            } while ((Get-Date) -lt $pollDeadline)

            $finalLogContent = Invoke-Command -Session $session -ScriptBlock {
                param($TaskName, $ScriptPath, $LogPath)

                try { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue } catch { }
                try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }

                $content = $null
                if (Test-Path $LogPath) {
                    $content = Get-Content -Path $LogPath -Raw
                    Remove-Item $LogPath -Force -ErrorAction SilentlyContinue
                }

                if (Test-Path $ScriptPath) {
                    Remove-Item $ScriptPath -Force -ErrorAction SilentlyContinue
                }

                $content
            } -ArgumentList $setupInfo.TaskName, $setupInfo.ScriptPath, $setupInfo.LogPath

            if (-not $wuOutcome -and $finalLogContent) {
                try {
                    $finalLog = $finalLogContent | ConvertFrom-Json -ErrorAction Stop
                    if ($finalLog.Result) {
                        $wuOutcome = $finalLog.Result
                        if ($finalLog.Status) { $workerStatus = $finalLog.Status }
                    }
                }
                catch {
                    # ignore parse errors
                }
            }

            if (-not $wuOutcome) {
                throw 'Updateauftrag lieferte keine Ergebnisse. Bitte Task- und WindowsUpdate-Logs prüfen.'
            }

            if ($workerStatus -eq 'Failed' -and -not $wuOutcome.Error) {
                $wuOutcome | Add-Member -NotePropertyName Error -NotePropertyValue 'Updateauftrag meldete Fehler ohne Detailangaben' -Force
            }

            if ($wuOutcome.Error -and (-not $wuOutcome.Available -or $wuOutcome.Available.Count -eq 0) -and (-not $wuOutcome.InstallLog -or $wuOutcome.InstallLog.Count -eq 0)) {
                $result.Status = 'Failed'
                $result.Error = $wuOutcome.Error
                Write-Stage -Server $Server -Message $wuOutcome.Error -Level 'ERROR'
            }
            elseif (-not $wuOutcome.Available -or $wuOutcome.Available.Count -eq 0) {
                Write-Stage -Server $Server -Message 'Keine neuen Updates gefunden' -Level 'SUCCESS'
                $result.AvailableUpdates = ''
                $result.Status = 'Compliant'
            }
            else {
                $kbList = $wuOutcome.Available | ForEach-Object { $_.KB } | Where-Object { $_ }
                $result.AvailableUpdates = ($kbList -join ',')
                $displayList = $wuOutcome.Available | ForEach-Object {
                    $kb = if ($_.KB) { $_.KB } else { 'KB-unbekannt' }
                    "{0} - {1}" -f $kb, $_.Title
                }
                Write-Stage -Server $Server -Message ("Folgende Updates werden installiert: {0}" -f ($displayList -join '; '))

                foreach ($logEntry in $wuOutcome.InstallLog) {
                    if ($logEntry.Stage -eq 'Install') {
                        $kbLabel = if ($logEntry.KB) { $logEntry.KB } else { 'KB-unbekannt' }
                        Write-Stage -Server $Server -Message ("{0}: {1} -> {2}" -f $kbLabel, $logEntry.Title, $logEntry.Status)
                    }
                    else {
                        Write-Stage -Server $Server -Message ("{0}: {1}" -f $logEntry.Stage, $logEntry.Status)
                    }
                }

                if ($wuOutcome.Installed -and $wuOutcome.Installed.Count -gt 0) {
                    $result.Installed = ($wuOutcome.Installed -join ',')
                }

                if ($wuOutcome.Error) {
                    $result.Status = if ($wuOutcome.Installed -and $wuOutcome.Installed.Count -gt 0) { 'Partial' } else { 'Failed' }
                    $result.Error = $wuOutcome.Error
                    Write-Stage -Server $Server -Message $wuOutcome.Error -Level 'WARN'
                }
                else {
                    $result.Status = if ($wuOutcome.Installed -and $wuOutcome.Installed.Count -gt 0) { 'Installed' } else { 'Partial' }
                }

                Write-Stage -Server $Server -Message 'Prüfe, ob ein Neustart erforderlich ist'
                $needsReboot = [bool]$wuOutcome.RebootRequired

                if ($needsReboot) {
                    Write-Stage -Server $Server -Message 'Server benötigt Neustart, löse Neustart aus'
                    $result.RebootTriggered = $true
                    Invoke-Command -Session $session -ScriptBlock {
                        Restart-Computer -Force
                    }

                    Remove-PSSession -Session $session
                    $session = $null

                    $deadline = (Get-Date).AddMinutes($RebootTimeoutMinutes)
                    Write-Stage -Server $Server -Message 'Warte auf erfolgreiche Anmeldung nach Neustart'
                    do {
                        Start-Sleep -Seconds $PollIntervalSeconds
                        $reachable = Test-Connection -ComputerName $Server -Quiet -Count 1 -ErrorAction SilentlyContinue
                    } while (-not $reachable -and (Get-Date) -lt $deadline)

                    if (-not $reachable) {
                        throw "Server hat nach Neustart nicht geantwortet"
                    }

                    $result.RebootCompleted = $true
                    Write-Stage -Server $Server -Message 'Server wieder erreichbar. Prüfe Endzustand.' -Level 'SUCCESS'

                    Write-Stage -Server $Server -Message 'Stelle Remoting-Sitzung wieder her'
                    $reconnectDeadline = (Get-Date).AddMinutes(5)
                    do {
                        try {
                            $session = New-PSSession -ComputerName $Server -Credential $Credential -ErrorAction Stop
                            $reconnected = $true
                        }
                        catch {
                            $reconnected = $false
                            Start-Sleep -Seconds 15
                        }
                    } while (-not $reconnected -and (Get-Date) -lt $reconnectDeadline)

                    if (-not $reconnected) {
                        throw 'Remoting-Sitzung konnte nach dem Neustart nicht wiederhergestellt werden'
                    }
                }

                Write-Stage -Server $Server -Message 'Lese finale Systeminformationen'
                $postState = Invoke-Command -Session $session -ScriptBlock {
                    $info = Get-ComputerInfo | Select-Object -First 1 OsName, OsVersion, WindowsVersion, OsBuildNumber
                    $hotfixes = (Get-HotFix | Select-Object -ExpandProperty HotFixID)
                    [pscustomobject]@{
                        OsName      = $info.OsName
                        OsVersion   = $info.OsVersion
                        WindowsVersion = $info.WindowsVersion
                        OsBuild     = $info.OsBuildNumber
                        HotFixes    = $hotfixes
                    }
                }

                $result.OsVersionAfter = $postState.OsVersion
                $result.OsBuildAfter = $postState.OsBuild
                $newHotFixes = $postState.HotFixes | Where-Object { $_ -notin $preState.HotFixes }
                if ($newHotFixes -and $newHotFixes.Count -gt 0) {
                    $result.Installed = ($newHotFixes -join ',')
                    $result.Status = 'Patched'
                    Write-Stage -Server $Server -Message ("Installierte KBs: {0}" -f ($newHotFixes -join ', ')) -Level 'SUCCESS'
                }
                elseif ($wuOutcome -and $wuOutcome.Installed -and $wuOutcome.Installed.Count -gt 0) {
                    $result.Status = 'Patched'
                    Write-Stage -Server $Server -Message ("Updates erfolgreich abgeschlossen: {0}" -f ($wuOutcome.Installed -join ', ')) -Level 'SUCCESS'
                }
                elseif ($wuOutcome -and $wuOutcome.Error -and $result.Status -ne 'Failed') {
                    $result.Status = 'Failed'
                    if (-not $result.Error) { $result.Error = $wuOutcome.Error }
                    Write-Stage -Server $Server -Message $result.Error -Level 'WARN'
                }
                elseif ($wuOutcome -and $wuOutcome.Available -and $wuOutcome.Available.Count -gt 0) {
                    $failedEntries = $wuOutcome.InstallLog | Where-Object { $_.Stage -eq 'Install' -and $_.Status -in @('Failed','Aborted') }
                    if ($failedEntries -and $failedEntries.Count -gt 0) {
                        $result.Status = 'Partial'
                        $failedList = $failedEntries | ForEach-Object { if ($_.KB) { $_.KB } else { $_.Title } }
                        $result.Error = "Fehler bei folgenden Updates: $($failedList -join ', ')"
                        Write-Stage -Server $Server -Message $result.Error -Level 'WARN'
                    }
                    else {
                        $result.Status = 'Partial'
                        $result.Error = 'Keine Bestätigung für installierte Updates erhalten. Bitte Windows Update Log prüfen.'
                        Write-Stage -Server $Server -Message $result.Error -Level 'WARN'
                    }
                }
            }
        }
        finally {
            if ($session) {
                Remove-PSSession -Session $session
            }
        }
    }
    catch {
        $result.Status = 'Failed'
        $result.Error = $_.Exception.Message
        Write-Stage -Server $Server -Message $result.Error -Level 'ERROR'
    }
    finally {
        $stopwatch.Stop()
        $result.DurationMinutes = [Math]::Round(($stopwatch.Elapsed.TotalMinutes), 2)
    }

    return [pscustomobject]$result
}

$overallResults = @()
$serverCount = $Servers.Count
$index = 0

foreach ($server in $Servers) {
    $index++
    $percent = [int](($index / $serverCount) * 100)
    Write-Progress -Activity 'Patch-Deployment' -Status "Bearbeite $server ($index von $serverCount)" -PercentComplete $percent
    $overallResults += Invoke-ServerPatch -Server $server -Credential $Credential -RebootTimeoutMinutes $RebootTimeoutMinutes -PollIntervalSeconds $PollIntervalSeconds
}

Write-Progress -Activity 'Patch-Deployment' -Completed

$overallResults | Sort-Object Server | Format-Table -AutoSize
$overallResults | Export-Csv -Path $ReportPath -Encoding UTF8 -Delimiter ';' -NoTypeInformation

Write-Host "Report gespeichert unter: $ReportPath" -ForegroundColor Green
