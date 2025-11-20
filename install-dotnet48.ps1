[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:TEMP\\dotnet48-install.log"
)

function Test-Administrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-NetFrameworkRelease {
    $releaseKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
    return $releaseKey.Release
}

function Test-NetFramework48Installed {
    $release = Get-NetFrameworkRelease
    return ($release -and $release -ge 528040)
}

if (-not (Test-Administrator)) {
    Write-Error "This script must be run with elevated permissions. Open PowerShell as Administrator and try again."
    exit 1
}

if (Test-NetFramework48Installed) {
    Write-Host ".NET Framework 4.8 or later is already installed."
    exit 0
}

if (-not (Test-Path -Path $InstallerPath)) {
    Write-Error "Installer path '$InstallerPath' does not exist. Provide a valid path to the offline installer (e.g., from a network share)."
    exit 1
}

$arguments = @("/q", "/norestart", "/log", "`"$LogPath`"")
Write-Host "Starting .NET Framework 4.8 installation..."
$process = Start-Process -FilePath $InstallerPath -ArgumentList $arguments -Wait -PassThru

switch ($process.ExitCode) {
    0 {
        Write-Host "Installation completed successfully."
    }
    3010 {
        Write-Host "Installation completed. A restart is required to finish the setup, but the server was not restarted."
        try {
            Add-Content -Path $LogPath -Value "Note: Installer reported exit code 3010 (restart required). Restart intentionally not performed by script."
        } catch {
            Write-Warning "Could not record restart note to log file at ${LogPath}: $($_.Exception.Message)"
        }
    }
    default {
        Write-Error "Installer returned exit code $($process.ExitCode). Check the log at $LogPath for details."; exit $process.ExitCode
    }
}

if (Test-NetFramework48Installed) {
    Write-Host ".NET Framework 4.8 installation verified."
} else {
    Write-Warning "Could not verify .NET Framework 4.8 after installation."
}
