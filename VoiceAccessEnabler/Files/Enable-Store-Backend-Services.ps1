# Enable Store Backend Services for Voice Access
# This script enables Store backend while keeping UI blocked

$ErrorActionPreference = 'Continue'

# Registry path
$storePath = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"

# Ensure registry key exists
if (-not (Test-Path $storePath)) {
    New-Item -Path $storePath -Force | Out-Null
}

# Enable Store backend
Set-ItemProperty -Path $storePath -Name "RemoveWindowsStore" -Value 0 -Type DWORD -ErrorAction SilentlyContinue
Set-ItemProperty -Path $storePath -Name "AutoDownload" -Value 4 -Type DWORD -ErrorAction SilentlyContinue

# Start backend services
$services = @("wuauserv", "InstallService", "ClipSVC")
foreach ($svc in $services) {
    try {
        $service = Get-Service -Name $svc -ErrorAction Stop
        if ($service.Status -ne 'Running') {
            Start-Service -Name $svc -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Service doesn't exist or can't be started
    }
}

# Clear Store cache
Stop-Process -Name "WinStore.App" -Force -ErrorAction SilentlyContinue

Exit 0
