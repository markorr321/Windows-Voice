# Disable Store Backend Services
# Reverses the Enable script changes

$ErrorActionPreference = 'Continue'

# Registry path
$storePath = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"

# Restore Intune Store blocks
Set-ItemProperty -Path $storePath -Name "RemoveWindowsStore" -Value 1 -Type DWORD -ErrorAction SilentlyContinue
Set-ItemProperty -Path $storePath -Name "AutoDownload" -Value 2 -Type DWORD -ErrorAction SilentlyContinue

# Restore winget msstore source
try {
    $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
    if ($wingetPath) {
        & winget source reset --force 2>&1 | Out-Null
    }
}
catch {
    # winget not available
}

Exit 0
