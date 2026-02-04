# Windows Voice Access Configuration - Simple Installation
# No PSADT, No Scheduled Tasks

# Show user notification
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show(
    "Voice Access is being enabled!

Store backend services will be temporarily enabled for Voice Access setup.

Go to: Settings → Accessibility → Speech → Voice Access

Intune will automatically revert these changes on the next policy sync (typically within 8 hours).",
    "Windows Voice Access Configuration",
    "OK",
    "Information"
) | Out-Null

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Run Enable script
$enableScript = Join-Path $scriptDir "Files\Enable-Store-Backend-Services.ps1"
if (Test-Path $enableScript) {
    & powershell.exe -ExecutionPolicy Bypass -NoProfile -File $enableScript
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Store backend enabled successfully" -ForegroundColor Green
    } else {
        Write-Error "Enable script failed with exit code: $LASTEXITCODE"
        Exit 1
    }
} else {
    Write-Error "Enable script not found at: $enableScript"
    Exit 1
}

# Create detection file
$detectionPath = "C:\ProgramData\VoiceAccessEnabler"
New-Item -Path $detectionPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
$installTime = Get-Date
Set-Content -Path "$detectionPath\installed.txt" -Value $installTime.ToString() -Force

Write-Host "Installation complete!" -ForegroundColor Green
Write-Host "Detection file created at: $detectionPath\installed.txt" -ForegroundColor Cyan
Exit 0
