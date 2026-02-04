# Show-UserMessage.ps1
# Displays a simple message box to the logged-in user
# This script is designed to run as the logged-in user via scheduled task

param(
    [Parameter(Mandatory=$true)]
    [string]$Message,

    [Parameter(Mandatory=$false)]
    [string]$Title = "Voice Access Enabler",

    [Parameter(Mandatory=$false)]
    [ValidateSet('Information','Warning','Error')]
    [string]$Icon = 'Information'
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Map icon type to MessageBoxIcon
$iconType = switch ($Icon) {
    'Information' { [System.Windows.Forms.MessageBoxIcon]::Information }
    'Warning' { [System.Windows.Forms.MessageBoxIcon]::Warning }
    'Error' { [System.Windows.Forms.MessageBoxIcon]::Error }
    default { [System.Windows.Forms.MessageBoxIcon]::Information }
}

# Show the message box
[System.Windows.Forms.MessageBox]::Show(
    $Message,
    $Title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    $iconType
) | Out-Null

Exit 0
