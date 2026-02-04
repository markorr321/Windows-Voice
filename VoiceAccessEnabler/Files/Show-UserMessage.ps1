# Show-UserMessage.ps1
# Displays a custom form with formatted message to the logged-in user
# This script is designed to run as the logged-in user via scheduled task

param(
    [Parameter(Mandatory=$true)]
    [string]$Message,

    [Parameter(Mandatory=$false)]
    [string]$Title = "Voice Access Enabler",

    [Parameter(Mandatory=$false)]
    [ValidateSet('Information','Warning','Error','None')]
    [string]$Icon = 'None'
)

$ErrorActionPreference = 'Stop'
$logFile = "$env:TEMP\Show-UserMessage.log"

try {
    "$(Get-Date): Starting Show-UserMessage.ps1" | Out-File $logFile -Append
    "$(Get-Date): Title: $Title" | Out-File $logFile -Append
    "$(Get-Date): Message length: $($Message.Length)" | Out-File $logFile -Append

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    "$(Get-Date): Assemblies loaded" | Out-File $logFile -Append

    # Create custom form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    "$(Get-Date): Form created" | Out-File $logFile -Append

    # Create label for message with bold font
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(450, 240)
    $label.Text = $Message
    # Create bold font
    $boldFont = New-Object System.Drawing.Font('Microsoft Sans Serif', 13, [System.Drawing.FontStyle]::Bold)
    $label.Font = $boldFont
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.AutoSize = $false

    "$(Get-Date): Label created" | Out-File $logFile -Append

    # Create OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(200, 270)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)

    # Add controls to form
    $form.Controls.Add($label)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    "$(Get-Date): About to show dialog" | Out-File $logFile -Append

    # Show form
    $result = $form.ShowDialog()

    "$(Get-Date): Dialog closed with result: $result" | Out-File $logFile -Append

    Exit 0
}
catch {
    "$(Get-Date): ERROR: $($_.Exception.Message)" | Out-File $logFile -Append
    "$(Get-Date): Stack trace: $($_.ScriptStackTrace)" | Out-File $logFile -Append
    Exit 1
}
