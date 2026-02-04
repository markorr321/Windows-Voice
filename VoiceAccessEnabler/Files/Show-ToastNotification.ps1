# Show-ToastNotification.ps1
# Displays Windows toast notification to logged-in user from SYSTEM context

param(
    [Parameter(Mandatory=$true)]
    [string]$Title,

    [Parameter(Mandatory=$true)]
    [string]$Message,

    [string]$AppId = "Windows Voice Access Configuration"
)

# Get the logged-in user
$loggedOnUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
if (-not $loggedOnUser) {
    Write-Host "No user logged in, skipping toast notification"
    Exit 0
}

# Create the toast notification XML
$toastXml = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default" />
</toast>
"@

# Load required assemblies
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

# Create the toast
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($toastXml)
$toast = New-Object Windows.UI.Notifications.ToastNotification $xml

# Show the toast
$toastNotifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)
$toastNotifier.Show($toast)

Write-Host "Toast notification displayed to $loggedOnUser"
Exit 0
