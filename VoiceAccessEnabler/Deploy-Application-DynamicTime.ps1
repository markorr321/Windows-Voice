<#
.SYNOPSIS

PSApppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION

- The script is provided as a template to perform an install or uninstall of an application(s).
- The script either performs an "Install" deployment type or an "Uninstall" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.

PSApppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2023 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham and Muhammad Mashwani).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType

The type of deployment to perform. Default is: Install.

.PARAMETER DeployMode

Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru

Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode

Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging

Disables logging to file for the script. Default is: $false.

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"

.EXAMPLE

Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS

None

You cannot pipe objects to this script.

.OUTPUTS

None

This script does not generate any output.

.NOTES

Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
- 69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
- 70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1

.LINK

https://psappdeploytoolkit.com
#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [String]$DeploymentType = 'Install',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [String]$DeployMode = 'Interactive',
    [Parameter(Mandatory = $false)]
    [switch]$AllowRebootPassThru = $false,
    [Parameter(Mandatory = $false)]
    [switch]$TerminalServerMode = $false,
    [Parameter(Mandatory = $false)]
    [switch]$DisableLogging = $false
)

Try {
    ## Set the script execution policy for this process
    Try {
        Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
    }
    Catch {
    }

    ##*===============================================
    ##* VARIABLE DECLARATION
    ##*===============================================
    ## Variables: Application
    [String]$appVendor = 'IT Department'
    [String]$appName = 'Windows Voice Access Configuration'
    [String]$appVersion = '1.1'
    [String]$appArch = 'x64'
    [String]$appLang = 'EN'
    [String]$appRevision = '01'
    [String]$appScriptVersion = '1.1.0'
    [String]$appScriptDate = '02/04/2026'
    [String]$appScriptAuthor = 'IT Admin'
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [String]$installName = ''
    [String]$installTitle = ''

    ##* Do not modify section below
    #region DoNotModify

    ## Variables: Exit Code
    [Int32]$mainExitCode = 0

    ## Variables: Script
    [String]$deployAppScriptFriendlyName = 'Deploy Application'
    [Version]$deployAppScriptVersion = [Version]'3.9.3'
    [String]$deployAppScriptDate = '02/05/2023'
    [Hashtable]$deployAppScriptParameters = $PsBoundParameters

    ## Variables: Environment
    If (Test-Path -LiteralPath 'variable:HostInvocation') {
        $InvocationInfo = $HostInvocation
    }
    Else {
        $InvocationInfo = $MyInvocation
    }
    [String]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

    ## Dot source the required App Deploy Toolkit Functions
    Try {
        [String]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) {
            Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]."
        }
        If ($DisableLogging) {
            . $moduleAppDeployToolkitMain -DisableLogging
        }
        Else {
            . $moduleAppDeployToolkitMain
        }
    }
    Catch {
        If ($mainExitCode -eq 0) {
            [Int32]$mainExitCode = 60008
        }
        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
        ## Exit the script, returning the exit code to SCCM
        If (Test-Path -LiteralPath 'variable:HostInvocation') {
            $script:ExitCode = $mainExitCode; Exit
        }
        Else {
            Exit $mainExitCode
        }
    }

    #endregion
    ##* Do not modify section above
    ##*===============================================
    ##* END VARIABLE DECLARATION
    ##*===============================================

    If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
        ##*===============================================
        ##* PRE-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Installation'

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Installation tasks here>


        ##*===============================================
        ##* INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Installation'

        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
                $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
            }
        }

        ## <Perform Installation tasks here>

        ## Show Progress
        Show-InstallationProgress -StatusMessage 'Enabling Store backend services for Voice Access...'

        ## Run Enable Script
        $enableScript = Join-Path -Path $dirFiles -ChildPath 'Enable-Store-Backend-Services.ps1'
        If (Test-Path -LiteralPath $enableScript -PathType 'Leaf') {
            $result = Execute-Process -Path "$envWinDir\System32\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-ExecutionPolicy Bypass -NoProfile -File `"$enableScript`"" -WindowStyle 'Hidden' -PassThru
            If ($result.ExitCode -eq 0) {
                Write-Log -Message 'Store backend services enabled successfully'
            } Else {
                Write-Log -Message "Enable script failed with exit code: $($result.ExitCode)" -Severity 3
                Show-InstallationPrompt -Message 'Failed to enable Store backend services. Please contact IT support.' -ButtonRightText 'OK' -Icon Error
                Exit-Script -ExitCode $result.ExitCode
            }
        } Else {
            Write-Log -Message "Enable script not found at: $enableScript" -Severity 3
            Show-InstallationPrompt -Message 'Installation files are missing. Please contact IT support.' -ButtonRightText 'OK' -Icon Error
            Exit-Script -ExitCode 1
        }


        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Installation'

        ## <Perform Post-Installation tasks here>

        ## Create registry detection for supersedence support
        $regPath = "HKLM:\SOFTWARE\VoiceAccessEnabler"
        If (-not (Test-Path -LiteralPath $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction 'SilentlyContinue' | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "Version" -Value $appVersion -Type String -ErrorAction 'SilentlyContinue'
        Set-ItemProperty -Path $regPath -Name "DisplayName" -Value $appName -Type String -ErrorAction 'SilentlyContinue'
        Set-ItemProperty -Path $regPath -Name "Publisher" -Value $appVendor -Type String -ErrorAction 'SilentlyContinue'
        Set-ItemProperty -Path $regPath -Name "InstallDate" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Type String -ErrorAction 'SilentlyContinue'
        Write-Log -Message "Registry detection created at: $regPath"

        ## Calculate actual time remaining until Intune policy refresh
        $timeRemaining = 60  # Default fallback
        Try {
            $configTask = Get-ScheduledTask | Where-Object {$_.TaskName -eq "Schedule created by dm client to refresh settings"} | Select-Object -First 1
            If ($configTask) {
                $taskInfo = Get-ScheduledTaskInfo $configTask
                If ($taskInfo.NextRunTime) {
                    $minutesRemaining = [Math]::Round(($taskInfo.NextRunTime - (Get-Date)).TotalMinutes)
                    If ($minutesRemaining -gt 0) {
                        $timeRemaining = $minutesRemaining
                        Write-Log -Message "Calculated time remaining: $timeRemaining minutes (Next policy refresh: $($taskInfo.NextRunTime))"
                    }
                }
            }
        } Catch {
            Write-Log -Message "Could not calculate exact time remaining, using default: $($_.Exception.Message)" -Severity 2
        }

        ## Save setup instructions to logged-in user's Desktop (checks OneDrive redirect)
        Try {
            $desktopPath = $null
            $loggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
            If ($loggedInUser) {
                $username = $loggedInUser.Split('\')[-1]
                $userProfile = "C:\Users\$username"
                ## Check for OneDrive Desktop first
                $oneDriveDesktop = Get-ChildItem -Path $userProfile -Filter "OneDrive*" -Directory -ErrorAction SilentlyContinue |
                    Where-Object { Test-Path (Join-Path $_.FullName 'Desktop') } |
                    Select-Object -First 1
                If ($oneDriveDesktop) {
                    $desktopPath = Join-Path $oneDriveDesktop.FullName 'Desktop'
                } Else {
                    $desktopPath = Join-Path $userProfile 'Desktop'
                }
            }
            Write-Log -Message "Desktop path resolved to: $desktopPath"
            $instructionsFile = Join-Path -Path $desktopPath -ChildPath 'Voice Access Setup Instructions.txt'
            $instructionsContent = @"
Voice Access Setup Instructions
================================

You have $timeRemaining minutes to complete the Voice Access setup wizard.

To set up Voice Access:
1. Open Settings
2. Go to Accessibility > Interaction > Speech > Voice Access
3. Toggle Voice Access to on
4. Follow the setup wizard

Intune will automatically revert Store backend changes within $timeRemaining minutes.

After completing Voice Access setup, you may delete this file.
"@
            Set-Content -Path $instructionsFile -Value $instructionsContent -Force -ErrorAction 'SilentlyContinue'
            Write-Log -Message "Setup instructions saved to: $instructionsFile"
        } Catch {
            Write-Log -Message "Could not save instructions to Desktop: $($_.Exception.Message)" -Severity 2
        }

        ## Display completion message to logged-in user
        If ($deployMode -eq 'NonInteractive') {
            Write-Log -Message 'Running in NonInteractive mode - skipping dialog'
        } Else {
            # Running in Interactive mode - show dialog directly
            Write-Log -Message 'Running in Interactive mode - showing dialog directly'
            Show-InstallationPrompt -Message "Voice Access is now enabled!`n`nYou have $timeRemaining minutes to complete the Voice Access setup wizard.`n`nSetup instructions have been saved to your Desktop:`nVoice Access Setup Instructions.txt`n`nYou must click OK to complete the installation.`nPlease wait for the install to finish before configuring Voice Access." -ButtonRightText 'OK' -TopMost $true
            Write-Log -Message 'Completion message displayed to user'
        }
    }
    ElseIf ($deploymentType -ieq 'Uninstall') {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Uninstallation'

        ## <Perform Pre-Uninstallation tasks here>


        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Uninstallation'

        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }

        ## <Perform Uninstallation tasks here>

        ## Remove registry detection key for supersedence support
        Write-Log -Message 'Removing registry detection key for supersedence...'
        Remove-Item -Path "HKLM:\SOFTWARE\VoiceAccessEnabler" -Recurse -Force -ErrorAction 'SilentlyContinue'
        Write-Log -Message 'Uninstallation complete. Intune policies will automatically revert Store backend changes.'


        ##*===============================================
        ##* POST-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Uninstallation'

        ## <Perform Post-Uninstallation tasks here>


    }
    ElseIf ($deploymentType -ieq 'Repair') {
        ##*===============================================
        ##* PRE-REPAIR
        ##*===============================================
        [String]$installPhase = 'Pre-Repair'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Repair tasks here>

        ##*===============================================
        ##* REPAIR
        ##*===============================================
        [String]$installPhase = 'Repair'

        ## Handle Zero-Config MSI Repairs
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }
        ## <Perform Repair tasks here>

        ##*===============================================
        ##* POST-REPAIR
        ##*===============================================
        [String]$installPhase = 'Post-Repair'

        ## <Perform Post-Repair tasks here>


    }
    ##*===============================================
    ##* END SCRIPT BODY
    ##*===============================================

    ## Call the Exit-Script function to perform final cleanup operations
    Exit-Script -ExitCode $mainExitCode
}
Catch {
    [Int32]$mainExitCode = 60001
    [String]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}
