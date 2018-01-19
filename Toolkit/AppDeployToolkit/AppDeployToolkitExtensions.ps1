<#
.SYNOPSIS
	This script is a template that allows you to extend the toolkit with your own custom functions.
    # LICENSE #
    PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows. 
    Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
    This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. 
    You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is automatically dot-sourced by the AppDeployToolkitMain.ps1 script.
.NOTES
    Toolkit Exit Code Ranges:
    60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
    69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
    70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
)

##*===============================================
##* VARIABLE DECLARATION
##*===============================================

# Variables: Script
[string]$appDeployToolkitExtName = 'PSAppDeployToolkitExt'
[string]$appDeployExtScriptFriendlyName = 'App Deploy Toolkit Extensions'
[version]$appDeployExtScriptVersion = [version]'1.5.0'
[string]$appDeployExtScriptDate = '02/12/2017'
[hashtable]$appDeployExtScriptParameters = $PSBoundParameters

##*===============================================
##* FUNCTION LISTINGS
##*===============================================


#region Function Trigger-AppEvalCycle
Function Trigger-AppEvalCycle {
    <#
    .SYNOPSIS
    Schedule a SCCM 2012 Application Evaluation Cycle task to be triggered in the specified time.
    .DESCRIPTION
    This function is called when the user selects to defer the installation. It does the following:
    1. Removes the scheduled task configuration XML, if it already exists on the machine.
    2. Creates a temporary directory on the local machine, if the folder doesn’t exists.
    3. Creates an scheduled task configuration XML file on the temporary directory.
    4. Checks if a scheduled task with that name already exists on the machine, if it exists then delete it.
    5. Create a new scheduled task based on the XML file created on step 3.
    6. Removes the scheduled task configuration XML.
    7. Once the specified time is reached a scheduled task runs a SCCM 2012 Application Evaluation Cycle will start and it will trigger the installation/uninstallation to start if the machine is still part of the install/uninstall collection.
    .PARAMETER Time
    Specify the time, in hours, to run the scheduled task.
    .EXAMPLE
    Trigger-AppEvalCycle -Time 24
    .NOTES
    This is an internal script function and should typically not be called directly.
    It is used to ensure that when the users defers the installation a new installation attempt will be made in the specified time if the machine is still part of the install/uninstall collection.
    Version 1.0 – Jeppe Olsen.
    #>
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNullorEmpty()]
    [int]$Time = 2
    )
    
    Begin {
    ## Get the name of this function and write header
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    
    ## Specify the scheduled task configuration in XML format
    [string]$taskRunDateTime = (((Get-Date).AddHours($Time)).ToUniversalTime()).ToString(“yyyy-MM-ddTHH:mm:ss.fffffffZ”)

    
    #specify the task scheduler executable
    [string] $execSchTasks = "$envWinDir\System32\schtasks.exe"
    
    #Specify the task name
    [string]$taskName = $installName + ‘_AppEvalCycle’
    
    }
    Process {
    [string]$xmlTask = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo/>
  <Triggers>
    <TimeTrigger id="1">
        <StartBoundary>$taskRunDateTime</StartBoundary>
        <Enabled>true</Enabled>
    </TimeTrigger>
        </Triggers>
    <Principals>
        <Principal id="Author">
            <UserId>S-1-5-18</UserId>
            <RunLevel>HighestAvailable</RunLevel>
        </Principal>
    </Principals>
    <Settings>
        <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
        <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
        <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
        <AllowHardTerminate>true</AllowHardTerminate>
        <StartWhenAvailable>false</StartWhenAvailable>
        <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
        <IdleSettings>
            <StopOnIdleEnd>false</StopOnIdleEnd>
            <RestartOnIdle>false</RestartOnIdle>
        </IdleSettings>
        <AllowStartOnDemand>true</AllowStartOnDemand>
        <Enabled>true</Enabled>
        <Hidden>false</Hidden>
        <RunOnlyIfIdle>false</RunOnlyIfIdle>
        <WakeToRun>false</WakeToRun>
        <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
        <Priority>7</Priority>
    </Settings>
    <Actions Context="Author">
        <Exec id="StartPowerShellJob">
            <Command>cmd</Command>
            <Arguments>/c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000121}' /NOINTERACTIVE</Arguments>
        </Exec>
        <Exec>
            <Command>schtasks</Command>
            <Arguments>/delete /tn $taskName /f</Arguments>
        </Exec>
    </Actions>
</Task>
"@
    
    #Export the xml to file
    try{
        [string] $schXmlFile = "$dirAppDeployTemp\$taskName"
        if(-not (Test-Path $dirAppDeployTemp)){New-Item $dirAppDeployTemp -ItemType Directory -Force}
        [string] $xmlTask | Out-File -FilePath $schXmlFile -Force -ErrorAction Stop
    }
    catch{
        Write-Log -Message "Failed to export the scheduled task XML file [$schXmlFile].   `n$(Resolve-Error)" -Severity 3 -Source ${CmdLetName}
        Return
    }
    # Create scheduled task
    Write-Log -Message "Creating scheduled task to run Application Deployment Evaluation Cycle at $taskRunDateTime"
    [psobject] $taskResult = Execute-Process -Path $execSchTasks -Parameters "/create /f /tn $taskName /xml `"$schXmlFile`"" -WindowStyle Hidden -CreateNoWindow -PassThru
    If($taskResult.ExitCode -ne 0){
        Write-log -Message "Failed to create the scheduled task, with exit code : $($taskResult.ExitCode)" -Severity 3 -Source ${CmdletName}
        Return
    }

    Remove-Item $schXmlFile -Force
    }
    End {
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion



#region Function Show-InstallationWelcome
Function Show-InstallationWelcome {
    <#
    .SYNOPSIS
        Show a welcome dialog prompting the user with information about the installation and actions to be performed before the installation can begin.
    .DESCRIPTION
        The following prompts can be included in the welcome dialog:
         a) Close the specified running applications, or optionally close the applications without showing a prompt (using the -Silent switch).
         b) Defer the installation a certain number of times, for a certain number of days or until a deadline is reached.
         c) Countdown until applications are automatically closed.
         d) Prevent users from launching the specified applications while the installation is in progress.
        Notes:
         The process descriptions are retrieved from WMI, with a fall back on the process name if no description is available. Alternatively, you can specify the description yourself with a '=' symbol - see examples.
         The dialog box will timeout after the timeout specified in the XML configuration file (default 1 hour and 55 minutes) to prevent SCCM installations from timing out and returning a failure code to SCCM. When the dialog times out, the script will exit and return a 1618 code (SCCM fast retry code).
    .PARAMETER CloseApps
        Name of the process to stop (do not include the .exe). Specify multiple processes separated by a comma. Specify custom descriptions like this: "winword=Microsoft Office Word,excel=Microsoft Office Excel"
    .PARAMETER Silent
        Stop processes without prompting the user.
    .PARAMETER CloseAppsCountdown
        Option to provide a countdown in seconds until the specified applications are automatically closed. This only takes effect if deferral is not allowed or has expired.
    .PARAMETER ForceCloseAppsCountdown
        Option to provide a countdown in seconds until the specified applications are automatically closed regardless of whether deferral is allowed.
    .PARAMETER PersistPrompt
        Specify whether to make the prompt persist in the center of the screen every 10 seconds. The user will have no option but to respond to the prompt. This only takes effect if deferral is not allowed or has expired.
    .PARAMETER BlockExecution
        Option to prevent the user from launching the process/application during the installation.
    .PARAMETER AllowDefer
        Enables an optional defer button to allow the user to defer the installation.
    .PARAMETER AllowDeferCloseApps
        Enables an optional defer button to allow the user to defer the installation only if there are running applications that need to be closed.
    .PARAMETER DeferTimes
        Specify the number of times the installation can be deferred.
    .PARAMETER DeferDays
        Specify the number of days since first run that the installation can be deferred. This is converted to a deadline.
    .PARAMETER DeferDeadline
        Specify the deadline date until which the installation can be deferred.
        Specify the date in the local culture if the script is intended for that same culture.
        If the script is intended to run on EN-US machines, specify the date in the format: "08/25/2013" or "08-25-2013" or "08-25-2013 18:00:00"
        If the script is intended for multiple cultures, specify the date in the universal sortable date/time format: "2013-08-22 11:51:52Z"
        The deadline date will be displayed to the user in the format of their culture.
    .PARAMETER CheckDiskSpace
        Specify whether to check if there is enough disk space for the installation to proceed.
        If this parameter is specified without the RequiredDiskSpace parameter, the required disk space is calculated automatically based on the size of the script source and associated files.
    .PARAMETER RequiredDiskSpace
        Specify required disk space in MB, used in combination with CheckDiskSpace.
    .PARAMETER MinimizeWindows
        Specifies whether to minimize other windows when displaying prompt. Default: $true.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'iexplore,winword,excel'
        Prompt the user to close Internet Explorer, Word and Excel.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'winword,excel' -Silent
        Close Word and Excel without prompting the user.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'winword,excel' -BlockExecution
        Close Word and Excel and prevent the user from launching the applications while the installation is in progress.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'winword=Microsoft Office Word,excel=Microsoft Office Excel' -CloseAppsCountdown 600
        Prompt the user to close Word and Excel, with customized descriptions for the applications and automatically close the applications after 10 minutes.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'winword.exe,msaccess.exe,excel.exe' -PersistPrompt
        Prompt the user to close Word, MSAccess and Excel if the processes match the exact name specified (use .exe for exact matches).
        By using the PersistPrompt switch, the dialog will return to the center of the screen every 10 seconds so the user cannot ignore it by dragging it aside.
    .EXAMPLE
        Show-InstallationWelcome -AllowDefer -DeferDeadline '25/08/2013'
        Allow the user to defer the installation until the deadline is reached.
    .EXAMPLE
        Show-InstallationWelcome -CloseApps 'winword,excel' -BlockExecution -AllowDefer -DeferTimes 10 -DeferDeadline '25/08/2013' -CloseAppsCountdown 600
        Close Word and Excel and prevent the user from launching the applications while the installation is in progress.
        Allow the user to defer the installation a maximum of 10 times or until the deadline is reached, whichever happens first.
        When deferral expires, prompt the user to close the applications and automatically close them after 10 minutes.
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
        [CmdletBinding()]
        Param (
            ## Specify process names separated by commas. Optionally specify a process description with an equals symbol, e.g. "winword=Microsoft Office Word"
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [string]$CloseApps,
            ## Specify whether to prompt user or force close the applications
            [Parameter(Mandatory=$false)]
            [switch]$Silent = $false,
            ## Specify a countdown to display before automatically closing applications where deferral is not allowed or has expired
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [int32]$CloseAppsCountdown = 0,
            ## Specify a countdown to display before automatically closing applications whether or not deferral is allowed
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [int32]$ForceCloseAppsCountdown = 0,
            ## Specify whether to prompt to save working documents when the user chooses to close applications by selecting the "Close Programs" button
		    [Parameter(Mandatory=$false)]
		    [switch]$PromptToSave = $false,
            ## Specify whether to make the prompt persist in the center of the screen every 10 seconds.
            [Parameter(Mandatory=$false)]
            [switch]$PersistPrompt = $false,
            ## Specify whether to block execution of the processes during installation
            [Parameter(Mandatory=$false)]
            [switch]$BlockExecution = $false,
            ## Specify whether to enable the optional defer button on the dialog box
            [Parameter(Mandatory=$false)]
            [switch]$AllowDefer = $false,
            ## Specify whether to enable the optional defer button on the dialog box only if an app needs to be closed
            [Parameter(Mandatory=$false)]
            [switch]$AllowDeferCloseApps = $false,
            ## Specify the number of times the deferral is allowed
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [int32]$DeferTimes = 0,
            ## Specify the number of days since first run that the deferral is allowed
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [int32]$DeferDays = 0,
            ## Specify the deadline (in format dd/mm/yyyy) for which deferral will expire as an option
            [Parameter(Mandatory=$false)]
            [string]$DeferDeadline = '',
            ## Specify whether to check if there is enough disk space for the installation to proceed. If this parameter is specified without the RequiredDiskSpace parameter, the required disk space is calculated automatically based on the size of the script source and associated files.
            [Parameter(Mandatory=$false)]
            [switch]$CheckDiskSpace = $false,
            ## Specify required disk space in MB, used in combination with $CheckDiskSpace.
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [int32]$RequiredDiskSpace = 0,
            ## Specify whether to minimize other windows when displaying prompt
            [Parameter(Mandatory=$false)]
            [ValidateNotNullorEmpty()]
            [boolean]$MinimizeWindows = $true
        )
        
        Begin {
            ## Get the name of this function and write header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
            #Determine the size of the package
            if($RequiredDiskSpace -eq 0){
                $fso = New-Object -ComObject 'Scripting.FileSystemObject' -ErrorAction 'Stop'
                $RequiredDiskSpace = [math]::Round((($fso.GetFolder($scriptParentPath).Size) / 1MB))
            }
        }
        Process {
            [string]$dirTempSupportFiles = "$dirAppDeployTemp\${appVendor}_${appName}_${appVersion}_${appArch}_${appLang}_${appRevision}"
            #Copy toolkit to tmp folder accessable by the user
            if(-not (Test-Path $dirTempSupportFiles)){New-Item $dirTempSupportFiles -ItemType Directory -Force}
            Copy-Item -Path "$dirSupportFiles\closeapps\*" -Destination $dirTempSupportFiles -Recurse -Force
            
            #wrap the parameters in a file
            $WelcomeOptions = @{
                "CloseApps" = $CloseApps;
                "Silent" = $Silent;
                "CloseAppsCountdown" = $CloseAppsCountdown;
                "ForceCloseAppsCountdown" = $ForceCloseAppsCountdown;
                "PromptToSave" = $PromptToSave;
                "PersistPrompt" = $PersistPrompt;
                "AllowDefer" = $AllowDefer;
                "AllowDeferCloseApps" = $AllowDeferCloseApps;
                "DeferTimes" = $DeferTimes;
                "DeferDays" = $DeferDays;
                "DeferDeadline" = $DeferDeadline;
                "CheckDiskSpace" = $CheckDiskSpace;
                "RequiredDiskSpace" = $RequiredDiskSpace;
                "MinimizeWindows" = $MinimizeWindows;
            }
            $WelcomeOptions|Export-Clixml "$dirTempSupportFiles\WelcomeOptions.xml"

            $appOptions = @{
                "appVendor" = $appVendor;
                "appName" = $appName;
                "appVersion" = $appVersion;
                "appArch" = $appArch;
                "appLang" = $appLang;
                "appRevision" = $appRevision;
                "appScriptVersion" = $appScriptVersion;
                "appScriptDate" = $appScriptDate;
                "appScriptAuthor" = $appScriptAuthor;
            }
            $appOptions|Export-Clixml "$dirTempSupportFiles\appOptions.xml"
            #Start at scheduled task as the user
            Write-log -Message "Trying to show a welcome dialog to the user" -Source ${cmdletname}
            $ExitCode = Execute-ProcessAsUser -Path "$dirTempSupportFiles\Deploy-Application.EXE" -Parameters "-DeploymentType Install -DeployMode Interactive" -Wait -PassThru -RunLevel 'LeastPrivilege'
            
            #Get the result of the dialog
            if(Test-Path "$dirTempSupportFiles\result.xml"){
                $results = Import-Clixml "$dirTempSupportFiles\result.xml"
                $action = $results.action
                $global:configInstallationDeferTime = $results.deferHours
                $deferTimes = $results.deferTimes
                $deferDeadlineUniversal = $results.deferDeadlineUniversal

                #Update the defer history
                if($action -eq 'defer'){
                    if($DeferTimes -or $DeferDeadlineUniversal){
                        set-deferHistory -DeferTimesRemaining $DeferTimes -DeferDeadline $deferDeadlineUniversal
                        
                    }
                }

                #Make sure defer or timeout was the action
                if($action -eq 'defer' -or $action -eq 'timeout'){
                    #Create a scheduled task to run the app eval cycle
                    Trigger-AppEvalCycle -Time $global:configInstallationDeferTime
                    Exit-Script -ExitCode $configInstallationDeferExitCode
                }

            }
            #The exit code indicate that we should defer, but there hasn't been saved any information about why. 
            #This would be the case of a active powerpoint presentation
            if($ExitCode -eq $configInstallationDeferExitCode -or $ExitCode -eq $configInstallationUIExitCode){
                Trigger-AppEvalCycle -Time $global:configInstallationDeferTime
                Exit-Script -ExitCode $ExitCode
            }

            #Clean up the temp files
            Remove-Item $dirTempSupportFiles -Recurse -Force
            return $ExitCode

        }
        End {
            Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
        }
    }
#endregion

#region Function Show-InstallationRestartPrompt
Function Show-InstallationRestartPrompt {
    <#
    .SYNOPSIS
        Displays a restart prompt with a countdown to a forced restart.
    .DESCRIPTION
        Displays a restart prompt with a countdown to a forced restart.
    .PARAMETER CountdownSeconds
        Specifies the number of seconds to countdown before the system restart.
    .PARAMETER CountdownNoHideSeconds
        Specifies the number of seconds to display the restart prompt without allowing the window to be hidden.
    .PARAMETER NoCountdown
        Specifies not to show a countdown, just the Restart Now and Restart Later buttons.
        The UI will restore/reposition itself persistently based on the interval value specified in the config file.
    .EXAMPLE
        Show-InstallationRestartPrompt -Countdownseconds 600 -CountdownNoHideSeconds 60
    .EXAMPLE
        Show-InstallationRestartPrompt -NoCountdown
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullorEmpty()]
        [int32]$CountdownSeconds = 60,
        [Parameter(Mandatory=$false)]
        [ValidateNotNullorEmpty()]
        [int32]$CountdownNoHideSeconds = 30,
        [Parameter(Mandatory=$false)]
        [switch]$NoCountdown = $false
    )
        
        Begin {
            ## Get the name of this function and write header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
        }
        Process {
            [string]$dirTempSupportFiles = "$dirAppDeployTemp\${appVendor}_${appName}_${appVersion}_${appArch}_${appLang}_${appRevision}"
            #Copy toolkit to tmp folder accessable by the user
            if(-not (Test-Path $dirTempSupportFiles)){New-Item $dirTempSupportFiles -ItemType Directory -Force}
            Copy-Item -Path "$dirSupportFiles\Reboot\*" -Destination $dirTempSupportFiles -Recurse -Force
            
            $appOptions = @{
                "appVendor" = $appVendor;
                "appName" = $appName;
                "appVersion" = $appVersion;
                "appArch" = $appArch;
                "appLang" = $appLang;
                "appRevision" = $appRevision;
                "appScriptVersion" = $appScriptVersion;
                "appScriptDate" = $appScriptDate;
                "appScriptAuthor" = $appScriptAuthor;
            }
            $appOptions|Export-Clixml "$dirTempSupportFiles\appOptions.xml"

            $rebootOptions = @{
                "CountdownSeconds" = $CountdownSeconds;
                "CountdownNoHideSeconds" = $CountdownNoHideSeconds;
                "NoCountdown" = $NoCountdown;
            }
            $rebootOptions|Export-Clixml "$dirTempSupportFiles\rebootOptions.xml"
            #Start at scheduled task as the user
            Write-log -Message "Trying to show a reboot dialog to the user" -Source ${cmdletname}
            $ExitCode = Execute-ProcessAsUser -Path "$dirTempSupportFiles\Deploy-Application.EXE" -Parameters "-DeploymentType Install -DeployMode Interactive" -wait -PassThru -RunLevel 'LeastPrivilege'
            
            #Clean up the temp files
            #Remove-Item $dirTempSupportFiles -Recurse -Force
            return $ExitCode

        }
        End {
            Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
        }
    }
#endregion

##*===============================================
##* END FUNCTION LISTINGS
##*===============================================

##*===============================================
##* SCRIPT BODY
##*===============================================

If ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $appDeployToolkitExtName
}
Else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $appDeployToolkitExtName
}

##*===============================================
##* END SCRIPT BODY
##*===============================================