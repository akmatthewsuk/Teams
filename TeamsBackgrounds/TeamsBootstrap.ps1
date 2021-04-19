Param(
	$installFolder
)

<#  
    .NOTES
===========================================================================
    ## License ##    

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details. 

    
===========================================================================
Created by:    Andrew Matthews
Organization:  To The Cloud And Beyond
Filename:      TeamsBootstrap.ps1
Documentation: https://tothecloudandbeyond.substack.com/
Execution Tested on: Windows 10 2009
Requires:      Installation with an MSI package
Versions:
1.0 Initial Release
 -

===========================================================================
.SYNOPSIS

The bootstrap for the Teams Background updater 

.DESCRIPTION
The bootstrap performs the following actions.
# Section 1 
 - Create a log file
 - load the config
# Section 2
 - Create the scheduled task
 - Run the scheduled task


.INPUTS
The bootstrap is controled by a config file (config.xml)

.OUTPUTS
Log file output in cmtrace format
#>

################################################
#Declare Constants and other Script Variables
################################################

#Log Levels
[string]$LogLevelError = "Log_Error"
[string]$LogLevelWarning = "Log_Warning"
[string]$LogLevelInfo = "Log_Information"

[string]$LogPath = "C:\Program Files\Deploy\Log"
[string]$TxtLogfilePrefix = "TeamsBoostrap" # Log file in cmtrace format

$LogCacheArray = New-Object System.Collections.ArrayList
$MaxLogCachesize = 10
$MaxLogWriteAttempts = 5

$TaskExecutableType_WScript = "WScript"
$TaskExecutableType_PowerShell = "PowerShell"

$WScriptpath = "C:\WINDOWS\system32\wscript.exe"
$PSpath = Join-Path -Path "C:\Windows\System32\WindowsPowerShell\v1.0" -ChildPath "powershell.exe"

################################################
#Declare Functions
################################################
<# Create a New log entry and invoke the log cache flush if required #>
Function New-LogEntry {
    param (
        [Parameter(Mandatory=$true)]    
        [string]$LogEntry,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Log_Error","Log_Warning","Log_Information")]
        [string]$LogLevel,
        [Parameter(Mandatory=$false)]
        [Bool]$ImmediateLog,
        [Parameter(Mandatory=$false)]
        [Bool]$FlushLogCache
    )

    #Create the CMTrace Time stamp
    $TxtLogTime = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $TxtLogDate = "$(Get-Date -Format MM-dd-yyyy)"

    #Create the Script line number variable
    $ScriptLineNumber = "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"

    #Add the log entry to the cache
    switch ($LogLevel) {
        $LogLevelError {  
            New-LogCacheEntry -LogEntry $LogEntry -LogTime $TxtLogTime -LogDate $TxtLogDate -ScriptLineNumber $ScriptLineNumber -LogLevel $LogLevel
        }
        $LogLevelWarning { 
            New-LogCacheEntry -LogEntry $LogEntry -LogTime $TxtLogTime -LogDate $TxtLogDate -ScriptLineNumber $ScriptLineNumber -LogLevel $LogLevel
        }
        $LogLevelInfo { 
            New-LogCacheEntry -LogEntry $LogEntry -LogTime $TxtLogTime -LogDate $TxtLogDate -ScriptLineNumber $ScriptLineNumber -LogLevel $LogLevel
        }
        default {
            New-LogCacheEntry -LogEntry $LogEntry -LogTime $TxtLogTime -LogDate $TxtLogDate -ScriptLineNumber $ScriptLineNumber -LogLevel $LogLevelInfo
        }
    }

    #Set the Write log entries to the default state of false
    $WriteLogEntries = $True
    #Determine whether the log needs to be immediately written
    If ($PSBoundParameters.ContainsKey('ImmediateLog')) {
        If($ImmediateLog -eq $false) {
            #Do not invoke the log flush       
        } Else {
            #If the action is immediate log then flush the log entries
            $WriteLogEntries = $True
        }
    } else {
        #If no value specified then for not flush the log cache
        $WriteLogEntries = $false
    }

    If ($PSBoundParameters.ContainsKey('FlushLogCache')) { 
        If($FlushLogCache -eq $false) {
            If($LogCacheArray.count -eq $MaxLogCachesize) {
                #If the max cache size has been hit then flush the log entries
                $WriteLogEntries = $true
            }
        } else { 
            $WriteLogEntries = $true
        }
    } else {
        If($LogCacheArray.count -eq $MaxLogCachesize) {
            #If the max cache size has been hit then flush the log entries
            $WriteLogEntries = $true
        }
    }


    If ($WriteLogEntries -eq $true) {
        #write the log entries
        Write-LogEntries
    }
}

<# Write the log entries to the log file #>
Function Write-LogEntries {
    Write-Host "**** Flushing $($LogCacheArray.count) Log Cache Entries ****"
    $LogTextRaw = ""
    #Rotate through the Log entries and compile a master variable
    ForEach($LogEntry in $LogCacheArray) {
        switch ($LogEntry.LogLevel) {
            $LogLevelError {  
                #Create the CMTrace Log Line
                $TXTLogLine = '<![LOG[' + $LogEntry.LogEntry + ']LOG]!><time="' + $LogEntry.LogTime + '" date="' + $LogEntry.LogDate + '" component="' + "$($LogEntry.LineNumber)" + '" context="" type="' + 3 + '" thread="" file="">'
            }
            $LogLevelWarning {
                $TXTLogLine = '<![LOG[' + $LogEntry.LogEntry + ']LOG]!><time="' + $LogEntry.LogTime + '" date="' + $LogEntry.LogDate + '" component="' + "$($LogEntry.LineNumber)" + '" context="" type="' + 2 + '" thread="" file="">'
            }
            $LogLevelInfo {
                $TXTLogLine = '<![LOG[' + $LogEntry.LogEntry + ']LOG]!><time="' + $LogEntry.LogTime + '" date="' + $LogEntry.LogDate + '" component="' + "$($LogEntry.LineNumber)" + '" context="" type="' + 1 + '" thread="" file="">'
            }
            default {
                $TXTLogLine = '<![LOG[' + $LogEntry.LogEntry + ']LOG]!><time="' + $LogEntry.LogTime + '" date="' + $LogEntry.LogDate + '" component="' + "$($LogEntry.LineNumber)" + '" context="" type="' + 1 + '" thread="" file="">'
            }
        }
        If($LogTextRaw.Length -eq 0) {
            $LogTextRaw = $TXTLogLine
        } else {
            $LogTextRaw = $LogTextRaw + "`r`n" + $TXTLogLine
        }
    }

    #Write the Log entries Log line
    $LogWritten = $false
    $LogWriteAttempts = 0
    do {
        $LogWriteAttempts = $LogWriteAttempts + 1
        $WriteLog = $True
        Try {
            Add-Content -Value $LogTextRaw -Path $TxtLogFile -ErrorAction Stop
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $WriteLog = $false
            Write-Host "Log entry flush failed"
            Write-Host $ErrorMessage
        }
        If ($WriteLog-eq $false) {
            If ($LogWriteAttempts -eq $MaxLogWriteAttempts) {
                Write-Host "Maximum log write attempts exhausted - saving log entries for the next attempt"
                $LogWritten = $true
            }
            #Wait five seconds before looping again
            Start-Sleep -Seconds 5
        } else {
            $LogWritten = $true
            Write-Host "Wrote $($LogCacheArray.count) cached log entries to the log file"
            $LogCacheArray.Clear()
        }
    } Until ($LogWritten -eq $true) 
        
}

<# Create a new entry in the log cache #>
Function New-LogCacheEntry {
    param (
        [Parameter(Mandatory=$true)]    
        [string]$LogEntry,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Log_Error","Log_Warning","Log_Information")]
        [string]$LogLevel,
        [Parameter(Mandatory=$true)]
        [string]$LogTime,
        [Parameter(Mandatory=$true)]
        [string]$LogDate,
        [Parameter(Mandatory=$true)]
        [string]$ScriptLineNumber
    )

    $LogCacheEntry = New-Object -TypeName PSObject -Property @{
        'LogEntry' = $LogEntry
        'LogLevel' = $LogLevel
        'LogTime' = $LogTime
        'LogDate' = $LogDate
        'Linenumber' = $ScriptLineNumber
    }

    $LogCacheArray.Add($LogCacheEntry) | Out-Null

}

<# Create a new log file for a Txt Log #>
Function New-TxtLog {
    param (
        [Parameter(Mandatory=$true)]    
        [string]$NewLogPath,
        [Parameter(Mandatory=$true)]    
        [string]$NewLogPrefix
    )

    #Create the log path if it does not exist
    if (!(Test-Path $NewLogPath))
    {
        New-Item -itemType Directory -Path $NewLogPath
    }

    #Create the new log name using the prefix
    [string]$NewLogName = "$($NewLogPrefix)-$(Get-Date -Format yyyy-MM-dd)-$(Get-Date -Format HH-mm).log"
    #Create the fill path
    [String]$NewLogfile = Join-Path -path $NewLogPath -ChildPath $NewLogName
    #Create the log file
    New-Item -Path $NewLogPath -Name $NewLogName -Type File -force | Out-Null

    #Return the LogfileName
    Return $NewLogfile
}

<# Exit the script#>
Function Exit-Script {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ExitText
    )
    
    #Write the exit text to the log, flush the log cache and exit
    New-LogEntry -LogEntry $ExitText -FlushLogCache $true
    Exit
}

################################################
# SECTION 1: Script Initialization
################################################

# SECTION 1 STEP 1: Create a Log file

$ReturnedLogFile = New-TxtLog -NewlogPath $LogPath -NewLogPrefix $TxtLogfilePrefix
If(($ReturnedLogFile | Measure-Object).count -eq 1) {
    $TxtLogfile = $ReturnedLogFile
    New-LogEntry -LogEntry "Writing Log file to $($TxtLogfile)"
} else {
    Foreach($file in $ReturnedLogFile) {
        #Workaround for the returned value being returned as an array object
        New-LogEntry -LogEntry "Checking that the log file $($file) exists"
        if(test-path -Path $file -PathType Leaf) {
            $TxtLogfile = $file
            New-LogEntry -LogEntry "Writing Log file to $($TxtLogfile)"
        }
    }
}

# SECTION 1 STEP 2: Load the Config.xml
New-LogEntry -LogEntry "Install folder location: $installFolder"
New-LogEntry -LogEntry "Loading configuration location: $($installFolder)Config.xml"
try {
    [Xml]$config = Get-Content "$($installFolder)Config.xml"
} catch {
    $ErrorMessage = $_.Exception.Message
    New-LogEntry -LogEntry "Error loading the config XML" -LogLevel $LogLevelError 
    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError -FlushLogCache $true
    Exit-Script -ExitText "Unable to load the Config XML - Script exiting"
}

################################################
# SECTION 2: Configure Scheduled Task
################################################

#SECTION 2 STEP 1: Creating Scheduled Tasks
New-LogEntry -LogEntry "*** Creating Scheduled Tasks ***"
ForEach ($Task in $Config.Config.Tasks.Task) {
    New-LogEntry -LogEntry "+++ Processing Task - $($Task.ID) +++"
        
    #Set the process scheduled task flag
    $ProcessScheduledTask = $True
    #Set the replace Scheduled task flag
    $ReplaceScheduledTask = $False #if true at the end of the processing then an existing task will be replaced

    #Set the Task Name
    $TaskName = $Task.TaskName
    New-LogEntry -LogEntry "TaskName: $($TaskName)"

    #If the Task does not already exist then create the task otherwise check whether the task needs to be recreated (I.E. the script has changed)
    if(!(Get-ScheduledTask | where-object {$_.TaskName -eq $TaskName})) {
        $ScheduledTaskExists = $False
        New-LogEntry -logentry "Scheduled Task $($TaskName) does not exist - proceeding to register the scheduled task" 
    } else {
        $ScheduledTaskExists = $True
        $ExistingTask = Get-ScheduledTask | where-object {$_.TaskName -eq $TaskName}
        New-LogEntry -logentry "Scheduled Task $($TaskName) already exists - Checking task" -LogLevel $LogLevelWarning 
    }

    #Construct the path to the task script and confirm the script file exists
    If ($ProcessScheduledTask -eq $true) {
                
        $TaskScript = Join-Path -Path $Task.TaskScriptFolder -ChildPath $Task.TaskScript

        #Check that the Scheduled task script exists
        If (Test-Path -Path $TaskScript -PathType Leaf) {
            New-LogEntry -LogEntry "Task Script: $($TaskScript)"
        } else {
            $ProcessScheduledTask = $False
            New-LogEntry -LogEntry "Task Script ($($TaskScript)) not found" -LogLevel $LogLevelError
        }
    }

    #Determine whether this is a PowerShell or wscript action
    If(!($Task.TaskExecutable.Length -eq 0)) {
        switch ($Task.TaskExecutable) {
            "powershell" {
                $TaskExecutableType = $TaskExecutableType_PowerShell
            }
            "wscript" {
                $TaskExecutableType = $TaskExecutableType_WScript
            }
            Default {
                $ProcessScheduledTask = $False
                New-LogEntry -LogEntry "Task Excutable Type ($($Task.TaskExecutable)) was unknown" -LogLevel $LogLevelError
            }
        }
    } else {
        #Default to PowerShell
        $TaskExecutableType = $TaskExecutableType_PowerShell
    }

    #Create the arguments for the PowerShell Task action
    If ($ProcessScheduledTask -eq $true) {
        If($TaskExecutableType -eq $TaskExecutableType_PowerShell) {
            #Set the Task Argument String
            $TaskArgument = '-WindowStyle Hidden -NonInteractive -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "' + $TaskScript + '"'
                
            #Add the Task Config file if it exists
            If (!($Task.TaskConfig.Length -eq 0)) {
                $TaskConfigPath = Join-Path -Path $Task.TaskScriptFolder -ChildPath $Task.TaskConfig
                New-LogEntry -LogEntry "Task config: $($TaskConfigPath)" 
                $TaskArgument = $TaskArgument + ' -Config "' + $TaskConfigPath + '"'
            }
            #Add other task arguments
            If (!($Task.TaskArgument.Length -eq 0)) {
                $TaskArgument = $TaskArgument + " " + $Task.TaskArgument
            }
        } elseif ($TaskExecutableType = $TaskExecutableType_WScript) {
            $TaskArgument = '"' + $TaskScript + '"'
        }
    }

    #Check whether the current task argument matches the current arguments
    If ($ProcessScheduledTask -eq $true) {
            
        New-LogEntry -LogEntry "Task Argument String: $($TaskArgument)"
        If($ScheduledTaskExists -eq $true) {
            #Check whether the task action argument matches the required task
            New-LogEntry -logentry "Checking the arguments for the existing task $($ExistingTask.TaskName)"
            $ActionArguments = $false
            $ActionExecute = $false
            #Note this assumes that there is only one action but because the returned value is an array, all the actions need to be checked.
            Foreach ($Action in $ExistingTask.Actions) {
                New-LogEntry -logentry "Exsting argument $($Action.Arguments)"
                New-LogEntry -logentry "New Argument $($TaskArgument)"
                #compare the arguments
                If ($Action.Arguments.tolower() -eq $TaskArgument.ToLower()) {
                    $ActionArguments = $True
                    New-LogEntry -logentry "Match found for existing arguments"
                }
                New-LogEntry -logentry "Exsting executed process $($Action.Execute)"
                If($TaskExecutableType -eq $TaskExecutableType_PowerShell) {
                    New-LogEntry -logentry "New executed process $($PSpath)"
                    #Compare the executed process
                    If ($Action.Execute.tolower() -eq $PSpath.ToLower()) {
                        $ActionExecute = $True
                        New-LogEntry -logentry "Match found for existing execution"
                    }
                } elseif ($TaskExecutableType = $TaskExecutableType_WScript) {
                    New-LogEntry -logentry "New executed process $($WScriptpath)"
                    #Compare the executed process
                    If ($Action.Execute.tolower() -eq $WScriptpath.ToLower()) {
                        $ActionExecute = $True
                        New-LogEntry -logentry "Match found for existing execution"
                    }
                }
            }
            If (($ActionArguments -eq $false) -or ($ActionExecute -eq $false)) {
                $ReplaceScheduledTask = $True
                New-LogEntry -logentry "Arguments for Task $($ExistingTask.TaskName) have changed - task must be re-registered" -LogLevel $LogLevelWarning
            } else {
                New-LogEntry -logentry "Arguments for Task $($ExistingTask.TaskName) have not changed - task does not need to be re-registered"
            }

        }
    }

    If ($ProcessScheduledTask -eq $true) {
        If($TaskExecutableType -eq $TaskExecutableType_PowerShell) {
            #Set the Task Action - This action requires a try catch block because it can fail
            New-LogEntry -LogEntry "Creating Scheduled Task Action Object for a PowerShell action"            
            try {
                $TaskAction = New-ScheduledTaskAction -execute $PSpath -Argument $TaskArgument
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Action Object" -LogLevel $LogLevelError
                New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                $ProcessScheduledTask = $False
            }
        } elseif ($TaskExecutableType = $TaskExecutableType_WScript) {
            #Set the Task Action - This action requires a try catch block because it can fail
            New-LogEntry -LogEntry "Creating Scheduled Task Action Object for a WScript Action"            
            try {
                $TaskAction = New-ScheduledTaskAction -execute $WScriptpath -Argument $TaskArgument
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Action Object" -LogLevel $LogLevelError
                New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                $ProcessScheduledTask = $False
            }
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        #Set the Task triggers
        New-LogEntry -LogEntry "Creating Task Trigger Array"
        #Create the task trigger array
        $TaskTriggerArray = New-Object System.Collections.ArrayList
        ForEach ($Run in $Task.TaskRun.Run) {
            switch ($Run.RunType) {
                "Daily" {
                    New-LogEntry -LogEntry "Creating Task Trigger: Daily at $($Run.RunTime)"   
                    Try {
                        $TaskTrigger = New-ScheduledTaskTrigger -Daily -At "$($Run.RunTime)" -DaysInterval 1
                    } catch {
                        $ErrorMessage = $_.Exception.Message
                        New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Trigger" -LogLevel $LogLevelError
                        New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                        $ProcessScheduledTask = $False
                        Break
                    }
                    $TaskTriggerArray.Add($TaskTrigger)
                }
                "Weekly" {
                    New-LogEntry -LogEntry "Creating Task Trigger: Weekly on $($Run.RunDay) at $($Run.RunTime)"    
                    Try {
                        $TaskTrigger = New-ScheduledTaskTrigger -Weekly -At "$($Run.RunTime)" -WeeksInterval 1 -DaysOfWeek $Run.RunDay
                    } catch {
                        $ErrorMessage = $_.Exception.Message
                        New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Trigger" -LogLevel $LogLevelError
                        New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                        $ProcessScheduledTask = $False
                        Break
                    }
                    $TaskTriggerArray.Add($TaskTrigger)
                }
                "AtLogon" {
                    New-LogEntry -LogEntry "Creating Task Trigger: At Logon" 
                    Try {
                        $TaskTrigger = New-ScheduledTaskTrigger -AtLogOn
                    } catch {
                        $ErrorMessage = $_.Exception.Message
                        New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Trigger" -LogLevel $LogLevelError
                        New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                        $ProcessScheduledTask = $False
                        Break
                    }
                    $TaskTriggerArray.Add($TaskTrigger)
                }
                "NetConnect" {
                    New-LogEntry -LogEntry "Creating Task Trigger: On Network Connection" 
                    #Create a custom task trigger based on a subscription
                    $TaskTriggerOK = $True
                    If($TaskTriggerOK -eq $true) {
                        New-LogEntry -LogEntry "Creating custom Task Trigger class"
                        Try{
                            $TaskTriggerClass = get-cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
                        } catch {
                            $ErrorMessage = $_.Exception.Message
                            New-LogEntry -LogEntry "Error occurred when creating custom Task Trigger class" -LogLevel $LogLevelError
                            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                            $TaskTriggerOK = $False
                        }
                    }

                    If($TaskTriggerOK -eq $true) {
                        New-LogEntry -LogEntry "Creating custom Task Trigger object"
                        Try{
                            $TaskTrigger = $TaskTriggerClass | New-CimInstance -ClientOnly
                        } catch {
                            $ErrorMessage = $_.Exception.Message
                            New-LogEntry -LogEntry "Error occurred when creating custom Task Trigger object" -LogLevel $LogLevelError
                            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                            $TaskTriggerOK = $False
                        }
                    }

                    If($TaskTriggerOK -eq $true) {
                        New-LogEntry -LogEntry "Enabling task trigger"
                        Try{
                            $TaskTrigger.Enabled = $true
                        } catch {
                            $ErrorMessage = $_.Exception.Message
                            New-LogEntry -LogEntry "Error occurred enabling the task trigger" -LogLevel $LogLevelError
                            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                            $TaskTriggerOK = $False
                        }
                    }

                    If($TaskTriggerOK -eq $true) {
                        New-LogEntry -LogEntry "Creating Task subscription for Network EventID 10000"
                        Try{
                            $TaskTrigger.Subscription = '<QueryList><Query Id="0" Path="Microsoft-Windows-NetworkProfile/Operational"><Select Path="Microsoft-Windows-NetworkProfile/Operational">*[System[Provider[@Name=''Microsoft-Windows-NetworkProfile''] and EventID=10000]]</Select></Query></QueryList>'
                        } catch {
                            $ErrorMessage = $_.Exception.Message
                            New-LogEntry -LogEntry "Error occurred when creating Task subscription" -LogLevel $LogLevelError
                            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                            $TaskTriggerOK = $False
                        }
                    }
                    
                    If($TaskTriggerOK -eq $false) {
                        New-LogEntry -LogEntry "Error detected when creating the event subscription trigger"
                        $ProcessScheduledTask = $False
                    } else {
                        $TaskTriggerArray.Add($TaskTrigger)
                    }
                }
                "None"{
                    New-LogEntry -LogEntry "The run type is set to none - do not create a task schedule item"
                }
                Default {
                    #Default to Once in ten minutes
                    New-LogEntry -LogEntry "Task run Type ($($Run.RunType)) is unknown" -LogLevel $LogLevelWarning
                    New-LogEntry -LogEntry "Creating Task Trigger: Once" 
                    Try {
                        $TaskTrigger = New-ScheduledTaskTrigger -Once -At (Get-date).AddMinutes(10)
                    } catch {
                        $ErrorMessage = $_.Exception.Message
                        New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Trigger" -LogLevel $LogLevelError
                        New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                        $ProcessScheduledTask = $False
                        Break
                    }
                    $TaskTriggerArray.Add($TaskTrigger)
                }
            }
        }

        #confirm the size of the task trigger array
        If ($TaskTriggerArray.count -eq 0) {
            New-LogEntry -LogEntry "No Task triggers were added to the array" -LogLevel $LogLevelWarning
            $CreateWithoutTaskTrigger = $True
        } else {
            $CreateWithoutTaskTrigger = $false
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        Switch -exact ($Task.TaskPrincipal) {
            "System" {
                #Set the task principal
                New-LogEntry -LogEntry "Creating Task Principal" 
                Try {
                    $TaskPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount
                } catch {
                    $ErrorMessage = $_.Exception.Message
                    New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Principal" -LogLevel $LogLevelError
                    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                    $ProcessScheduledTask = $False
                }
            }
            "User" {
                #Set the task principal
                New-LogEntry -LogEntry "Creating Task Principal" 
                Try {
                    $TaskPrincipal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Users" -RunLevel Limited
                } catch {
                    $ErrorMessage = $_.Exception.Message
                    New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Principal" -LogLevel $LogLevelError
                    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                    $ProcessScheduledTask = $False
                }
            }
            default {
                #Default to System
                New-LogEntry -LogEntry "Creating Task Principal" 
                Try {
                    $TaskPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount
                } catch {
                    $ErrorMessage = $_.Exception.Message
                    New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Principal" -LogLevel $LogLevelError
                    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                    $ProcessScheduledTask = $False
                }
            }
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        #Creating the Task Settings
        New-LogEntry -LogEntry "Creating Task Settings Set" 
        Try {
            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
        } catch {
            $ErrorMessage = $_.Exception.Message
            New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Setting Set" -LogLevel $LogLevelError
            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
            $ProcessScheduledTask = $False
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        #Create the task object
        New-LogEntry -LogEntry "Creating the Scheduled Task object" 
        Try {
            If ($CreateWithoutTaskTrigger -eq $true) {
                $TaskObject = New-ScheduledTask -Action $TaskAction -Principal $TaskPrincipal -Settings $TaskSettings
            } else {
                $TaskObject = New-ScheduledTask -Action $TaskAction -Trigger $TaskTriggerArray -Principal $TaskPrincipal -Settings $TaskSettings
            }
        } catch {
            $ErrorMessage = $_.Exception.Message
            New-LogEntry -LogEntry "Error occurred creating a Scheduled Task Object" -LogLevel $LogLevelError
            New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
            $ProcessScheduledTask = $False
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        if($ScheduledTaskExists -eq $True){
            If($ReplaceScheduledTask -eq $True) {
                $RegisterScheduleTask = $True
                #Remove the existing scheduled task
                New-LogEntry -logentry "Unregistering the old scheduled task" -LogLevel $LogLevelWarning
                try {
                    Unregister-ScheduledTask -TaskName $ExistingTask.TaskName -TaskPath $ExistingTask.TaskPath -Confirm:$false
                } catch {
                    $ErrorMessage = $_.Exception.Message
                    New-LogEntry -LogEntry "Error occurred unregistering an existing Scheduled Task" -LogLevel $LogLevelError
                    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                    $ProcessScheduledTask = $False
                }
            } else {
                $RegisterScheduleTask = $false
                New-LogEntry -logentry "Existing task does not need to be re-registered"
            }
        } else {
            $RegisterScheduleTask = $True
        }
    }

    If ($ProcessScheduledTask -eq $true) {
        If($RegisterScheduleTask -eq $True) {
            #Registering Scheduled Task
            
            New-LogEntry -LogEntry "Registering Scheduled Task $($TaskName)" 
            Try {
                Register-ScheduledTask $TaskName -InputObject $TaskObject
            } catch {
                $ErrorMessage = $_.Exception.Message
                New-LogEntry -LogEntry "Error occurred registering a Scheduled Task" -LogLevel $LogLevelError
                New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
                $ProcessScheduledTask = $False
            }
        }
    }
    
    If ($ProcessScheduledTask -eq $true) {
        If($Task.StartImmediately.ToLower() -eq "yes") {
            New-LogEntry -LogEntry "Starting $($TaskName) Scheduled Task" 
            Try {
                Start-ScheduledTask -TaskName $TaskName
            } catch {
                $ErrorMessage = $_.Exception.Message
                New-LogEntry -LogEntry "Error occurred starting a Scheduled Task" -LogLevel $LogLevelError
                New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
            }
        }
    }

}



################################################
# FINAL SECTION: Script Exit
################################################

#Flush the log cache before exiting
Exit-Script -ExitText "Script complete"