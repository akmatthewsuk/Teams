Param(
	[Parameter(Mandatory=$true)]    
    [string]$ConfigFile
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
Filename:      New-DeployToast.ps1
Documentation: https://tothecloudandbeyond.substack.com/
Execution Tested on: Windows 10 2009
Requires:      Setup as a scheduled task
Versions:
1.0 - Intial Release
 - First release created using old code
===========================================================================
.SYNOPSIS

Displays Toast Notifications during Autopilot Deployment

.DESCRIPTION
Section 0 - Script initialisation
Section 1 - Update the teams backgrounds if required
Section 2 - Close Out

.INPUTS
The execution engine is controled by a config file (config.xml)

.OUTPUTS
Outputs a log file in CMTrace format
#>

################################################
#Declare Constants and other Script Variables
################################################

#Log Levels
[string]$LogLevelError = "Log_Error"
[string]$LogLevelWarning = "Log_Warning"
[string]$LogLevelInfo = "Log_Information"

#[string]$LogPath = "C:\Program Files\Deploy\DeployLog"
#use the $env:TEMP folder as the log location
[string]$LogPath = "$((join-path -path $env:userprofile -childpath "appdata\local\temp\TeamsBackgrounds"))"
[string]$TxtLogfilePrefix = "TeamsBackgrounds" # Log file in cmtrace format

$LogCacheArray = New-Object System.Collections.ArrayList
$MaxLogCachesize = 10
$MaxLogWriteAttempts = 5

################################################
#Declare Functions
################################################

<# Create a New log entry in log files #>
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
    Write-Host "$($LogEntry) $($ScriptLineNumber)"
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

Function Write-LogEntry {
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
    #Determine the action based on the log level
    switch ($LogLevel) {
        $LogLevelError {  
            #Create the CMTrace Log Line
            $TXTLogLine = '<![LOG[' + $LogEntry + ']LOG]!><time="' + $TxtLogTime + '" date="' + $TxtLogDate + '" component="' + "$($ScriptLineNumber)" + '" context="" type="' + 3 + '" thread="" file="">'
        }
        $LogLevelWarning {
            $TXTLogLine = '<![LOG[' + $LogEntry + ']LOG]!><time="' + $TxtLogTime + '" date="' + $TxtLogDate + '" component="' + "$($ScriptLineNumber)" + '" context="" type="' + 2 + '" thread="" file="">'
        }
        $LogLevelInfo {
            $TXTLogLine = '<![LOG[' + $LogEntry + ']LOG]!><time="' + $TxtLogTime + '" date="' + $TxtLogDate + '" component="' + "$($ScriptLineNumber)" + '" context="" type="' + 1 + '" thread="" file="">'
        }
        default {
            $TXTLogLine = '<![LOG[' + $LogEntry + ']LOG]!><time="' + $TxtLogTime + '" date="' + $TxtLogDate + '" component="' + "$($ScriptLineNumber)" + '" context="" type="' + 1 + '" thread="" file="">'
        }
    }

    #Write the CMTrace Log line
    Add-Content -Value $TXTLogLine -Path $TxtLogFile -force
}

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
        } Catch {
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
    New-Item -Path $NewLogfile -Type File -force | Out-Null

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

<# 
    Main routine
#>


################################################
# SECTION 1: Script Initialization
################################################

# SECTION 1 STEP 1: Create a Log file
New-LogEntry -LogEntry "Starting Teams Backgrounds Updater"
New-LogEntry -LogEntry "### Section 1: Script Initialisation ###"
New-LogEntry -LogEntry "*** Step 1: Creating Log File ***"
$TxtLogFile = New-TxtLog -NewlogPath $LogPath -NewLogPrefix $TxtLogfilePrefix


# SECTION 1 STEP 2: Load the Config.xml
New-LogEntry -LogEntry "*** Step 2: Loading config ***"
New-LogEntry -LogEntry "Loading configuration file: $($ConfigFile)"
try {
    [Xml]$config = Get-Content $ConfigFile
} catch {
    $ErrorMessage = $_.Exception.Message
    New-LogEntry -LogEntry "Error loading the config XML" -LogLevel $LogLevelError 
    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError -FlushLogCache $true
    Exit-Script -ExitText "Unable to load the Config XML - Script exiting"
}

# SECTION 1 STEP 3: Obtain the current logged in user
New-LogEntry -LogEntry "*** Step 3: Checking Logged In User ***"
#Grab the logged in user from WMI
Try {
    $LoggedInUserWMI = Get-WmiObject -class win32_computersystem | select-object username
   
} Catch {
    $ErrorMessage = $_.Exception.Message
    New-LogEntry -LogEntry "Error retrieving Logged in User from WMI" -LogLevel $LogLevelError
    New-LogEntry -LogEntry $ErrorMessage -LogLevel $LogLevelError
}
New-LogEntry -LogEntry "Logged in user account: $($LoggedInUserWMI.Username)"
#If the username is not a valid username then error handling is required
If($null -eq $LoggedInUserWMI.Username) {
    New-LogEntry -LogEntry "No user is currently logged in" -LogLevel $LogLevelWarning
    Exit-Script -ExitText "No user is currently logged in - Toast notification cannot be sent"
} else {
    If ($LoggedInUserWMI.Username -like "*\*")
    {
        $LoggedinUser = $LoggedInUserWMI.username.Split("\")
        #Get the Profile path of the User
        $LoggedinUserProfilePath = "c:\users\" + $LoggedinUser[1]
    } else {
        New-LogEntry -LogEntry "Logged in user format not recognized" -LogLevel $LogLevelError
        Exit-Script "Error checking logged in user"
    }
}

################################################
# Section 2: Update Backgrounds
################################################
New-LogEntry -LogEntry "*** Section 2 Step 1: Checking Teams Background Folder ***"
#Section 2 Step 1: Get the Teams Backgrounds folder from the config
If(!($config.config.teams.backgroundpath -eq 0))  {
    $TeamsBackgroundsFolder = $config.config.teams.backgroundpath
    #Check the folder exists
    If(Test-Path -Path $TeamsBackgroundsFolder -PathType Container) {
        New-LogEntry -LogEntry "Using Teams Background folder $($TeamsBackgroundsFolder)"
    } else {
        New-LogEntry -LogEntry "The Teams Background Folder specified in the config ($($TeamsBackgroundsFolder)) does not exist" -LogLevel $LogLevelError
        Exit-Script "Teams Background folder is missing"
    }
} else {
    New-LogEntry -LogEntry "The Teams Background Folder is missing from the config" -LogLevel $LogLevelError
    Exit-Script "Error checking the Teams Background folder in the config"
}

#Check whether the Teams backgrounds folder exists
New-LogEntry -LogEntry "*** Section 2 Step 2: Checking Teams Background Folder ***"
$TeamsBackgroundsUserFolder = Join-Path -Path $LoggedinUserProfilePath -ChildPath "\appdata\Roaming\Microsoft\Teams\Backgrounds\Uploads"
If (!(Test-Path -path $TeamsBackgroundsUserFolder -PathType Container)) {
    New-LogEntry -LogEntry "The Teams Backgrounds Folder ($TeamsBackgroundsUserFolder) does not exist"
    New-LogEntry -LogEntry "Creating Teams Backgrounds folder in the user profile"
    New-Item -Path $TeamsBackgroundsfolder -ItemType "directory"
} else {
    New-LogEntry -LogEntry "The Teams Backgrounds Folder ($TeamsBackgroundsUserFolder ) exists"
}

#Loop through the backgrounds and process the background actions
New-LogEntry -LogEntry "*** Section 2 Step 3: Checking Teams Background Folder ***"
Foreach($Background in $config.config.teams.backgrounds.background) {

    #Create the source and destination files
    $SourceFile = Join-Path -Path $TeamsBackgroundsFolder -ChildPath $Background.Image
    $TargetFile = Join-Path -Path $TeamsBackgroundsUserFolder -ChildPath $Background.Image
    #Determine the action based on the config switch
    Switch ($Background.Action) {
        "add" {
            New-LogEntry -LogEntry "Processing Add Action for $($Background.Image)"
            #Check whether the background file exists in the Source folder
            If(Test-Path -Path $SourceFile -PathType Leaf) {
                New-LogEntry -LogEntry "Copying the Source file ($($SourceFile)) to $TeamsBackgroundsUserFolder"
                Copy-Item -Path $SourceFile -Destination $TeamsBackgroundsUserFolder -Force
            } else {
                New-LogEntry -LogEntry "The source file ($($SourceFile)) does not exist - ignoring background" -LogLevel $LogLevelWarning
            }
        }
        "remove" {
            #Check whether the background file exists in the target folder
            New-LogEntry -LogEntry "Processing Remove Action for $($Background.Image)"
            If(Test-Path -Path $TargetFile -PathType Leaf) {
                New-LogEntry -LogEntry "Removing the Target file ($($Background.Image)) from $TeamsBackgroundsUserFolder"
                Remove-Item -Path $TargetFile -Force
            } else {
                New-LogEntry -LogEntry "The Target file ($($TargetFile)) does not exist - ignoring background" -LogLevel $LogLevelWarning
            }
        }
        default {
            New-LogEntry -LogEntry "Action type ($($Background.Action)) unknown - ignoring background" -LogLevel $LogLevelWarning
        }
    }

    

    

}

################################################
# Section 3: Exit
################################################
New-LogEntry -LogEntry  "### Section 3: Graceful Exit ###"
Exit-Script -ExitText "script complete"
