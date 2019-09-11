#------------------------------[ HELP INFO ]-------------------------------

<#
.SYNOPSIS
Exports, imports, or deletes local group members on a local or remote computer.

.DESCRIPTION
This script can be used to:

- export local groups and/or members from a local or remote computer to a text file,
- import local groups and/or members to a local or remote computer from a text file, or
- delete local groups and/or members specified in a text file on a local or remote computer.

During export, you can filter groups by name and/or description using regular expressions. You can also specify the exclusion lists (also by names and descriptions) using the configuration file. In a similar manner you can specify which group members to include or exclude during the export operation.

The script allows you to specify the non-default parameters and settings via an optional configuration file (see the description of the '-ConfigFile' parameter). Please keep in mind that the command-line parameters override the config file settings.

The script has a couple of external module dependencies:

https://powershellgallery.com/packages/ConfigFile/
https://powershellgallery.com/packages/ScriptVersion/

If the system on which the script runs has limited or does not have access to the Internet, download these packages manually and use the '-ModulePath' to specify location of the root module folder. (You can find a number of articles explaining how to install a PowerShell module without internet access online.)

You can log the information about the performed operations and/or errors to the log and/or error log files.

.PARAMETER Export
Export local groups and/or members to a data file.

.PARAMETER Import
Import local groups and/or members from a data file.

.PARAMETER Delete
Delete local groups and/or members specified in a data file.

.PARAMETER DataFile
Defines the full path to the date file holding imported or exported data. The data file contains three text entries separated by tabs (or explicitly defined separator):

- Name of the local group
- Name of the account in the format: WinNT://DOMAIN/user
- Group description

The group description is only needed the first time the group appears in the file. For groups with no members, the name of the account may be left blank. Here is an example of the data file (with the tab characters identified by '{TAB}'):

Group A{TAB}{TAB}Empty group with no members.
Group B{TAB}WinNT://CONTOSO/jdoe{TAB}Group with two members.
Group B{TAB}WinNT://CONTOSO/jschmoe{TAB}
Group C{TAB}WinNT://CONTOSO/mjane{TAB}Group with one member.

When exporting data, if the 'ExcludeEmptyGroups' switch is not set, all data file entries will contain group members; otherwise, there will be a separate entry for each exported group with no member information.

The group description is needed only the first time the group appears in the data file.

When deleting data, if the data file entry does not include member information, the whole group will be deleted. If member information is included, the member will be removed from the group.

.PARAMETER Separator
Defines column separator used in the data file. The default separator is TAB. Do not use comma as a separator, because it can appear in data, such as group names or descriptions.

.PARAMETER ExcludeEmptyGroups
Use this switch to tell the script to not export empty groups. Applies to exports only.

.PARAMETER GroupNameRegex
Specifies the regular expression that the group names must match. Applies to exports only.

.PARAMETER GroupDescriptionRegex
Specifies the regular expression that the group descriptions must match. Applies to exports only.

.PARAMETER AccountRegex
Specifies the regular expression that the member account name must match. Applies to exports only.

.PARAMETER NoMembers
Use this switch to export groups only (without members). Applies to exports only.

.PARAMETER WithHeader
When set, the data file will or will be assumed to include column header in the first row.

.PARAMETER ComputerName
Name of the computer on which the operation must be performed. If not specified, the current computer will be used.

.PARAMETER ModulePath
Optional path to directory holding the modules used by this script. This can be useful if the script runs on the system with no or restricted access to the Internet. By default, the module path will point to the 'Modules' folder in the script's folder.

.PARAMETER Log
When set to true, informational messages will be written to a log file. The default log file will be created in the backup folder and will be named after the name of Computer and operation with the '.log' extension, such as 'MYSERVER.Export.log' or 'MYSERVER.Import.log'.

.PARAMETER LogFile
Use this switch to specify a custom log file location. When this parameter is set to a non-null and non-empty value, the '-Log' switch can be omitted.

.PARAMETER LogAppend
Use this switch to appended log entries to the existing log file, if it exists (by default, the old log file will be overwritten).

.PARAMETER ErrLog
When set to true, error messages will be written to an error log file. The default error log file will be created in the backup folder and will be named after the name of Computer and operation with the '.err.log' extension, such as 'MYSERVER.Export.err.log' or 'MYSERVER.Import.err.log'.

.PARAMETER ErrLogFile
Use this switch to specify a custom error log file location. When this parameter is set to a non-null and non-empty value, the '-ErrLog' switch can be omitted.

.PARAMETER ErrLogAppend
Use this switch to appended error log entries to the existing error log file, if it exists (by default, the old error log file will be overwritten).

.PARAMETER ProgressInterval
The number of items that must be processed between progress updates. Set to higher number (100, 1000) for better performance. Set to 0 to not display progress bar.

.PARAMETER Test
Use this switch to verify the process and data without making changes to the server (for the import or delete operations) or the data file (for exports).

.PARAMETER Detailed
Use this switch to log more details about the performed operations. Keep in mind that it can negatively affect the script performance (by the order of magnitude).

.PARAMETER Quiet
Use this switch to suppress regular log entries sent to a console (errors and warnings will still be displayed). You can use this switch to reduce the progress bar flickering due to extensive console scrolling.

.PARAMETER NoLogo
Specify this command-line switch to not print version and copyright info.

.PARAMETER ConfigFile
Path to the optional custom config file. The default config file is named after the script with the '.json' extension, such as 'CloneLocalGroups.ps1.json'.

.NOTES
Version    : 1.0.5
Author     : Alek Davis
Created on : 2019-08-22
License    : MIT License
LicenseLink: https://github.com/alekdavis/CloneLocalGroups/blob/master/LICENSE
Copyright  : (c) 2019 Alek Davis

.LINK
https://github.com/alekdavis/CloneLocalGroups

.INPUTS
None.

.OUTPUTS
None.

.EXAMPLE
.\CloneLocalGroups.ps1
Exports local groups with members to the default data file.

.EXAMPLE
.\CloneLocalGroups.ps1 -Import -DataFile "D:\Data\MYSERVER.txt"
Imports local groups with members from the specified data file.

.EXAMPLE
.\CloneLocalGroups.ps1 -Import -DataFile "D:\Data\MYSERVER.txt" -Test
Perform a dry run for the import operation without creating groups or assigning group members.

.EXAMPLE
.\CloneLocalGroups.ps1 -Delete -DataFile "D:\Data\MYSERVER.txt"
Deletes local groups and/or members defined in the specified data file.

.EXAMPLE
Get-Help .\CloneLocalGroups.ps1
View help information.
#>

#------------------------------[ IMPORTANT ]-------------------------------

<#
PLEASE MAKE SURE THAT THE SCRIPT STARTS WITH THE COMMENT HEADER ABOVE AND
THE HEADER IS FOLLOWED BY AT LEAST ONE BLANK LINE; OTHERWISE, GET-HELP AND
GETVERSION COMMANDS WILL NOT WORK.
#>

#------------------------[ RUN-TIME REQUIREMENTS ]-------------------------

#Requires -Version 4.0
#Requires -RunAsAdministrator

#------------------------[ COMMAND-LINE SWITCHES ]-------------------------

# Script command-line arguments (see descriptions in the .PARAMETER comments
# above).
[CmdletBinding(DefaultParameterSetName="default")]
param (
    [Parameter(ParameterSetName="Export")]
    [Alias("E")]
    [switch]
    $Export,

    [Parameter(ParameterSetName="Import")]
    [Alias("I")]
    [switch]
    $Import,

    [Parameter(ParameterSetName="Delete")]
    [Alias("D")]
    [switch]
    $Delete,

    [Parameter(Position=0)]
    [string]
    $DataFile,

    [string]
    $Separator = "`t",

    [switch]
    $WithHeader,

    [Parameter(ParameterSetName="Export")]
    [switch]
    $ExcludeEmptyGroups,

    [Parameter(ParameterSetName="Export")]
    [switch]
    $NoMembers,

    [Parameter(ParameterSetName="Export")]
    [string]
    $GroupNameRegex = ".",

    [Parameter(ParameterSetName="Export")]
    [string]
    $GroupDescriptionRegex = ".*",

    [Parameter(ParameterSetName="Export")]
    [string]
    $AccountRegex = ".",

    [string]
    $ModulePath = "$PSScriptRoot\Modules",

    [switch]
    $Log,

    [switch]
    $LogAppend,

    [string]
    $LogFile,

    [switch]
    $ErrLog,

    [switch]
    $ErrLogAppend,

    [string]
    $ErrLogFile,

    [Alias("Config")]
    [string]
    $ConfigFile,

    [Alias("Computer","Machine","Server")]
    [string]
    $ComputerName = $env:COMPUTERNAME,

    [int]
    [ValidateRange(0, [int]::MaxValue)]
    [Alias("Progress")]
    $ProgressInterval = 1,

    
    [Parameter(ParameterSetName="Import")]
    [Parameter(ParameterSetName="Delete")]
    [Alias("T")]
    [switch]
    $Test,

    [switch]
    $Detailed,

    [Alias("Q")]
    [switch]
    $Quiet,

    [switch]
    $NoLogo
)

#------------------------[ CONFIGURABLE VARIABLES ]------------------------

# List of regex matches of group names to be excluded from export.
$ExcludeGroupNameRegex = @()

# List of regex matches of group descriptions to be excluded from export.
$ExcludeGroupDescriptionRegex = @()

# List of regex matches of account names to be excluded from export.
$ExcludeAccountRegex = @()

# File extensions.
$LogFileExt     = "log"
$ErrLogFileExt  = "err.log"
$DataFileExt    = "txt"

# Batch size for data file line counting operation (1000 seems to be optimal).
$MaxLineCount   = 1000

#----------------------[ NON-CONFIGURABLE VARIABLES ]----------------------

$Operation = $null

#-------------------------[ WINDOWS ERROR CODES ]--------------------------

$ERRCODE_USER_ALREADY_A_MEMBER = 0x80070562
$ERRCODE_USER_IS_NOT_A_MEMBER  = 0x80070561

#------------------------------[ EXIT CODES]-------------------------------

$EXITCODE_ERROR = 1

#------------------------------[ FUNCTIONS ]-------------------------------

#--------------------------------------------------------------------------
# LoadModule
#   Installs (if needed) and loads a PowerShell module.
function LoadModule {
    param(
        [string]
        $ModuleName
    )

    # Download module if needed and import it into the process.
    if (!(Get-Module -Name $ModuleName)) {

        if (!(Get-Module -Listavailable -Name $ModuleName)) {
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module $ModuleName -ErrorAction Stop -Force
    }
}

#--------------------------------------------------------------------------
# GetTimestamp
#   Returns current timestamp in a consistent format.
function GetTimestamp {
    param (
        [datetime]
        $date = $null,

        [bool]
        $withMsec = $true
    )

    if (!$date) {
        $date = Get-Date
    }

    $format = "yyyy/MM/dd HH:mm:ss"

    if ($withMsec) {
        $format += ".fff"
    }

    return $date.ToString($format)
}

#--------------------------------------------------------------------------
# CountLines
#   Returns the number of lines in a text file.
function CountLines {
    param (
        [string]
        $path,

        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $max = 1000
    )

    $count = 0

    # Do it by one batch at a time (1000 seems to be the optimal batch size).
    Get-Content $path -read $max | ForEach-Object { $count += $_.Length };

    return $count
}


#--------------------------------------------------------------------------
# UpdateProgress
#   Displays progress bar with message.
function UpdateProgress {
    param (
        [string]
        $activity,

        [string]
        $group,

        [string]
        $user,

        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $processed,

        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $total = 0,

        [datetime]
        $startTime
    )

    # If interval is set to 0, do not show progress.
    if ($Script:ProgressInterval -lt 1) {
        return
    }

    $params = @{}

    $update = $false

    # Always update progress with the last item
    if ($processed -eq $total) {
        $update = $true
    }
    # Always update if interval is 1.
    elseif ($Script:ProgressInterval -eq 1) {
        $update = $true
    }
    # Update every n-th (interval) item.
    else {
        # Update always for the first item.
        if ($processed -eq 1) {
            $update = $true
        }
        elseif (($processed % $Script:ProgressInterval) -eq 0) {
            $update = $true
        }
    }

    if (!$update) {
        return
    }

    # Once we reach the last record, mark progress as completed.
    if (($total -gt 0) -and ($processed -ge $total)) {
        $params.Add("Completed", $true)
    }

    $currentOperation = $null

    if ($group -and $user) {
        $currentOperation = "$group : $user"
    }
    elseif ($group) {
        $currentOperation = $group
    }
    elseif ($user) {
        $currentOperation = $user
    }

    if ($currentOperation) {
        $params.Add("CurrentOperation", $currentOperation)
    }

    $status = $null

    # If we have the total number of items, we can calculate speed and ETA.
    if ($total -gt 0) {
        # How many items are remaining?
        $remaining = $total - $processed

        if ($remaining -gt 0) {

            # Calculate percentage of processed items.
            $percent = ($processed / $total) * 100
            $displayPercent = "{0:n1}" -f $percent

            # Set the progress bar to the appropriate percentage.
            $params.Add("PercentComplete", $percent)

            # For the first item, we do not have enough info for accurate calculations.
            if ($processed -gt 1) {
                $currentTime = Get-Date

                # Time spent to procss all items so far.
                $timeSpan = New-TimeSpan -Start $startTime -End $currentTime

                # How many milliseconds takes to process one item?
                $speed      = $timeSpan.TotalMilliseconds / $processed

                # Rate is dependent on speed: items/sec, items/min, items/hr.
                $rate       = 0
                $displayRate= $null

                # Takes less then a second per item.
                if ($speed -lt 1000) {
                    $rate           = 1000 / $speed
                    $displayRate    = "{0:n0}" -f $rate
                    $displayRate   += "/sec"
                # Takes less then a minute per item.
                } elseif ($speed -lt 60000) {
                    $rate           = 60000 / $speed
                    $displayRate    = "{0:n0}" -f $rate
                    $displayRate   += "/min"
                # Takes less more than a minute per item.
                } else {
                    $rate           = 360000 / $speed
                    $displayRate    = "{0:n0}" -f $rate
                    $displayRate   += "/hr"
                }

                # How many milliseconds to process the remaining items left?
                $timeLeft           = [TimeSpan]::FromMilliseconds($remaining * $speed)
                $displayTimeleft    = $null

                # Format ETA.
                if ($timeLeft.TotalSeconds -lt 60) {
                    $sec            = "{0:n0}" -f $timeLeft.Seconds
                    $displayTimeleft = "$sec sec"
                } elseif ($timeLeft.TotalMinutes -lt 60) {
                    $min            = "{0:n0}" -f $timeLeft.Minutes
                    $displayTimeleft = "$min min"

                    if ($timeLeft.Seconds -gt 0) {
                        $sec            = "{0:n0}" -f $timeLeft.Seconds
                        $displayTimeleft += " $sec sec"
                    }
                } else {
                    $hr            = "{0:n0}" -f $timeLeft.TotalHours
                    $displayTimeleft = "$hr hr"

                    if ($timeLeft.Minutes -gt 0) {
                        $min            = "{0:n0}" -f $timeLeft.Minutes
                        $displayTimeleft += " $min min"
                    }
                }

                $status = "Processing $processed of $total ($displayPercent% complete, ETA: $displayTimeleft at $displayRate)"
            }
            else {
                $status = "Processing $processed of $total ($displayPercent% complete)"
            }
        }
        else {
            $params.Add("PercentComplete", 100)

            $status = "Done"
        }
    }
    else {
        $status = "Processing $processed"
    }

    Write-Progress @params -Activity $activity -Status $status
}

#--------------------------------------------------------------------------
# Indent
#   Adds indent(s) in front of the text string.
function Indent {
    param (
        [string]
        $message,
        [int]
        $indentLevel = 1
    )

    $indents = ""

    if ($message) {
        for ($i=0; $i -lt $indentLevel; $i++) {
            $indents = $indents + "  "
        }
    }
    else {
        $message = ""
    }

    $message = $indents + $message

    return $message
}

#--------------------------------------------------------------------------
# Log
#   Writes a message to the script's console window or a text file.
function Log {
    param (
        [string]
        $message = "",
        [string]
        $path = $null,
        [System.ConsoleColor]
        $textColor = $Host.ui.RawUI.ForegroundColor
    )

    #if ($Script:Log -and $Script:LogFile -and $writeToFile) {
    if ($path) {
        $Error.Clear()

        try {
            $message | Out-File $path -Append -Encoding utf8
        }
        catch {
            # There was a problem writing to the file.
            LogException $_

            $Error.Clear()
            return $false
        }
    }
    else {
        if ([int]$textColor -eq -1) {
            [System.ConsoleColor]$foregroundColor = [System.ConsoleColor]::White
        }
        else {
            [System.ConsoleColor]$foregroundColor = $textColor
        }

        Write-Host $message -ForegroundColor $foregroundColor
    }

    return $true
}

#--------------------------------------------------------------------------
# LogException
#   Writes exception info to one or more of the following:
#   - active console window
#   - script's log file
#   - script's error log file
function LogException {
    param (
        [object]
        $ex,
        [bool]
        $writeToLog = $true,
        [bool]
        $writeToErrLog = $true
    )

    # Always write exceptions to console.
    Write-Error $ex -ErrorAction Continue

    # Write to log file if settings align.
    if ($writeToLog -and $Script:Log -and $Script:LogFile) {
        try {
            Out-File -FilePath $Script:LogFile -Append -Encoding utf8 -InputObject $ex
        }
        catch {
            # There was a problem writing to the log file, so don't use the log
            # file from this point on.
            $Script:Log     = $false
            $Script:LogFile = $null

            try {
                # Log error that we failed to write to the log file.
                LogException $_ $false $writeToErrLog
            }
            catch {
            }

            $Error.Clear()
        }
    }

    # Write to error log file if settings align.
    if ($writeToErrLog -and $Script:ErrLog -and $Script:ErrLogFile) {
        try {
            Out-File -FilePath $Script:ErrLogFile -Append -Encoding utf8 -InputObject $ex
        }
        catch {
            # There was a problem writing to the error log file, so don't use the error log
            # file from this point on.
            $Script:ErrLog     = $false
            $Script:ErrLogFile = $null

            try {
                # Log error that we failed to write to the error log file.
                LogException $_ $writeToLog $false
            }
            catch {
            }

            $Error.Clear()
        }
    }
}

#--------------------------------------------------------------------------
# LogException
#   Writes an error message to one or more of the following:
#   - active console window
#   - script's log file
#   - script's error log file
function LogError {
    param (
        [string]
        $message = "",
        [bool]
        $writeToLog = $true,
        [bool]
        $writeToErrLog = $true
    )

    # Always send error messages to console.
    (Log $message $null Red) | Out-Null

    # If we failed to write to the log file, don't try again.
    if ($writeToLog -and $Script:Log -and $Script:LogFile) {
        if (!(Log $message $Script:LogFile)) {
            $Script:Log     = $false
            $Script:LogFile = $null
        }
    }

    # If we failed to write to the error log file, don't try again.
    if ($writeToErrLog -and $Script:ErrLog -and $Script:ErrLogFile) {
        if (!(Log $message $Script:ErrLogFile)) {
            $Script:ErrLog     = $false
            $Script:ErrLogFile = $null
        }
    }
}

#--------------------------------------------------------------------------
# LogWarning
#   Writes a warning message to one or more of the following:
#   - active console window
#   - script's log file
function LogWarning {
    param (
        [string]
        $message = "",
        [bool]
        $writeToFile = $true,
        [bool]
        $writeToConsole = $true
    )

    if (!$Script:Quiet) {
        return
    }

    # Always send warning messages to console.
    (Log $message $null Yellow) | Out-Null

    if (!$writeToFile -or !$Script:Log -or !$Script:LogFile) {
        return
    }

    # If we failed to write to the log file, don't try again.
    if (!(Log $message $Script:LogFile))
    {
        $Script:Log     = $false
        $Script:LogFile = $null
    }
}

#--------------------------------------------------------------------------
# LogDetailed
#   Writes an informational message to one or more of the following:
#   - active console window
#   - script's log file
function LogDetailed {
    param (
        [string]
        $message = "",
        [bool]
        $writeToFile = $true,
        [bool]
        $writeToConsole = $true
    )

    if (!$Script:Detailed -or $Script:Quiet) {
        return
    }

    # Send message to console.
    if ($writeToConsole) {
        (Log $message $null) | Out-Null
    }

    if (!$writeToFile -or !$Script:Log -or !$Script:LogFile) {
        return
    }

    # If we failed to write to the log file, don't try again.
    if (!(Log $message $Script:LogFile)) {
        $Script:Log     = $false
        $Script:LogFile = $null
    }
}

#--------------------------------------------------------------------------
# LogMessage
#   Writes an informational message to one or more of the following:
#   - active console window
#   - script's log file
function LogMessage {
    param (
        [string]
        $message = "",
        [bool]
        $writeToFile = $true,
        [bool]
        $writeToConsole = $true
    )

    if ($Script:Quiet) {
        return
    }

    # Send message to console.
    if ($writeToConsole) {
        (Log $message $null) | Out-Null
    }

    if (!$writeToFile -or !$Script:Log -or !$Script:LogFile) {
        return
    }

    # If we failed to write to the log file, don't try again.
    if (!(Log $message $Script:LogFile)) {
        $Script:Log     = $false
        $Script:LogFile = $null
    }
}

#--------------------------------------------------------------------------
# InitLog
#   Initializes log settings.
function InitLog {
    if ($Script:LogFile) {
        $Script:Log = $true
    }

    if ($Script:ErrLogFile) {
        $Script:ErrLog = $true
    }

    if ((!$Script:Log) -and (!$Script:ErrLog)) {
        return $true
    }

    if ($Script:Log -and !$Script:LogFile) {
        $Script:LogFile = Join-Path $PSScriptRoot `
            "$($Script:ComputerName).$($Script:Operation).$($Script:LogFileExt)"
    }

    if ($Script:ErrLog -and !$Script:ErrLogFile) {
        $Script:ErrLogFile = Join-Path $PSScriptRoot `
            "$($Script:ComputerName).$($Script:Operation).$($Script:ErrLogFileExt)"
    }

    # Process both log and error log files in a similar manner.
    $logSettings = @(
        @($Script:LogFile, $Script:LogAppend, "log file", $false, $false),
        @($Script:ErrLogFile, $Script:ErrLogAppend, "error log file", $true, $false))

    $validatedFolders = @{}

    foreach ($logSetting in $logSettings) {

        # If there is no log file path, skip to next.
        if (!$logSetting[0]) {
            continue
        }

        $filePath       = $logSetting[0]
        $fileDir        = Split-Path -Path $filePath -Parent
        $fileAppend     = $logSetting[1]
        $fileType       = $logSetting[2]
        $writeToLog     = $logSetting[3]
        $writeToErrLog  = $logSetting[4]

        if (!$fileAppend) {
            if (Test-Path -Path $filePath -PathType Leaf) {
                try {
                    Remove-Item -Path $filePath -Force
                }
                catch {
                    LogError ("Cannot delete old " + $fileType + ":") $writeToLog $writeToErrLog
                    LogError (Indent $filePath) $writeToLog $writeToErrLog

                    LogException $_ $writeToLog $writeToErrLog

                    return $false
                }
            }
        }

        # Make sure log folder exists.
        if (!($validatedFolders.ContainsKey($fileDir))) {
            LogDetailed ("Validating " + $fileType + " folder:") $writeToLog
            LogDetailed (Indent $fileDir) $writeToLog

            try {
                if (!(Test-Path -Path $fileDir -PathType Container)) {
                    New-Item -Path $fileDir -ItemType Directory -Force | Out-Null
                }

                $validatedFolders[$fileDir] = $true
            }
            catch {
                LogException $_ $writeToLog $writeToErrLog

                $Error.Clear()

                return $false;
            }
        }
    }

    return $true
}

#--------------------------------------------------------------------------
# PrintVersion
#   Prints script version info to the console and/or the log file.
function PrintVersion {
    param (
        [bool]
        $writeToFile = $true,
        [bool]
        $writeToConsole = $true
    )

    $versionInfo = Get-ScriptVersion
    $scriptName  = (Get-Item $PSCommandPath).Basename

    LogMessage ($scriptName +
        " v" + $versionInfo["Version"] +
        " " + $versionInfo["Copyright"]) $writeToFile $writeToConsole
}

#--------------------------------------------------------------------------
# Init
#   Initializes global variables.
function Init {
    param(
    )

    $modes = 0

    # Make sure that only one mode is selected (can happen via config file).
    if ($Script:Export) { $modes++ }
    if ($Script:Import) { $modes++ }
    if ($Script:Delete) { $modes++ }

    # Export is the default operation.
    if ($modes -eq 0) {
        $Script:Export = $true
    }
    elseif ($modes -gt 1) {
        throw [System.Management.Automation.ParameterBindingException] `
            ("Multiple operation modes (-Export, -Import, -Delete) specified. " +
            "Please makes sure that only one mode is set between the " +
            "command-line arguments and the script's configuration file.")
    }

    # Set the operation string for display purposes.
    if ($Script:Export) {
        $Script:Operation = "Export"
    }
    elseif ($Script:Import) {
        $Script:Operation = "Import"
    }
    else {
        $Script:Operation = "Delete"
    }

    # If not specified, generate the path to default data file.
    if (!$Script:DataFile) {
        $Script:DataFile = Join-Path $PSScriptRoot `
            "$($Script:ComputerName).$($Script:DataFileExt)"
    }
}

#--------------------------------------------------------------------------
# IgnoreGroup
#   Checks if the specified group must be ignored.
function IgnoreGroup {
    param(
        [string]
        $name,

        [string]
        $description
    )

    # Match by name was already done by the query that returned the list of groups.

    # Does the description match the corresponding regex?
    if ($description -notmatch $Script:GroupDescriptionRegex) {
        LogWarning "Group description:"
        LogWarning (Indent $group.Name)
        LogWarning "does not match the GroupDescriptionRegex value:"
        LogWarning (Indent $Script:GroupDescriptionRegex)

        return $true
    }

    # Does the name match one in the corresponding regex list of exclusions?
    foreach ($excludeGroup in $Script:ExcludeGroupNameRegex) {
        if ($name -match $excludeGroup) {
            LogWarning "Group name:"
            LogWarning (Indent $name)
            LogWarning "matches one of the ExcludeGroupNameRegex values:"
            LogWarning (Indent $excludeGroup)

            return $true
        }
    }

    # Does the description match one in the corresponding regex list of exclusions?
    foreach ($excludeGroup in $Script:ExcludeGroupDescriptionRegex) {
        if ($description -match $excludeGroup) {
            LogWarning "Group description:"
            LogWarning (Indent $description)
            LogWarning "matches one of the ExcludeGroupDescriptionRegex values:"
            LogWarning (Indent $excludeGroup)

            return $true
        }
    }

    return $false
}

#--------------------------------------------------------------------------
# IgnoreMember
#   Checks if the specified member must be ignored.
function IgnoreMember {
    param(
        [string]
        $account
    )

    # Does the account name match the corresponding regex?
    if ($account -notmatch $Script:AccountRegex) {
        LogWarning "User account:"
        LogWarning (Indent $account)
        LogWarning "does not match the AccountRegex value:"
        LogWarning (Indent $Script:AccountRegex)

        return $true
    }

    # Does the name match one in the corresponding regex list of exclusions?
    foreach ($excludeAccount in $Script:ExcludeAccountRegex) {
        if ($account -match $excludeAccount) {
            LogWarning "User account:"
            LogWarning (Indent $account)
            LogWarning "matches one of the ExcludeAccountRegex values:"
            LogWarning (Indent $excludeAccount)

            return $true
        }
    }

    return $false
}

#--------------------------------------------------------------------------
# CreateGroup
#   Creates a local group.
function CreateGroup {
    param(
        [string]
        $server,

        [string]
        $name,

        [string]
        $description
    )

    LogDetailed "Creating local group:"
    LogDetailed (Indent $name)

    try {
        $group = ([ADSI]$server).Create("Group", $name)
        $group.SetInfo()
        $group.Description = $description
        $group.SetInfo()
    }
    catch {
        LogError "Cannot create group:"
        LogError (Indent $name)

        LogException $_

        $Error.Clear()

        return $false
    }

    return $true
}

#--------------------------------------------------------------------------
# Import
#   Imports local groups or group members from a text file.
function Import {
    param(
        [datetime]
        $startTime
    )

    $count   = 0
    $total   = 0
    $current = 0

    $skipMembers = $false

    $activity   = "Importing local groups"
    $server     = "WinNT://$($Script:ComputerName)"

    $lastGroupName  = ""
    $msg            = $null

    # Calculate the number of lines in file for the progress info.
    if ($Script:ProgressInterval -gt 0) {
        LogDetailed "Counting lines in the data file..."

        $total = CountLines $Script:DataFile $Script:MaxLineCount

        LogDetailed "Lines in the data file:"
        LogDetailed (Indent $total.ToString())
    }

    $reader = [System.IO.File]::OpenText($Script:DataFile)

    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $current++

            # Skip the header line.
            if (($current -eq 1) -and ($Script:WithHeader)) {
                continue
            }

            $groupName  = $null
            $account    = $null
            $description= $null

            try {
                $groupName, $account, $description = $line -split $Script:Separator

                if ($groupName)   { $groupName   = $groupName.Trim() }
                if ($account)     { $account     = $account.Trim() }
                if ($description) { $description = $description.Trim() }
            }
            catch {
                LogError "Invalid line format:"
                LogError (Indent $line)

                LogException $_

                $Error.Clear()
            }

            # If there is no group name, skip this line.
            if (!$groupName) {
                continue
            }

            UpdateProgress $activity $groupName $account $current $total $startTime

            $groupCN = "$server/$groupName,group"

            # If this is a new group in the list, we may need to create it.
            if ($groupName -ne $lastGroupName) {

                LogDetailed "Group:"
                LogDetailed (Indent $groupName)

                $skipMembers = $false

                if (-not ([ADSI]::Exists($groupCN))) {
                    if (!$Script:Test) {
                        if (-not (CreateGroup $server $groupName $description)) {
                            # Could not create the groups: no point in assigning members.
                            $skipMembers = $true
                        }
                    }
                }

                $lastGroupName = $groupName
                $msg = "Members:"
            }

            if ($skipMembers) {
                continue
            }

            # If account field is missing, skip.
            if (!$account) {
                $count++
                continue
            }

            # Log header for members.
            if ($msg) {
                LogDetailed $msg
                $msg = $null
            }

            LogDetailed (Indent $account)
            try {
                if (!$Script:Test) {
                   ([ADSI]$groupCN).Add($account) | Out-Null
                }
                $count++
            }
            catch {
                if ($_.Exception.InnerException.ErrorCode -eq $ERRCODE_USER_ALREADY_A_MEMBER) {
                    $count++
                }
                else {
                    LogError "Cannot add account:"
                    LogError (Indent $account)
                    LogError "to group:"
                    LogError (Indent $groupName)

                    LogException $_
                }

                $Error.Clear()
            }
        }
    }
    finally {
        $reader.Close()

        UpdateProgress $activity $null $null $total $total $startTime
    }

    return $count
}

#--------------------------------------------------------------------------
# Export
#   Exports local groups or group members to a text file.
function Export {
    param(
        [datetime]
        $startTime
    )

    $count = 0
    $activity = "Exporting local groups"

    # Delete old data file if it already exists.
    if (Test-Path -Path $Script:DataFile -PathType Leaf) {
        if (!$Script:Test) {
            try {
                Remove-Item -Path $Script:DataFile -Force
            }
            catch {
                LogError "Cannot delete old data file:"
                LogError (Indent $Script:DataFile)

                LogException $_

                $Error.Clear()

                return $count
            }
        }
    }

    [ADSI]$server = "WinNT://$($Script:ComputerName)"

    # Get list of groups with matching names (default match supports any name).
    $groups = ($server.Children |
        Where-Object({($_.class -eq 'group') -and
            ($_.name.value -match $Script:GroupNameRegex)}))

    # If no matching groups found, we're done.
    if (!$groups) {
        return $count;
    }

    # This is the most efficient way of appending an output text file.
    if (!$Script:Test) {
        $writer = New-Object System.IO.StreamWriter $Script:DataFile
    }

    try {
        # Write header line.
        if (!$Script:Test) {
            if ($Script:WithHeader) {
                $data = "GROUP$($Script:Separator)ACCOUNT$($Script:Separator)DESCRIPTION"
                $writer.WriteLine($data)
            }
        }

        foreach ($group in $groups) {
            LogDetailed "Group:"
            LogDetailed (Indent $group.Name)

            # We may need to save the name (only in the first record for this group).
            $description = $group.Properties["Description"]

            if ($null -eq $description) {
                $description = ""
            }

            # Check group against other conditions (description and exclusion lists).
            if (IgnoreGroup $group.Name $description) {
                continue
            }

            # By default, include empty groups.
            if (!$Script:ExcludeEmptyGroups) {
                $data = "$($group.Name)$($Script:Separator)$($Script:Separator)$description"

                if (!$Script:Test) {
                    $writer.WriteLine($data)
                }

                $description = ""

                $count++

                UpdateProgress $activity $group.Name $null $count 0 $startTime
            }

            [ADSI]$adsiGroup = "$($group.Parent)/$($group.Name),group"

            # Get the list of group members.
            $members = $adsiGroup.psbase.Invoke("Members")

            # If group is empty, skip the rest.
            if (!$members) {
                continue
            }

            $msg = "Members:"

            # Process each group member.
            foreach ($member in $members) {
                $account = $member.GetType().InvokeMember("ADSPath", "GetProperty", $null, $member, $null)

                # Check if the member must be excluded from export.
                if (IgnoreMember $account) {
                    continue
                }

                # When exporting groups only, skip members.
                if ($Script:NoMembers) {
                    if ($Script:ExcludeEmptyGroups) {
                        # We haven't exported the group info, yet.
                        $data = "$($group.Name)$($Script:Separator)$($Script:Separator)$description"

                        if (!$Script:Test) {
                            $writer.WriteLine($data)
                        }

                        $count++

                        UpdateProgress $activity $group.Name $null $count 0 $startTime
                    }

                    break;
                }

                # Display Members heading.
                if ($msg) {
                    LogDetailed $msg
                    $msg = $null
                }

                LogDetailed (Indent $account)

                $data = "$($group.Name)$($Script:Separator)$account$($Script:Separator)$description"

                if (!$Script:Test) {
                    $writer.WriteLine($data)
                }

                # No need for description after the first group member was exported.
                $description = ""

                $count++

                UpdateProgress $activity $group.Name $account $count 0 $startTime
            }
        }
    }
    finally {
        if (!$Script:Test) {
            $writer.Close()
        }

        UpdateProgress $activity $null $null $count $count $startTime
    }

    return $count
}

#--------------------------------------------------------------------------
# Delete
#   Deletes local groups or group members via a text file.
function Delete {
    param(
        [datetime]
        $startTime
    )

    $count   = 0
    $total   = 0
    $current = 0

    $skipMembers = $false

    $activity   = "Deleting local groups"
    $server     = "WinNT://$($Script:ComputerName)"

    $lastGroupName  = ""
    $msg            = $null

    # Calculate the number of lines in file for the progress info.
    if ($Script:ProgressInterval -gt 0) {
        LogDetailed "Counting lines in the data file..."

        $total = CountLines $Script:DataFile $Script:MaxLineCount

        LogDetailed "Lines in the data file:"
        LogDetailed (Indent $total.ToString())
    }

    $reader = [System.IO.File]::OpenText($Script:DataFile)

    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $current++

            # Skip the header line.
            if (($current -eq 1) -and ($Script:WithHeader)) {
                continue
            }

            $groupName  = $null
            $account    = $null
            $description= $null

            try {
                $groupName, $account, $description = $line -split $Script:Separator

                if ($groupName) { $groupName = $groupName.Trim() }
                if ($account)   { $account   = $account.Trim() }
            }
            catch {
                LogError "Invalid line format:"
                LogError (Indent $line)

                LogException $_

                $Error.Clear()
            }

            # If there is no group name, skip this line.
            if (!$groupName) {
                continue
            }

            UpdateProgress $activity $groupName $account $current $total $startTime

            $groupCN = "$server/$groupName,group"

            if ($groupName -ne $lastGroupName) {

                LogDetailed "Group:"
                LogDetailed (Indent $groupName)

                $skipMembers = $false
            }

            # If we do not have a user account, delete the group.
            if (!$account) {
                if ([ADSI]::Exists($groupCN)) {
                    try {
                        if (!$Script:Test) {
                            ([ADSI]$server).Delete('group', $groupName) | Out-Null
                        }

                        # No need to delete members if the group was deleted.
                        $skipMembers = $false
                        $count++
                    }
                    catch {
                        LogError "Cannot delete group:"
                        LogError (Indent $groupName)

                        LogException $_
                    }
                }
                else {
                    # No need to delete members if the group does not exist.
                    $skipMembers = $false
                    $count++
                }

                $lastGroupName = $groupName
                $msg = "Members:"
            }
            # Delete account from group.
            else {
                if ($lastGroupName -ne $groupName) {
                    LogDetailed "Group:"
                    LogDetailed (Indent $groupName)

                    $skipMembers = $false
                    $msg = "Members:"
                }

                if ($skipMembers) {
                    $msg = $null
                    continue
                }

                # Log header for members.
                if ($msg) {
                    LogDetailed $msg
                    $msg = $null
                }

                LogDetailed (Indent $account)

                if (![ADSI]::Exists($groupCN)){
                    $skipMembers = $true
                }
                else {
                    try {
                        if (!$Script:Test) {
                            ([ADSI]$groupCN).Remove($account) | Out-Null
                        }
                        $count++
                    }
                    catch {
                        if ($_.Exception.InnerException.ErrorCode -eq $ERRCODE_USER_IS_NOT_A_MEMBER) {
                            $count++
                        }
                        else {
                            LogError "Cannot remove account:"
                            LogError (Indent $account)
                            LogError "from group:"
                            LogError (Indent $groupName)

                            LogException $_
                        }

                        $Error.Clear()
                    }
                }

                $lastGroupName = $groupName
           }
        }
    }
    finally {
        $reader.Close()

        UpdateProgress $activity $null $null $total $total $startTime
    }

    return $count
}

#---------------------------------[ MAIN ]---------------------------------

# We will trap errors in the try-catch blocks.
$ErrorActionPreference = 'Stop'

# Make sure we have no pending errors.
$Error.Clear()

# Save time for logging purposes.
$startTime = Get-Date
$displayStartTime = GetTimestamp $startTime $false

# Add custom folder(s) to the module path.
if ($ModulePath) {
    if ($env:PSModulePath -notmatch ";$") {
        $env:PSModulePath += ";"
    }

    $paths = $ModulePath -split ";"

    foreach ($path in $paths){
        $path = $path.Trim();

        if (-not ($env:PSModulePath.ToLower().Contains(";$path;".ToLower()))) {
            $env:PSModulePath += "$path;"
        }
    }
}

# Load modules for reading config file settings and script version info.
# https://www.powershellgallery.com/packages/ScriptVersion
# https://www.powershellgallery.com/packages/ConfigFile
$modules = @("ScriptVersion", "ConfigFile")
foreach ($module in $modules) {
    try {
        LoadModule -ModuleName $module
    }
    catch {
        $_
        Write-Error "Cannot load module $module."
        
        exit $EXITCODE_ERROR
    }
}

# Since we have not initialized log file, yet, print to console only.
if ((!$Quiet) -and (!$NoLogo)) {
    PrintVersion $false
}

LogMessage "Script started at:" $false
LogMessage (Indent $displayStartTime) $false

# Load config settings from a config file (if any).
try {
    Import-ConfigFile -ConfigFilePath $ConfigFile -DefaultParameters $PSBoundParameters
}
catch {
    $_
    Write-Error "Cannot initialize run-time configuration settings."
    exit $EXITCODE_ERROR
}

# Initialize global variables.
Init

# Initialize log file.
if (!(InitLog)) {
    exit $EXITCODE_ERROR
}

# Print messages that we already sent to console to the log file.
if ((!$Quiet) -and (!$NoLogo)) {
    PrintVersion $true $false
}

LogMessage "Script started at:" $true $false
LogMessage (Indent $displayStartTime) $true $false

# Print command-line arguments to the log file (not needed for console).
if ($args.Count -gt 0) {
    $commandLine = ""
    for ($i = 0; $i -lt $args.Count; $i++) {
        if ($args[$i].Contains(" ")) {
            $commandLine = $commandLine + '"' + $args[$i] + '" '
        }
        else {
            $commandLine = $commandLine + $args[$i] + ' '
        }
    }

    LogMessage "Command-line arguments:" $true $false
    LogMessage (Indent $commandLine.Trim()) $true $false
}

# The rest of the log messages will be sent to console and log file
# (subject to other checks, e.g. if log file is turned off or quiet mode).

LogMessage "Operation:"
LogMessage (Indent $Operation.ToUpper())

LogMessage ("Data file:")
LogMessage (Indent $DataFile)

if ($LogFile) {
    LogMessage ("Log file:")
    LogMessage (Indent $LogFile)
}

if ($ErrLogFile) {
    LogMessage ("Error log file:")
    LogMessage (Indent $ErrLogFile)
}

$count = 0

try {
    $operationStartTime = Get-Date

    if ($Export) {
        LogMessage "Exporting..."
        $count = Export $operationStartTime
    }
    elseif ($Import) {
        LogMessage "Importing..."
        $count = Import $operationStartTime
    }
    else {
        LogMessage "Deleting..."
        $count = Delete $operationStartTime
    }
}
catch {
    LogException $_
}

LogMessage "Processed items:"
LogMessage (Indent $count.ToString())

$endTime = Get-Date

LogMessage "Script ended at:"
LogMessage (Indent (GetTimestamp $endTime $false))

LogMessage "Script ran for (hr:min:sec.msec):"
LogMessage (Indent (New-TimeSpan -Start $startTime -End $endTime).ToString("hh\:mm\:ss\.fff"))
LogMessage "Done."

# THE END
#--------------------------------------------------------------------------
