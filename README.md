# CloneLocalGroups.ps1
This [PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/overview) script can be used to export, import, or delete local groups and group members on a local or remote computer.

## Overview
Use the `CloneLocalGroups.ps1` script:

- export local groups and/or members from a local or remote computer to a text file,
- import local groups and/or members to a local or remote computer from a text file, or
- delete local groups and/or members specified in a text file on a local or remote computer.

During export, you can filter groups by name and/or description using regular expressions. You can also specify the exclusion lists (also by names and descriptions) using the configuration file. In a similar manner you can specify which group members to include or exclude during the export operation.

### Script execution
You must launch PlexBackup _as administrator_ while being logged in under the same account your Plex Media Server runs.

If you haven't done this already, you may need to adjust the PowerShell script execution policy to allow scripts to run. To check the current execution policy, run the following command from the PowerShell prompt:

```PowerShell
Get-ExecutionPolicy
```
If the execution policy does not allow running scripts, do the following:

- Start Windows PowerShell with the "Run as Administrator" option. 
- Run the following command: 

```PowerShell
Set-ExecutionPolicy RemoteSigned
```

This will allow running unsigned scripts that you write on your local computer and signed scripts downloaded from the Internet (okay, this is not a signed script, but if you copy it locally, make a non-destructive change--e.g. add a space character, remove the space character, and save the file--it should work). For additional information, see [Running Scripts](https://docs.microsoft.com/en-us/previous-versions//bb613481(v=vs.85)) at Microsoft TechNet Library.

### Dependencies

The script relies on the following modules:

- [ScriptVersion](https://www.powershellgallery.com/packages/ScriptVersion)
- [ConfigFile](https://www.powershellgallery.com/packages/ConfigFile)

To verify that the modules get installed, run the script manually. You may be [prompted](https://docs.microsoft.com/en-us/powershell/gallery/how-to/getting-support/bootstrapping-nuget) to update the [NuGet](https://www.nuget.org/downloads) version (or you can do it yourself in advance).

If the system on which the script runs has limited or does not have access to the Internet, download these packages manually and use the `-ModulePath` to specify location of the root module folder (see description of the `-ModulePath` parameter). You can find a number of articles explaining how to install a PowerShell module without internet access online.

### Runtime parameters
The default value of the the script's runtime parameters are defined in code, but you can override some of them via command-line arguments or config file settings.

### Config file
Config file is optional. The default config file must be named after the script file with the `.json` extension, such as `CloneLocalGroups.ps1.json`. If the file with this name does not exist in the backup script's folder, PlexBackup will not care. You can also specify a custom config file name (or more accurately, path) via the `ConfigFile` command-line argument ([see sample](PlexBackup.ps1.json)).

A config file must use [JSON formatting](https://www.json.org/), such as:

```JavaScript
{
    "_meta": {
        "version": "1.0",
        "strict": false,
        "description": "Sample configuration settings for PlexBackup.ps1."
    },
    "Export": { "value": null },
    "Import": { "value": null },
    "Delete": { "value": null },
    "DataFile": { "value": null },
    "Separator": { "value": null },
    "ExcludeEmptyGroups": { "value": null },
    "GroupNameRegex": { "value": null },
    "GroupDescriptionRegex": { "value": null },
    "AccountRegex": { "value": "WinNT://[a-zA-Z]+/[a-zA-Z0-9]+" },
    "MaxLineCount": { "value": null },
    "ModulePath": { "value": null },
    "Log": { "value": true },
    "LogAppend": { "value": null },
    "LogFile": { "value": null },
    "ErrLog": { "value": true },
    "ErrLogAppend": { "value": null },
    "ErrLogFile": { "value": null },
    "Computer": { "value": null },
    "HideProgressBar": { "value": null },
    "Test": { "value": null },
    "Detailed": { "value": null },
    "Quiet": { "value": null },
    "NoLogo": { "value": null }
}
```
Make sure you use proper JSON formatting (escape characters, etc) when defining the config values (e.g. you must escape each backslash characters with another backslash).

### Logging
Use the `Log` switch to write operation progress and informational messages to a log file. By default, the log file will be created in the script folder. The default log file name includes the name of computer and the performed operation, such as `MYSERVER.Import.log` or `MYSERVER.Export.log`. You can specify a custom log file path via the `LogFile` argument. By default, the script deletes an old log file (if one already exists), but if you specify the `LogAppend` switch, it will append new log messages to the existing log file.

### Error log
Use the `ErrLog` switch to write error messages to a dedicated error log file. By default, the error log file will be created in the script folder. The default error log file name is similar to the default log file name, except that it has the `.err.log` extension, such as `MYSERVER.Import.err.log` or `MYSERVER.Export.err.log`. You can specify a custom error log file path via the `ErrLogFile` argument. By default, the script deletes an old error log file (if one already exists), but if you specify the `ErrLogAppend` switch, it will append new error messages to the existing error log file. If no errors occur, the error log file will not be created.

## Syntax

```PowerShell
.\CloneLocalGroups.ps1 ` 
    [-Export] `
    [[-DataFile] <String>] `
    [-Separator <String>] `
    [-ModulePath <String>] `
    [-Log] `
    [-LogAppend] `
    [-LogFile <String>] `
    [-ErrLog] `
    [-ErrLogAppend] `
    [-ErrLogFile <String>] `
    [-ConfigFile <String>] `
    [-ComputerName <String>] `
    [-ExcludeEmptyGroups] `
    [-GroupNameRegex <String>] `
    [-GroupDescriptionRegex <String>] `
    [-AccountRegex <String>]
    [-HideProgressBar] `
    [-Detailed] `
    [-Quiet] `
    [-NoLogo] `
    [<CommonParameters>]

.\CloneLocalGroups.ps1 ` 
    -Import `
    [[-DataFile] <String>] `
    [-Separator <String>] `
    [-ModulePath <String>] `
    [-Log] `
    [-LogAppend] `
    [-LogFile <String>] `
    [-ErrLog] `
    [-ErrLogAppend] `
    [-ErrLogFile <String>] `
    [-ConfigFile <String>] `
    [-ComputerName <String>] `
    [-HideProgressBar] `
    [-Test] `
    [-Detailed] `
    [-Quiet] `
    [-NoLogo] `
    [<CommonParameters>]
```
### Arguments

`-Export`

Export local groups and/or members to a data file. This is the default operation mode (when neither `-Export`, nor `-Import`, nor `-Delete` is specified).

`-Import`

Import local groups and/or members from a data file.

`-Delete`

Delete local groups and/or members specified in a data file.

`-DataFile`

Defines the full path to the date file holding imported or exported data. The file must contain three text entries separated by tabs (or explicitly defined separator):

- Name of the local group
- Name of the account in the format: WinNT://DOMAIN/user
- Group description

The group description is only needed the first time the group appears in the file. For groups with no members, the name of the account may be left blank. Here is an example of the data file (with the tab characters identified by '{TAB}'):

```
Group A{TAB}{TAB}Empty group with no members.
Group B{TAB}WinNT://CONTOSO/jdoe{TAB}Group with two members.
Group B{TAB}WinNT://CONTOSO/jschmoe{TAB}
Group C{TAB}WinNT://CONTOSO/mjane{TAB}Group with one member.
```
When exporting data, if the 'ExcludeEmptyGroups' switch is not set, all data file entries will contain group members; otherwise, there will be a separate entry for each exported group with no member information.

The group description is needed only the first time the group appears in the data file.

When deleting data, if the data file entry does not include member information, the whole group will be deleted. If member information is included, the member will be removed from the group.

`-Separator`

Defines column separator used in the data file. The default separator is TAB. Do not use comma as a separator, because it can appear in data, such as group names or descriptions.

`-ExcludeEmptyGroups`

Use this switch to tell the script to not export empty groups.

`-GroupNameRegex`

Specifies the regular expression that the group names must match. Applies to exports only.

`-GroupDescriptionRegex`

Specifies the regular expression that the group descriptions must match. Applies to exports only.

`-AccountRegex`

Specifies the regular expression that the member account name must match. Applies to exports only.

`-NoMembers`

Use this switch to export groups only (without members). Applies to exports only.

`-WithHeader`

When set, the data file will or will be assumed to include column header in the first row.

`-ComputerName`

Name of the computer on which the operation must be performed. If not specified, the current computer will be used.

`-ModulePath`

Optional path to directory holding the modules used by this script. This can be useful if the script runs on the system with no or restricted access to the Internet. By default, the module path will point to the 'Modules' folder in the script's folder.

`-Log`

When set to true, informational messages will be written to a log file. The default log file will be created in the backup folder and will be named after the name of Computer and operation with the '.log' extension, such as 'MYSERVER.Export.log' or 'MYSERVER.Import.log'.

.PARAMETER LogFile
Use this switch to specify a custom log file location. When this parameter is set to a non-null and non-empty value, the '-Log' switch can be omitted.

`-LogAppend`

Use this switch to appended log entries to the existing log file, if it exists (by default, the old log file will be overwritten).

`-ErrLog`

When set to true, error messages will be written to an error log file. The default error log file will be created in the backup folder and will be named after the name of Computer and operation with the '.err.log' extension, such as 'MYSERVER.Export.err.log' or 'MYSERVER.Import.err.log'.

`-ErrLogFile`

Use this switch to specify a custom error log file location. When this parameter is set to a non-null and non-empty value, the '-ErrLog' switch can be omitted.

`-ErrLogAppend`

Use this switch to appended error log entries to the existing error log file, if it exists (by default, the old error log file will be overwritten).

`-ProgressInterval`

The number of items that must be processed between progress updates. Set to higher number (100, 1000) for better performance. Set to 0 to not display progress bar.

`-Test`

Use this switch to verify the process and data without making changes to the server. Applies to imports only.

`-Detailed`

Use this switch to log more details about the performed operations. Keep in mind that it can negatively affect the script performance (by the order of magnitude).

`-Quiet`

Use this switch to suppress regular log entries sent to a console (errors and warnings will still be displayed). You can use this switch to reduce the progress bar flickering due to extensive console scrolling.

`-NoLogo`

Specify this command-line switch to not print version and copyright info.

`-ConfigFile`

Path to the optional custom config file. The default config file is named after the script with the '.json' extension, such as 'CloneLocalGroups.ps1.json'.

### Examples

The following examples assume that the the default settings are used for the unspecified script arguments.

#### Example 1
```
.\CloneLocalGroups.ps1
```
Export local groups with members to the default data file.

#### Example 2
```
.\CloneLocalGroups.ps1 -Import -DataFile "D:\Data\MYSERVER.txt"
```
Import local groups with members from the specified data file.

#### Example 3
```
.\CloneLocalGroups.ps1 -Import -DataFile "D:\Data\MYSERVER.txt" -Test
```
Perform a dry run for the import operation without creating groups or assigning group members.

#### Example 4
```
.\CloneLocalGroups.ps1 -Delete -DataFile "D:\Data\MYSERVER.txt"
```
Delete local groups and/or members defined in the specified data file.

#### Example 5
```
Get-Help .\CloneLocalGroups.ps1
```
View help information.

