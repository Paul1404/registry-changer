<#
.SYNOPSIS
This script allows to manipulate registry settings of all users profiles on the computer.

.DESCRIPTION
The script provides options to import settings from an XML file or enter them manually. 
It also provides options to create a backup before applying the changes. 

.EXAMPLE
# To run the script and select an XML settings file:
.\Registry-Changer.ps1

# The script will prompt you to select an XML settings file and ask if you want to backup the registry.
# If you choose to proceed, it will update the registry settings for all user profiles according to the selected file.
# It will also create a unique log file for each run under a 'Log' subdirectory in the script's directory, storing the output messages and errors.

.PARAMETER N/A
This script doesn't accept any parameters. All inputs are interactive.

.NOTES
General notes:
- Please run the script with administrative privileges to ensure it can access and modify the registry.
- Always confirm that you have a recent backup of your registry before running the script, in case any issues arise.
- This script was last updated on 2023-06-25.

.LINK
https://github.com/Paul1404/registry-changer
#>


# Global variables
$logFilePath = ""
$MaxLogCount = 100 # Default maximum number of log files


<#
.SYNOPSIS
    This function checks if the currently installed PowerShell version meets or exceeds a specified minimum version.

.DESCRIPTION
    The Get-PowerShellVersion function takes a minimum version as a parameter and checks if the current PowerShell version 
    is equal to or greater than this minimum version. It uses the built-in variable $PSVersionTable.PSVersion to get the current 
    PowerShell version.

    If the current PowerShell version is less than the minimum version, the function throws an error message and stops execution.
    If the current PowerShell version is equal to or greater than the minimum version, the function outputs a success message.

.PARAMETER MinVersion
    This parameter takes a version number (in the format "Major.Minor") and sets it as the minimum acceptable PowerShell version.
    The default value is "5.1". 

.EXAMPLE
    Get-PowerShellVersion
    Checks if the current PowerShell version is equal to or greater than the default minimum version "5.1".

.EXAMPLE
    Get-PowerShellVersion -MinVersion "7.0"
    Checks if the current PowerShell version is equal to or greater than version "7.0".

.INPUTS
    None

.OUTPUTS
    String
    Outputs a string with either a success message detailing the current PowerShell version if it's acceptable, 
    or it throws an exception with an error message if the current version is lower than the specified minimum version.

.NOTES
    Be sure to run this function in a try-catch block to properly handle the error message thrown if the current PowerShell 
    version does not meet the specified minimum version.

.LINK
    https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions
#>
function Get-PowerShellVersion {
    param (
        [Version]$MinVersion = "5.1"
    )

    if ($PSVersionTable.PSVersion -lt $MinVersion) {
        throw "This script requires PowerShell version $MinVersion or higher. Your version is $($PSVersionTable.PSVersion)."
    } else {
        Write-CustomSuccess "PowerShell version check passed. Your version is $($PSVersionTable.PSVersion)."
    }
}


<#
.SYNOPSIS
    This function checks if the current Windows version meets or exceeds a specified minimum version.

.DESCRIPTION
    The Get-WindowsVersion function uses the Get-CimInstance cmdlet to retrieve information about the operating system 
    from the Win32_OperatingSystem class. It then compares the current Windows version with the specified minimum 
    version (MinVersion parameter).

    If the current Windows version is less than the specified minimum version, the function throws an exception with 
    an error message. If the current Windows version is equal to or greater than the minimum version, it outputs a 
    success message.

.PARAMETER MinVersion
    This parameter takes a version number (in the format "Major.Minor") and sets it as the minimum acceptable Windows version. 
    The default value is "10.0".

.EXAMPLE
    Get-WindowsVersion
    Checks if the current Windows version is equal to or greater than the default minimum version "10.0".

.EXAMPLE
    Get-WindowsVersion -MinVersion "11.0"
    Checks if the current Windows version is equal to or greater than version "11.0".

.INPUTS
    None

.OUTPUTS
    String
    Outputs a string with either a success message detailing the current Windows version if it's acceptable, 
    or it throws an exception with an error message if the current version is lower than the specified minimum version.

.NOTES
    This function requires the presence of the CIM (Common Information Model) infrastructure on the computer where it's run. 
    This infrastructure is typically present on Windows systems starting with Windows Server 2012 and Windows 8.

.LINK
    https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem
#>
function Get-WindowsVersion {
    param (
        [Version]$MinVersion = "10.0"
    )

    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $version = [Version]$osInfo.Version

    if ($version -lt $MinVersion) {
        throw "This script requires Windows version $MinVersion or higher. Your version is $($osInfo.Caption) $($osInfo.Version)."
    } else {
        Write-CustomSuccess "Windows version check passed. Your version is $($osInfo.Caption) $($osInfo.Version)."
    }
}



<#
.SYNOPSIS
    This function checks if the current user has administrative privileges.

.DESCRIPTION
    The Get-AdminPrivileges function uses the .NET classes Security.Principal.WindowsPrincipal and 
    Security.Principal.WindowsIdentity to determine if the current user has administrative privileges. 
    If the current user does not have administrative privileges, the function throws an exception with 
    an error message. If the user has administrative privileges, it outputs a success message.

.PARAMETER None
    This function does not take any parameters.

.EXAMPLE
    Get-AdminPrivileges
    Checks if the current user has administrative privileges.

.INPUTS
    None

.OUTPUTS
    String
    Outputs a string with either a success message indicating the user has administrative privileges, 
    or it throws an exception with an error message if the user does not have administrative privileges.

.NOTES
    This function is essential for scripts that need to perform tasks requiring administrative privileges, 
    such as modifying registry keys, starting or stopping services, or managing user accounts.

.LINK
    https://docs.microsoft.com/en-us/dotnet/api/system.security.principal.windowsprincipal
    https://docs.microsoft.com/en-us/dotnet/api/system.security.principal.windowsidentity
#>
function Get-AdminPrivileges {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script must be run as an Administrator. Please re-run the script as an Administrator."
    } else {
        Write-CustomSuccess "Admin privileges check passed."
    }
}


<#
.SYNOPSIS
    This function checks if there is an active internet connection.

.DESCRIPTION
    The Get-NetworkConnectivity function uses the Test-Connection cmdlet to ping cloudflares one.one.one.one. If the function cannot 
    ping 1.1.1.1 successfully, it throws an exception with an error message. If the ping is successful, 
    indicating an active internet connection, it outputs a success message.

.PARAMETER None
    This function does not take any parameters.

.EXAMPLE
    Get-NetworkConnectivity
    Checks if there is an active internet connection.

.INPUTS
    None

.OUTPUTS
    String
    Outputs a string with either a success message indicating an active internet connection, or it throws an exception with 
    an error message if there is no active internet connection.

.NOTES
    This function is useful for scripts that need to connect to remote servers or perform tasks requiring an internet connection.
    The script assumes that 1.1.1.1 is always available, which might not be true in all cases (for example, if 
    1.1.1.1 is blocked on your network).

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-connection
#>
function Get-NetworkConnectivity {
    if (!(Test-Connection -ComputerName one.one.one.one -Count 2 -Quiet)) {
        throw "This script requires an active internet connection. Please check your connection and try again."
    } else {
        Write-CustomSuccess "Network connectivity and Name Resolution check passed."
    }
}


<#
.SYNOPSIS
    This function checks if a specific drive has enough free disk space.

.DESCRIPTION
    The Get-DiskSpace function uses the Get-PSDrive cmdlet to obtain information about the specified drive. 
    It then checks if the free space on the drive is greater than or equal to the minimum required free space 
    (MinSpaceGB parameter).

    If the drive does not exist, or if the free space on the drive is less than the minimum required free space, 
    the function throws an exception with an error message. If the drive exists and has sufficient free space, 
    it outputs a success message.

.PARAMETER DriveLetter
    This parameter specifies the drive to check. The default value is "C:".

.PARAMETER MinSpaceGB
    This parameter specifies the minimum required free space on the drive, in gigabytes. The default value is 5.

.EXAMPLE
    Get-DiskSpace
    Checks if the C: drive has at least 5 GB of free space.

.EXAMPLE
    Get-DiskSpace -DriveLetter "D:" -MinSpaceGB 10
    Checks if the D: drive has at least 10 GB of free space.

.INPUTS
    None

.OUTPUTS
    String
    Outputs a string with either a success message indicating the drive has enough free space, or it throws an exception 
    with an error message if the drive does not exist or does not have enough free space.

.NOTES
    This function is useful for scripts that need to ensure sufficient disk space before performing tasks that might require 
    a significant amount of space (like downloading files, creating backups, etc.).

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-psdrive
#>
function Get-DiskSpace {
    param (
        [string]$DriveLetter = "C:",
        [int]$MinSpaceGB = 5
    )

    $drive = Get-PSDrive -Name $DriveLetter.Replace(':', '')

    if ($null -eq $drive) {
        throw "Drive $DriveLetter does not exist."
    } elseif ($drive.Free / 1GB -lt $MinSpaceGB) {
        throw "Drive $DriveLetter does not have enough free space. Minimum required: ${MinSpaceGB}GB."
    } else {
        Write-CustomSuccess "Drive $DriveLetter has enough free space."
    }
}



<#
.SYNOPSIS
    Creates a new log file for each script run.

.DESCRIPTION
    The New-LogFile function creates a new log file every time the script is run. Log files are stored in a directory named 'Log' in the same location as the script. 
    If the 'Log' directory does not exist, it is created.
    The log file is named 'Log_YYYYMMDD_HHmmss.txt', where YYYYMMDD_HHmmss is the timestamp at the time of the script run.
    After creating the log file, the function writes an audit trail entry including the timestamp, username, and machine name.

.NOTES
    This function depends on two other functions: Write-CustomError and Write-CustomOutput. These functions are used to output messages to the console and the log file.
    The $scriptRoot global variable should be defined before calling this function. It should contain the path of the script's directory.
    The $script:logFilePath script variable is set by this function. It contains the full path of the created log file.

.EXAMPLE
    New-LogFile

    This will create a new log file in the 'Log' directory under $scriptRoot with the current timestamp, and then write an audit trail entry to it.

#>
function New-LogFile {
    try {
        # Join the script root directory with 'Log' to get the log directory path
        $logDir = Join-Path -Path $scriptRoot -ChildPath 'Log'
        
        # If the log directory does not exist, create it
        if (!(Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir | Out-Null
        }

        # Get the current timestamp and format it as 'yyyyMMdd_HHmmss'
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        # Join the log directory path with the log file name to get the full log file path
        $script:logFilePath = Join-Path -Path $logDir -ChildPath "Log_$timestamp.txt"

        # Get the machine name and username
        $machineName = $env:COMPUTERNAME
        $userName = $env:USERNAME
        # Create the audit trail message and write it to the log file
        $auditMessage = "Log file created on $timestamp by $userName on machine $machineName"
        Write-CustomOutput $auditMessage

    } catch {
        # If an error occurs, write the error message to the log file
        Write-CustomError "Error occurred while creating new log file: $_"
    }
}


<#
.SYNOPSIS
    This function is used to display successful operations.

.DESCRIPTION
    The Write-CustomSuccess function generates a success message, which is written to the console in green color 
    indicating successful operations. The message, along with a timestamp and the success level, is also added 
    to a log file.

.PARAMETER Message
    This parameter specifies the success message to be displayed and logged.

.EXAMPLE
    Write-CustomSuccess "The operation was successful."
    Displays and logs the message "The operation was successful." as a success message.

.INPUTS
    String
    Takes a string as an input which is the success message to be displayed and logged.

.OUTPUTS
    None
    This function does not return a value. It writes the success message to the console and appends it to the log file.

.NOTES
    This function is useful for providing clear success indications during script execution.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/add-content
#>
function Write-CustomSuccess {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message
    )

    # Define success color and level
    $Color = "Green"
    $Level = "SUCCESS"
    
    # Get the current timestamp and format it as 'yyyyMMdd_HHmmss'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    # Create the log message including the timestamp, the log level and the message
    $logMessage = "$timestamp - $Level - $Message" # Include log level in log message
    
    # Write the success message to the console with the green color
    Write-Host $Message -ForegroundColor $Color
    Write-Host ""
    # Append the log message to the log file
    Add-Content -Path $script:logFilePath -Value $logMessage
}





<#
.SYNOPSIS
    Writes a custom log message with a timestamp and log level to the console and the log file.

.DESCRIPTION
    The Write-CustomOutput function is used to write custom log messages. Each message is outputted to the console and appended to the log file. 
    The function adds a timestamp and a log level to each message.
    The message color in the console can be customized.

.PARAMETER Message
    The message to be written to the console and the log file. 

.PARAMETER Color
    The color of the message in the console. The default value is 'White'. This parameter accepts any color name that is valid in the Write-Host cmdlet.

.PARAMETER Level
    The log level of the message. The default value is 'INFO'. This parameter can be used to indicate the importance or the type of the message (like 'ERROR', 'WARNING', 'DEBUG', etc.).

.NOTES
    The $script:logFilePath script variable should be defined before calling this function. It should contain the full path of the log file.
    
.EXAMPLE
    Write-CustomOutput -Message "Script started." -Color "Green" -Level "INFO"

    This will output the message "Script started." in green color to the console, and append the log message "YYYYMMDD_HHmmss - INFO - Script started." to the log file.

#>
function Write-CustomOutput {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "White",
        [string]$Level = "INFO" # Add log level parameter
    )
    
    # Get the current timestamp and format it as 'yyyyMMdd_HHmmss'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    # Create the log message including the timestamp, the log level and the message
    $logMessage = "$timestamp - $Level - $Message" # Include log level in log message
    
    # Write the message to the console with the specified color
    Write-Host $Message -ForegroundColor $Color
    Write-Host ""
    # Append the log message to the log file
    Add-Content -Path $script:logFilePath -Value $logMessage
}


<#
.SYNOPSIS
    Writes a custom error message with a timestamp, log level, and line number to the console and the log file.

.DESCRIPTION
    The Write-CustomError function is used to write custom error messages. Each message is outputted to the console in red (by default) and appended to the log file.
    The function adds a timestamp, a log level, and the script line number where the error occurred to each error message.

.PARAMETER Message
    The error message to be written to the console and the log file.

.PARAMETER Color
    The color of the error message in the console. The default value is 'Red'. This parameter accepts any color name that is valid in the Write-Host cmdlet.

.PARAMETER Level
    The log level of the error message. The default value is 'ERROR'. This parameter can be used to indicate the importance or the type of the message (like 'CRITICAL', 'FATAL', etc.).

.NOTES
    The $script:logFilePath script variable should be defined before calling this function. It should contain the full path of the log file.
    
.EXAMPLE
    Write-CustomError -Message "File not found." -Color "Red" -Level "ERROR"

    This will output the error message "File not found." in red color to the console, and append the error message "YYYYMMDD_HHmmss - ERROR - File not found.\nError occurred on line: X" to the log file, where X is the script line number where the error occurred.

#>
function Write-CustomError {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "Red",
        [string]$Level = "ERROR" # Add log level parameter
    )
    
    # Get the current timestamp and format it as 'yyyyMMdd_HHmmss'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    # Create the error message including the timestamp, the log level and the error message
    $errorMessage = "$timestamp - $Level - $Message" # Include log level in log message

    # Add the script line where the error occurred
    $line = $_.InvocationInfo.ScriptLineNumber
    $errorMessage += "`nError occurred on line: $line"

    # Write the error message to the console with the specified color
    Write-Host $Message -ForegroundColor $Color
    # Append the error message to the log file
    Add-Content -Path $script:logFilePath -Value $errorMessage
}



<#
.SYNOPSIS
    Deletes old log files, leaving only a specified maximum number of log files in the directory.

.DESCRIPTION
    The Clear-OldLogs function removes old log files in the specified directory, leaving only the latest logs up to the maximum log count.
    The function sorts the log files by their last write time, and the oldest files get removed first if the total log count exceeds the specified maximum log count.

.PARAMETER LogDirectory
    The directory where the log files are located. By default, this is the 'Log' subdirectory in the $scriptRoot directory.

.PARAMETER MaxLogCount
    The maximum number of log files to be kept in the directory. If the total log count exceeds this number, the oldest logs will be deleted. By default, this value is taken from the $script:MaxLogCount script variable.

.NOTES
    The $script:MaxLogCount and $scriptRoot script variables should be defined before calling this function.

.EXAMPLE
    Clear-OldLogs -LogDirectory "C:\Scripts\Log" -MaxLogCount 10

    This will delete the oldest log files in the "C:\Scripts\Log" directory, leaving only the 10 latest log files.

#>
function Clear-OldLogs {
    [CmdletBinding()]
    param (
        [string]$LogDirectory = (Join-Path -Path $scriptRoot -ChildPath 'Log'),
        [int]$MaxLogCount = $script:MaxLogCount # You can adjust this as needed
    )
    
    # Get all log files in the directory and sort them by last write time
    $oldLogs = Get-ChildItem -Path $LogDirectory -Filter 'Log_*.txt' | Sort-Object -Property LastWriteTime
    
    # If the total log count exceeds the maximum log count, delete the oldest logs
    if ($oldLogs.Count -gt $MaxLogCount) {
        $logsToDelete = $oldLogs.Count - $MaxLogCount
        $oldLogs | Select-Object -First $logsToDelete | ForEach-Object {
            Write-CustomOutput "Deleting old log file: $($_.FullName)"
            Remove-Item -Path $_.FullName -Force
        }
    }
}





<#
.SYNOPSIS
    Opens a file dialog for selecting a settings file and imports its content based on the file type.

.DESCRIPTION
    The Get-SettingsFileDialog function presents a file dialog that allows the user to select a settings file. 
    The settings file can be in XML, JSON, or CSV format. 
    Once a file is selected, the function reads the file's content based on the file type and returns the imported data.

.PARAMETER None
    This function doesn't take any parameters.

.NOTES
    This function requires the .NET System.Windows.Forms namespace to be loaded into the PowerShell session. 
    The function uses the OpenFileDialog class to open the file dialog.

.EXAMPLE
    $settings = Get-SettingsFileDialog

    This will open a file dialog, and after the user selects a settings file, the function will import the content 
    of the file and store it in the $settings variable.

#>
function Get-SettingsFileDialog {
    try {
        # Create OpenFileDialog object
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.InitialDirectory = $scriptRoot
        $dialog.Filter = "XML files (*.xml)|*.xml|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv"
        $dialog.Title = "Select a settings file"

        # Show the OpenFileDialog and check if the user clicked 'OK'
        if ($dialog.ShowDialog() -eq 'OK') {
            $selectedFilePath = $dialog.FileName

            # Import the file's content based on the file type
            switch ($dialog.FilterIndex) {
                1 { # XML
                    Import-Clixml -Path $selectedFilePath
                }
                2 { # JSON
                    Get-Content $selectedFilePath | ConvertFrom-Json
                }
                3 { # CSV
                    Import-Csv -Path $selectedFilePath
                }
                default {
                    throw "Invalid file type selected"
                }
            }
        } else {
            $null
        }
    } catch {
        Write-CustomError "Error occurred while selecting the settings file: $_"
    }
}




<#
.SYNOPSIS
    Prompts the user for registry key path, value name, and value data and validates the inputs.

.DESCRIPTION
    The Get-NewParameters function asks the user to enter a registry path, a value name, and value data. 
    It then validates each of these inputs to ensure they're not empty and that the registry path starts with a valid hive. 
    If validation passes, the function returns a custom object with properties for the registry path, value name, and value data.

.PARAMETER None
    This function doesn't take any parameters. It relies on user input collected via the Read-Host cmdlet.

.NOTES
    The function uses regular expressions to validate the registry path.

.EXAMPLE
    $newParameters = Get-NewParameters

    This example prompts the user to enter the registry path, value name, and value data, 
    and then stores the validated inputs in the $newParameters variable.

#>
function Get-NewParameters {
    # Prompt for registry path and validate
    $regPath = Read-Host -Prompt 'Enter the registry path'
    if ($regPath -notmatch '^(HKLM:|HKCU:|HKCR:|HKU:|HKCC:|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER|HKEY_CLASSES_ROOT|HKEY_USERS|HKEY_CURRENT_CONFIG)\\') {
        Write-CustomError "The registry path is not valid. It should start with a valid hive (like 'HKLM:' or 'HKEY_LOCAL_MACHINE')."
        exit 1
    }

    # Prompt for value name and validate
    $valueName = Read-Host -Prompt 'Enter the value name'
    if ([string]::IsNullOrEmpty($valueName)) {
        Write-CustomError "The value name cannot be empty."
        exit 1
    }

    # Prompt for value data and validate
    $valueData = Read-Host -Prompt 'Enter the value data'
    if ([string]::IsNullOrEmpty($valueData)) {
        Write-CustomError "The value data cannot be empty."
        exit 1
    }

    # Return new parameters as a custom object
    [PSCustomObject]@{
        RegPath = $regPath
        ValueName = $valueName
        ValueData = $valueData
    }
}



<#
.SYNOPSIS
    Creates a new settings file in XML, JSON, or CSV format with provided parameters.

.DESCRIPTION
    The New-SettingsFile function takes an object of parameters and a file type, then creates a new settings file in the given format. 
    The name of the new file is determined by incrementing the highest number among existing settings files.

.PARAMETER parameters
    A mandatory parameter representing the settings to be stored in the file.

.PARAMETER FileType
    An optional parameter representing the format of the settings file. It must be one of the following values: "XML", "JSON", or "CSV". The default is "XML".

.EXAMPLE
    $newSettings = Get-NewParameters
    New-SettingsFile -parameters $newSettings -FileType "JSON"

    This example first gets new parameters interactively from the user using the Get-NewParameters function. 
    Then, it creates a new settings file in JSON format with those parameters.

.NOTES
    The function automatically determines the name of the new file by finding the highest numbered existing settings file and incrementing by 1.
    In case of an error, the function calls the Write-CustomError function to log the error message.
#>
function New-SettingsFile {
    [CmdletBinding()]
    param (
        # Mandatory parameter representing the settings to be stored in the file.
        [Parameter(Mandatory=$true)]
        $parameters,

        # Optional parameter representing the format of the settings file.
        # It must be one of the following values: "XML", "JSON", or "CSV". The default is "XML".
        [ValidateSet("XML", "JSON", "CSV")]
        [string]$FileType = "XML"
    )

    try {
        # Get the base names of existing settings files and find the highest number
        $settingsFiles = Get-ChildItem -Path $scriptRoot -Filter "Settings*.$FileType" | Select-Object -ExpandProperty BaseName
        $maxNumber = ($settingsFiles | ForEach-Object { ($_ -replace '[^\d]', '') -as [int] } | Sort-Object -Descending | Select-Object -First 1) + 1

        # Create the name and path for the new settings file
        $newSettingsFileName = "Settings$maxNumber.$FileType"
        $newSettingsFilePath = Join-Path -Path $scriptRoot -ChildPath $newSettingsFileName

        # Create the new settings file in the given format
        switch ($FileType) {
            "XML" {
                $parameters | Export-Clixml -Path $newSettingsFilePath
            }
            "JSON" {
                $parameters | ConvertTo-Json | Out-File -FilePath $newSettingsFilePath
            }
            "CSV" {
                $parameters | Export-Csv -Path $newSettingsFilePath -NoTypeInformation
            }
        }

        # Return the path of the new settings file
        return $newSettingsFilePath
    } catch {
        # If an error occurs, log the error message
        Write-CustomError "Error occurred while creating new settings file: $_"
    }
}




<#
.SYNOPSIS
    Updates the user profiles with a given parameter set and includes a retry mechanism for failures.

.DESCRIPTION
    The Update-UserProfiles function updates user profiles by making changes to the registry. 
    If an error is encountered, the function will retry the update up to a maximum number of times (default is 3).

.PARAMETER parameters
    A mandatory parameter that should be a custom object containing registry path (RegPath), value name (ValueName), and value data (ValueData).

.PARAMETER maxRetries
    An optional parameter that defines the maximum number of retries if an error occurs during the update process. 
    The default value is 3.

.EXAMPLE
    $params = [PSCustomObject]@{
        RegPath = 'HKLM:\Software\MySoftware'
        ValueName = 'MyValue'
        ValueData = 'MyData'
    }
    Update-UserProfiles -parameters $params -maxRetries 5

    This example updates user profiles with the specified parameters and sets the maximum retries to 5.

.NOTES
    This function is designed to update user profiles in a robust manner, handling temporary errors by retrying the update operation.
    It writes progress and error information to the log through the Write-CustomOutput and Write-CustomError functions.
#>
function Update-UserProfiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $parameters,
        [int]$maxRetries = 3  # maximum number of retries
    )

    Write-CustomOutput "Starting to update user profiles..." # Log start of function

    Get-ChildItem "Registry::HKEY_USERS" | Where-Object { $_.PSChildName -match "S-1-5-21" } | ForEach-Object {
        $user = $_
        $retryCount = 0
        
        while ($retryCount -lt $maxRetries) {
            try {
                $userRegPath = "Registry::HKEY_USERS\$($user.PSChildName)\$($parameters.RegPath)"
                Write-CustomOutput "Processing user $($user.PSChildName)..." # Log start of user processing

                if (Test-Path -Path $userRegPath) {
                    if (Get-ItemProperty -Path $userRegPath -Name $parameters.ValueName -ErrorAction SilentlyContinue) {
                        Write-CustomOutput "The registry value '$($parameters.ValueName)' is already set for user $($user.PSChildName). Skipping..."
                        break
                    }

                    Write-CustomOutput "Setting the registry value for user $($user.PSChildName)..."
                    Set-ItemProperty -Path $userRegPath -Name $parameters.ValueName -Value $parameters.ValueData -Type DWord
                    Write-CustomOutput "Successfully set the registry value for user $($user.PSChildName)." # Log success of operation
                    break
                } else {
                    Write-CustomOutput "The registry path does not exist for user $($user.PSChildName). Skipping..."
                    break
                }
            } catch {
                Write-CustomError "Error encountered while processing user $($user.PSChildName): $_"
                $retryCount++
                Write-CustomOutput "Attempt $($retryCount) failed. Retrying..."
                if ($retryCount -eq $maxRetries) {
                    Write-CustomError "Maximum retries reached for user $($user.PSChildName)."
                    Write-CustomError "An error occurred during the execution of the script: $_"
                    if ($backupFilePath) {
                        Restore-Registry -BackupFilePath $backupFilePath
                        Write-CustomOutput "Successfully restored the registry from backup."
                }
                Start-Sleep -Seconds (2*$retryCount)  # wait a bit before next retry, with backoff
                }
            }
        }
    }
}




<#
.SYNOPSIS
    Prompts the user to confirm an action by displaying a message and expecting a 'Y' or 'N' response.

.DESCRIPTION
    The Confirm-Action function prompts the user to confirm an action by displaying a message and expecting a 'Y' or 'N' response.
    It returns $true if the user responds with 'Y' (yes), and $false if the user responds with 'N' (no).

.PARAMETER Message
    The message to display when asking for user confirmation.

.EXAMPLE
    if (Confirm-Action -Message "Are you sure you want to delete the file? (Y/N)") {
        # Perform deletion
    } else {
        # Abort action
    }

    This example prompts the user with the specified message and proceeds with the deletion if the user responds with 'Y'.

.NOTES
    - The user's response is case-insensitive.
    - Only 'Y' and 'N' responses are accepted. Any other response will be considered as 'N'.
    - It is recommended to use this function when requiring user confirmation to avoid unintended actions.

.FUNCTIONALITY
    Interactive Prompts

.INPUTS
    None. You cannot pipe input to Confirm-Action.

.OUTPUTS
    System.Boolean. The function outputs a Boolean value ($true or $false) based on the user's response.

#>
function Confirm-Action {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )

    $prompt = (Read-Host -Prompt $Message).ToUpper()
    if ($prompt -eq 'Y') {
        $true
    } else {
        $false
    }
}



<#
.SYNOPSIS
    Creates a backup of the registry.

.DESCRIPTION
    The Backup-Registry function creates a backup of the registry by invoking the 'regedit.exe' command-line tool with the appropriate arguments.

.PARAMETER BackupDirectory
    An optional parameter specifying the directory where the backup file will be stored.
    If not provided, the backup file will be saved in a 'Backup' subdirectory of the script's root directory.

.EXAMPLE
    Backup-Registry -BackupDirectory "C:\Backups"

    This example creates a backup of the registry and saves the backup file in the 'C:\Backups' directory.

.NOTES
    - This function utilizes the 'regedit.exe' command-line tool, which is a built-in Windows utility.
    - The backup file will be named 'RegistryBackup_<timestamp>.reg', where <timestamp> represents the current date and time.
    - The progress of the backup process is displayed with a loading message.
#>
function Backup-Registry {
    [CmdletBinding()]
    param (
        [string]$BackupDirectory = (Join-Path -Path $scriptRoot -ChildPath 'Backup')
    )

    try {
        if (!(Test-Path $BackupDirectory)) {
            New-Item -ItemType Directory -Path $BackupDirectory | Out-Null
        }
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupFilePath = Join-Path -Path $BackupDirectory -ChildPath "RegistryBackup_$timestamp.reg"

        Write-Host "Starting backup..."

        $startTime = Get-Date

        # Start the registry backup process
        $process = Start-Process -FilePath "regedit.exe" -ArgumentList "/E", "`"$backupFilePath`"" -NoNewWindow -PassThru

        $animationChars = '/-\|'
        $animationIndex = 0

        # Display a progress bar with animation and elapsed time while the process is running
        while (!$process.HasExited) {
            $animationChar = $animationChars[$animationIndex % $animationChars.Length]
            $elapsedTime = (Get-Date) - $startTime
            $status = "In progress... $($animationChar)  |`nElapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))"
            Write-Progress -Activity "Backing up registry" -Status $status -Id 1
            Start-Sleep -Milliseconds 200
            $animationIndex++
        }

        $elapsedTime = (Get-Date) - $startTime
        $status = "Complete `nElapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))"
        Write-Progress -Activity "Backing up registry" -Status $status -Completed

        Write-Host "Backup complete."

        return $backupFilePath
    } catch {
        Write-Host "Error occurred while backing up the registry: $_"
    }
}





<#
.SYNOPSIS
    Restores the registry from a backup file.

.DESCRIPTION
    The Restore-Registry function restores the registry by running the regedit.exe command with the specified backup file.

.PARAMETER BackupFilePath
    The mandatory parameter that specifies the path to the backup file.

.NOTES
    This function provides a convenient way to restore the registry from a backup file created using the Backup-Registry function.
    It displays progress information during the restore process and provides support for ShouldProcess to confirm the restoration action.

.EXAMPLE
    $backupFile = "C:\Backup\RegistryBackup_20230101_123456.reg"
    Restore-Registry -BackupFilePath $backupFile

    This example restores the registry from the specified backup file.

#>
function Restore-Registry {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [Parameter(Mandatory=$true)]
        [string]$BackupFilePath
    )

    if ($PSCmdlet.ShouldProcess("Are you sure you want to restore the registry from the backup file at '$BackupFilePath'?", "Restoring registry", "Restoring registry")) {
        try {
            Write-Host "Starting registry restore from backup..."

            $loadingChars = '|/-\'
            $index = 0

            # Start the registry restore process
            $process = Start-Process -FilePath "regedit.exe" -ArgumentList "/S", "`"$BackupFilePath`"" -NoNewWindow -PassThru

            # Display a loading message while the process is running
            while (!$process.HasExited) {
                Write-Host "`rRestoring registry from backup... $($loadingChars[$index % $loadingChars.Length])" -NoNewline
                Start-Sleep -Milliseconds 200
                $index++
            }

            Write-Host "`rRestore complete."

        } catch {
            Write-Host "Error occurred while restoring the registry from backup: $_"
        }
    }
}




<#
.SYNOPSIS
    Main script to manipulate registry settings of all user profiles on the computer.

.DESCRIPTION
    The main script provides options to import settings from an XML file or enter them manually.
    It also allows creating a backup before applying the changes to the registry settings of user profiles.
    The script performs various operations such as checking PowerShell and Windows versions, creating log files,
    managing old log files, getting settings from the user, creating new settings files, executing registry updates,
    and handling errors.

.PARAMETER N/A
    This script doesn't accept any parameters. All inputs are interactive.

.NOTES
    General notes:
    - Please run the script with administrative privileges to ensure it can access and modify the registry.
    - Always confirm that you have a recent backup of your registry before running the script, in case any issues arise.
    - This script was last updated on 2023-06-25.

.LINK
    https://github.com/Paul1404/registry-changer
#>

try {
    Add-Type -AssemblyName System.Windows.Forms

    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $savedParameters = Get-SettingsFileDialog

    # Call the New-LogFile function at the start of the script
    New-LogFile
    Clear-OldLogs

    Get-PowerShellVersion
    Get-WindowsVersion
    Get-AdminPrivileges
    Get-NetworkConnectivity
    Get-DiskSpace -DriveLetter "C:" -MinSpaceGB 1

    if ($null -eq $savedParameters) {
        $savedParameters = Get-NewParameters
        $settingsFilePath = New-SettingsFile -parameters $savedParameters
        Write-CustomOutput "Parameters saved to file: $settingsFilePath"
    }

    Write-Host "`nSelected parameters:" -ForegroundColor Yellow
    $formatTable = @{Expression={$_.RegPath};Label="RegPath";width=75},@{Expression={$_.ValueName};Label="ValueName";width=15},@{Expression={$_.ValueData};Label="ValueData";width=10}
    $savedParameters | Format-Table $formatTable

    if (Confirm-Action -Message 'Do you want to proceed with these settings? (Y/N)') {
        if (Confirm-Action -Message 'Do you want to backup the registry? (Y/N)') {
            $backupFilePath = Backup-Registry
            Write-CustomSuccess "`nRegistry has been backed up. Backup file: $backupFilePath"
        }
        Update-UserProfiles -parameters $savedParameters
        Write-CustomSuccess "Script execution complete."
    } else {
        Write-CustomOutput "Script execution cancelled by user."
    }
} catch {
    Write-CustomError "An error occurred during the execution of the script: $_"
} finally {
    Read-Host "Press Enter to exit"
}
