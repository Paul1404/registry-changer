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


# Global variable to hold the log file path
$logFilePath = ""

function Get-Environment {
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-CustomError "This script requires PowerShell version 5.1 or higher. Your version is $($PSVersionTable.PSVersion)."
        exit 1
    } else {
        Write-CustomOutput "PowerShell version check passed. Your version is $($PSVersionTable.PSVersion)."
    }

    # Check Windows version
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $version = [Version]$osInfo.Version

    if ($version.Major -lt 10) {
        Write-CustomError "This script requires Windows version 10.0 or higher. Your version is $($osInfo.Caption) $($osInfo.Version)."
        exit 1
    } else {
        Write-CustomOutput "Windows version check passed. Your version is $($osInfo.Caption) $($osInfo.Version)."
    }
}


# Function to create a new log file for each run
function New-LogFile {
    $logDir = Join-Path -Path $scriptRoot -ChildPath 'Log'
    
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }
    
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:logFilePath = Join-Path -Path $logDir -ChildPath "Log_$timestamp.txt"
}

# Custom functions to write output and errors
function Write-CustomOutput {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "White",
        [string]$Level = "INFO" # Add log level parameter
    )
    
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logMessage = "$timestamp - $Level - $Message" # Include log level in log message
    
    Write-Host $Message -ForegroundColor $Color
    Write-Host ""
    Add-Content -Path $script:logFilePath -Value $logMessage
}

function Write-CustomError {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "Red",
        [string]$Level = "ERROR" # Add log level parameter
    )
    
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $errorMessage = "$timestamp - $Level - $Message" # Include log level in log message

    # Add script line where the error occurred
    $line = $_.InvocationInfo.ScriptLineNumber
    $errorMessage += "`nError occurred on line: $line"

    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $script:logFilePath -Value $errorMessage
}

# A function to manage old log files
function Clear-OldLogs {
    [CmdletBinding()]
    param (
        [string]$LogDirectory = (Join-Path -Path $scriptRoot -ChildPath 'Log'),
        [int]$MaxLogCount = 100 # You can adjust this as needed
    )
    
    $oldLogs = Get-ChildItem -Path $LogDirectory -Filter 'Log_*.txt' | Sort-Object -Property LastWriteTime
    
    if ($oldLogs.Count -gt $MaxLogCount) {
        $logsToDelete = $oldLogs.Count - $MaxLogCount
        $oldLogs | Select-Object -First $logsToDelete | Remove-Item -Force
    }
}





# Function to select a settings XML file
function Get-SettingsFileDialog {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $scriptRoot
    $dialog.Filter = "XML files (*.xml)|*.xml"
    $dialog.Title = "Select a settings XML file"

    if ($dialog.ShowDialog() -eq 'OK') {
        $selectedFilePath = $dialog.FileName
        Import-Clixml -Path $selectedFilePath
    } else {
        $null
    }
}

# Function to read new parameters from the user
function Get-NewParameters {
    $regPath = Read-Host -Prompt 'Enter the registry path'
    if ($regPath -notmatch '^(HKLM:|HKCU:|HKCR:|HKU:|HKCC:|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER|HKEY_CLASSES_ROOT|HKEY_USERS|HKEY_CURRENT_CONFIG)\\') {
        Write-CustomError "The registry path is not valid. It should start with a valid hive (like 'HKLM:' or 'HKEY_LOCAL_MACHINE')."
        exit 1
    }

    $valueName = Read-Host -Prompt 'Enter the value name'
    if ([string]::IsNullOrEmpty($valueName)) {
        Write-CustomError "The value name cannot be empty."
        exit 1
    }

    $valueData = Read-Host -Prompt 'Enter the value data'
    if ([string]::IsNullOrEmpty($valueData)) {
        Write-CustomError "The value data cannot be empty."
        exit 1
    }

    [PSCustomObject]@{
        RegPath = $regPath
        ValueName = $valueName
        ValueData = $valueData
    }
}


# Function to create a new settings file
function New-SettingsFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $parameters
    )

    $settingsFiles = Get-ChildItem -Path $scriptRoot -Filter 'Settings*.xml' | Select-Object -ExpandProperty BaseName
    $maxNumber = ($settingsFiles | ForEach-Object { ($_ -replace '[^0-9]', '') -as [int] } | Sort-Object -Descending)[0]

    $nextNumber = $maxNumber + 1
    $settingsFileName = "Settings_$nextNumber.xml"
    $settingsFilePath = Join-Path -Path $scriptRoot -ChildPath $settingsFileName

    $parameters | Export-Clixml -Path $settingsFilePath
    return $settingsFilePath
}

# Function to update user profiles
function Update-UserProfiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $parameters
    )

    Write-CustomOutput "Starting to update user profiles..." # Log start of function

    Get-ChildItem "Registry::HKEY_USERS" | Where-Object { $_.PSChildName -match "S-1-5-21" } | ForEach-Object {
        $user = $_
        try {
            $userRegPath = "Registry::HKEY_USERS\$($user.PSChildName)\$($parameters.RegPath)"

            Write-CustomOutput "Processing user $($user.PSChildName)..." # Log start of user processing

            if (Get-ItemProperty -Path $userRegPath -Name $parameters.ValueName -ErrorAction SilentlyContinue) {
                Write-CustomOutput "The registry value '$($parameters.ValueName)' is already set for user $($user.PSChildName). Skipping..."
                return
            }

            if (!(Test-Path $userRegPath)) {
                Write-CustomOutput "Creating a new registry key for user $($user.PSChildName)..."
                New-Item -Path $userRegPath -Force | Out-Null
            }

            Write-CustomOutput "Setting the registry value for user $($user.PSChildName)..."
            Set-ItemProperty -Path $userRegPath -Name $parameters.ValueName -Value $parameters.ValueData -Type DWord
            Write-CustomOutput "Successfully set the registry value for user $($user.PSChildName)." # Log success of operation
        } catch {
            Write-CustomError "Error encountered while processing user $($user.PSChildName): $_"
        }
    }
}


# Function to confirm user action
function Confirm-Action {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )

    $prompt = Read-Host -Prompt $Message
    if ($prompt -eq 'Y') {
        $true
    } else {
        $false
    }
}

# Function to backup registry
function Backup-Registry {
    [CmdletBinding()]
    param (
        [string]$BackupDirectory = (Join-Path -Path $scriptRoot -ChildPath 'Backup')
    )

    if (!(Test-Path $BackupDirectory)) {
        New-Item -ItemType Directory -Path $BackupDirectory | Out-Null
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupFilePath = Join-Path -Path $BackupDirectory -ChildPath "RegistryBackup_$timestamp.reg"

    Write-CustomOutput "Starting backup..."
    Start-Process -FilePath "regedit.exe" -ArgumentList "/E", "`"$backupFilePath`"" -NoNewWindow -Wait
    Write-CustomOutput "Backup complete."
    
    return $backupFilePath
}

# Main script

Add-Type -AssemblyName System.Windows.Forms

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$savedParameters = Get-SettingsFileDialog

# Call the New-LogFile function at the start of the script
New-LogFile
Clear-OldLogs

Get-Environment

if ($null -eq $savedParameters) {
    $savedParameters = Get-NewParameters
    $settingsFilePath = New-SettingsFile -parameters $savedParameters
    Write-CustomOutput "Parameters saved to file: $settingsFilePath"
}

Write-CustomOutput 'Current saved parameters:'
$savedParameters | Format-Table -AutoSize

if (Confirm-Action -Message 'Do you want to proceed with these settings? (Y/N)') {
    if (Confirm-Action -Message 'Do you want to backup the registry? (Y/N)') {
        $backupFilePath = Backup-Registry
        Write-CustomOutput "Registry has been backed up. Backup file: $backupFilePath"
    }
    Update-UserProfiles -parameters $savedParameters
    Write-CustomOutput "Script execution complete."
} else {
    Write-CustomOutput "Script execution cancelled by user."
}

Read-Host "Press Enter to exit"
