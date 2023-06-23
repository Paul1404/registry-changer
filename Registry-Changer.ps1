<#
.SYNOPSIS
This script allows to manipulate registry settings of all users profiles on the computer.
.DESCRIPTION
The script provides options to import settings from an XML file or enter them manually. 
It also provides options to create a backup before applying the changes.
#>

# Custom functions to write output and errors
function Write-CustomOutput {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    
    Write-Host $Message -ForegroundColor $Color
    Write-Host ""
}

function Write-CustomError {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "Red"
    )
    
    Write-Host $Message -ForegroundColor $Color
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
    $valueName = Read-Host -Prompt 'Enter the value name'
    $valueData = Read-Host -Prompt 'Enter the value data'

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

    Get-ChildItem "Registry::HKEY_USERS" | Where-Object { $_.PSChildName -match "S-1-5-21" } | ForEach-Object {
        $user = $_
        try {
            $userRegPath = "Registry::HKEY_USERS\$($user.PSChildName)\$($parameters.RegPath)"

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
