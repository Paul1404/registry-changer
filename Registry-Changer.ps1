function Write-CustomOutput {
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    
    Write-Host $Message -ForegroundColor $Color
    Write-Host ""
}

function Write-CustomError {
    param (
        [string]$Message
    )
    
    Write-Host $Message -ForegroundColor Red
}

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

function New-SettingsFile($parameters) {
    $settingsFiles = Get-ChildItem -Path $scriptRoot -Filter 'Settings*.xml' | Select-Object -ExpandProperty BaseName
    $maxNumber = ($settingsFiles | ForEach-Object { ($_ -replace '[^0-9]', '') -as [int] } | Sort-Object -Descending)[0]

    $nextNumber = $maxNumber + 1
    $settingsFileName = "Settings_$nextNumber.xml"
    $settingsFilePath = Join-Path -Path $scriptRoot -ChildPath $settingsFileName

    $parameters | Export-Clixml -Path $settingsFilePath
}

function Update-UserProfiles($parameters) {
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

function Confirm-Proceed {
    $prompt = Read-Host "Do you want to proceed with these settings? (Y/N)"
    if ($prompt -eq 'Y') {
        $true
    } else {
        $false
    }
}

function Confirm-Backup {
    $prompt = Read-Host "Do you want to backup the registry? (Y/N)"
    if ($prompt -eq 'Y') {
        $true
    } else {
        $false
    }
}

function Backup-Registry {
    $backupDir = Join-Path -Path $scriptRoot -ChildPath 'Backup'
    if (!(Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupFilePath = Join-Path -Path $backupDir -ChildPath "RegistryBackup_$timestamp.reg"

    Write-CustomOutput "Starting backup..."
    
    # Start the backup operation in a separate process and wait for it to finish
    Start-Process -FilePath "regedit.exe" -ArgumentList "/E", "`"$backupFilePath`"" -NoNewWindow -Wait

    Write-CustomOutput "Backup complete."
}



Add-Type -AssemblyName System.Windows.Forms

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$savedParameters = Get-SettingsFileDialog

if ($null -eq $savedParameters) {
    $savedParameters = Get-NewParameters
    New-SettingsFile -parameters $savedParameters
    Write-CustomOutput "Parameters saved to file: $settingsFilePath"
}

Write-CustomOutput 'Current saved parameters:'
$savedParameters | Format-Table -AutoSize

if (Confirm-Proceed) {
    if (Confirm-Backup) {
        Backup-Registry
        Write-CustomOutput "Registry has been backed up."
    }
    Update-UserProfiles -parameters $savedParameters
    Write-CustomOutput "Script execution complete."
} else {
    Write-CustomOutput "Script execution cancelled by user."
}

Read-Host "Press Enter to exit"
