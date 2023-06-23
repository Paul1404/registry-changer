function CustomWrite-Output {
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message
    )
    
    begin {}
    process {
        Write-Output $Message
        Write-Output ""
    }
    end {}
}

function Show-SettingsFileDialog {
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

function Read-NewParameters {
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
                CustomWrite-Output "The registry value '$($parameters.ValueName)' is already set for user $($user.PSChildName). Skipping..."
                return
            }

            if (!(Test-Path $userRegPath)) {
                CustomWrite-Output "Creating a new registry key for user $($user.PSChildName)..."
                New-Item -Path $userRegPath -Force | Out-Null
            }

            CustomWrite-Output "Setting the registry value for user $($user.PSChildName)..."
            Set-ItemProperty -Path $userRegPath -Name $parameters.ValueName -Value $parameters.ValueData -Type DWord
        } catch {
            Write-Error "Error encountered while processing user $($user.PSChildName): $_"
        }
    }
}

function Ask-To-Proceed {
    $prompt = Read-Host "Do you want to proceed with these settings? (Y/N)"
    if ($prompt -eq 'Y') {
        $true
    } else {
        $false
    }
}

function Ask-To-Backup {
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

    CustomWrite-Output "Starting backup..."
    
    # Start the backup operation in a separate process and wait for it to finish
    Start-Process -FilePath "regedit.exe" -ArgumentList "/E", "`"$backupFilePath`"" -NoNewWindow -Wait

    CustomWrite-Output "Backup complete."
}



Add-Type -AssemblyName System.Windows.Forms

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$savedParameters = Show-SettingsFileDialog

if ($null -eq $savedParameters) {
    $savedParameters = Read-NewParameters
    New-SettingsFile -parameters $savedParameters
    CustomWrite-Output "Parameters saved to file: $settingsFilePath"
}

CustomWrite-Output 'Current saved parameters:'
$savedParameters | Format-Table -AutoSize

if (Ask-To-Proceed) {
    if (Ask-To-Backup) {
        Backup-Registry
        CustomWrite-Output "Registry has been backed up."
    }
    Update-UserProfiles -parameters $savedParameters
    CustomWrite-Output "Script execution complete."
} else {
    CustomWrite-Output "Script execution cancelled by user."
}

Read-Host "Press Enter to exit"
