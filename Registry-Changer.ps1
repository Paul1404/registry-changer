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
                Write-Output "The registry value '$($parameters.ValueName)' is already set for user $($user.PSChildName). Skipping..."
                return
            }

            if (!(Test-Path $userRegPath)) {
                Write-Output "Creating a new registry key for user $($user.PSChildName)..."
                New-Item -Path $userRegPath -Force | Out-Null
            }

            Write-Output "Setting the registry value for user $($user.PSChildName)..."
            Set-ItemProperty -Path $userRegPath -Name $parameters.ValueName -Value $parameters.ValueData -Type DWord
        } catch {
            Write-Error "Error encountered while processing user $($user.PSChildName): $_"
        }
    }
}

Add-Type -AssemblyName System.Windows.Forms

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$savedParametersFile = Join-Path -Path $scriptRoot -ChildPath 'SavedParameters.xml'

if (Test-Path $savedParametersFile) {
    $savedParameters = Import-Clixml -Path $savedParametersFile
    Write-Output 'Current saved parameters:'
    $savedParameters | Format-Table -AutoSize
} else {
    $savedParameters = Show-SettingsFileDialog
}

if ($null -eq $savedParameters) {
    $savedParameters = Read-NewParameters
    New-SettingsFile -parameters $savedParameters
    Write-Output "Parameters saved to file: $settingsFilePath"
}

Update-UserProfiles -parameters $savedParameters

Write-Output "Script execution complete."
Read-Host "Press Enter to exit"