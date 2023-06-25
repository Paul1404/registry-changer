@echo off

echo This script allows you to run a PowerShell script with elevated privileges.
echo It ensures that the user has the necessary permissions to execute the script.
echo.

:: Use PowerShell to prompt for file selection
for /f "delims=" %%I in ('powershell.exe -noprofile -c "Add-Type -AssemblyName System.Windows.Forms; $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog; $openFileDialog.Filter = 'PowerShell Script Files (*.ps1;*.psm1)|*.ps1;*.psm1|All Files (*.*)|*.*'; $openFileDialog.Title = 'Select the PowerShell script file'; if ($openFileDialog.ShowDialog() -eq 'OK') { $openFileDialog.FileName }"') do set "psFile=%%I"

if "%psFile%" == "" (
    echo No file selected. Script execution aborted.
    pause
    exit /b
)

echo.
echo Checking if the PowerShell script exists...
echo.

if exist "%psFile%" (
    echo The PowerShell script "%psFile%" exists.
    echo.
    echo Generating a temporary PowerShell script to run with elevated privileges...
    echo.

    :: Generate a temporary PowerShell script
    echo Start-Process powershell.exe -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File "%psFile%"' > "%temp%\run_elevated_temp.ps1"

    echo Running the PowerShell script with elevated privileges...
    echo.

    :: Run the temporary PowerShell script
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%temp%\run_elevated_temp.ps1"

    echo.
    echo Cleaning up the temporary PowerShell script...
    echo.

    :: Clean up the temporary PowerShell script
    del "%temp%\run_elevated_temp.ps1"

    echo.
    echo Script execution complete.
    echo.
    pause
) else (
    echo.
    echo The specified file "%psFile%" does not exist.
    echo Please make sure you have selected the correct file and try again.
    echo.
    PAUSE
)
