# 📝 Registry-Changer.ps1

## 🌟 Introduction
The `Registry-Changer.ps1` is a powerful 💪 PowerShell script designed to load settings from an XML file 📂 and apply them to the Windows registry for all user profiles on a machine. 🖥️

## 🚀 How to Use
1. Clone this repository or download the script to your local machine. 📥
2. Open a PowerShell terminal and navigate to the directory containing the script. 🗂️
3. Run the script with the command `.\Registry-Changer.ps1`. 🏃‍♂️

When the script is run, it will open a dialog box 🗃️ prompting you to select an XML file containing the settings. This XML file can be any file you choose that contains the appropriate parameters. 📝

If no XML file is chosen (dialog is canceled), the script will then prompt you to manually input new parameters. These parameters include:

- Registry path: The path to the registry key (not including the `HKEY_USERS\{SID}` part). 🔑
- Value name: The name of the registry value to set. 🏷️
- Value data: The data for the registry value. 📊

After the XML file is chosen, or after new parameters have been entered, the script will display the loaded settings and ask for your confirmation to continue. ✔️

If you choose to continue, the script will apply these parameters as registry settings for each user profile on the machine. Any errors encountered during this process will be displayed in the terminal. If you choose not to continue, the script will terminate. ❌

If you've entered new parameters manually, the script will save them to a new XML file for future use. The file will be named `Settings_{number}.xml`, where `{number}` is the next available number. 📋

## ⚠️ Caution
Always exercise caution when modifying the Windows registry, as incorrect changes can result in system instability or other unwanted behaviors. Always back up your registry or create a system restore point before making changes. 🛡️

Enjoy using `Registry-Changer.ps1`! 🎉 Remember, with great power comes great responsibility. Use wisely! 💼
