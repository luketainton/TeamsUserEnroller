# Import-TeamsUsers
A Powershell script that imports users from a CSV into a Microsoft Teams team.

# Setting up your device
This script runs via PowerShell, which is pre-installed on Windows. If you're not on Windows, please download **PowerShell Core** [here](https://github.com/PowerShell/PowerShell/releases). Once you've got PowerShell:
1. Open PowerShell as an administrator.
1. Allow remote scripts to execute by running `Set-ExecutionPolicy RemoteSigned`. If you don't do this, the script won't run.
1. Install the Microsoft Teams module. To do this, run `Install-Module -Name MicrosoftTeams`. Accept any prompts that you are given.
1. Install this module by running `Install-Module -Name Import-TeamsUsers`. Accept any prompts that you are given. 

# Running the script
1. Create a CSV file in the format `email,role`. The first line must be the headers `email,role`. You can copy the template if required.
1. Open PowerShell and run `Import-TeamsUsers -File <FILE>`, where `<FILE>` is the full path to the CSV file.

# Need help?
If you require assistance running the script, see the help by executing `Get-Help Import-TeamsUsers` (requires importing the module first - see step 4 in _Setting up your device_). If you still need help, please [send me an email](mailto:luke@tainton.uk?subject=I%20need%20help%20running%20Import-TeamsUsers).

# Issues? Want a new feature?
If you're having problems with the script or have an idea for a new feature, please check [here](https://github.com/luketainton/Import-TeamsUsers/issues) to see if someone else has the same problem or suggestion, and open an issue if one doesn't already exist. If you can implement a fix or feature request, please file a pull request!
