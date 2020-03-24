# Import-TeamsUsers
A Powershell script that imports users from a CSV into a Microsoft Teams team.

# Setting up your device
This script runs via PowerShell. If you're on Windows, you'll already have this. If not, please download it from the [releases page](https://github.com/PowerShell/PowerShell/releases). Once you've got PowerShell:
1. Open PowerShell as an administrator.
2. Allow remote scripts to execute by running `Set-ExecutionPolicy RemoteSigned`. If you don't do this, the script won't run.
3. Install the Microsoft Teams module. To do this, run `Install-Module -Name MicrosoftTeams`. Accept any prompts that you are given.

# Running the script
1. Download the repository to your PC.
2. Create a CSV file in the format `email,role`. The first line must be the headers `email,role`. You can copy the template if required.
3. Open PowerShell and change directory (`cd`) to the directory where you downloaded the repository.
3. Import the module (`Import-Module ./Import-TeamsUsers.psm1`).
4. Run `Import-TeamsUsers -File <FILE>`.

# Need help?
If you require assistance running the script, see the help by executing `Get-Help Import-TeamsUsers` (requires importing the module first - see steps 3 and 4 above). If you still need help, please [send me an email](mailto:luke@tainton.uk?subject=I%20need%20help%20running%20Import-TeamsUsers).

# Issues? Want a new feature?
If you're having problems with the script or have an idea for a new feature, please check [here](https://github.com/luketainton/Import-TeamsUsers/issues) to see if someone else has the same problem or suggestion, and open an issue if one doesn't already exist. If you can implement a fix or feature request, please file a pull request!
