# Import-TeamsUsers
A Powershell script that imports users from a CSV into a Microsoft Teams team.

# Setting up your device
This script runs via PowerShell. If you're on Windows, you'll already have this. If not, please download it from the [releases page](https://github.com/PowerShell/PowerShell/releases). Once you've got PowerShell:
1. Open PowerShell.
2. Allow remote scripts to execute by running `Set-ExecutionPolicy RemoteSigned`. If you don't do this, the script won't run.
3. Install the Microsoft Teams module. To do this, run `Install-Module -Name MicrosoftTeams`. Accept any prompts that you are given.

# Gathering information
You'll need to do a few things before you can run the script:
1. Have a CSV file with the users you want to add. This needs to be in the format `email,role`. You can copy the template if required.
2. Import the Microsoft Teams module. Run `Import-Module -Name MicrosoftTeams` in your Powershell terminal.
3. Authenticate to Microsoft Teams. Run `Connect-MicrosoftTeams` in your Powershell terminal and follow the instructions.
4. Get your group ID. Run `Get-Team -User <EMAIL>`, substituting `<EMAIL>` for your Office 365 email address, to list all teams you are a member of.

# Running the script
1. Copy or move the CSV file to the same folder that the `Import-TeamsUsers.ps1` file is in.
2. Open the `Import-TeamsUsers.ps1` file.
3. Modify the `$GroupId` variable to the Group ID you found from step 4 in the section above.
4. Modify the `$File` variables to the name of your CSV file.
5. Run the script. 

# Need help?
If you require assistance running the script, please [send me an email](mailto:luke@tainton.uk?subject=I%20need%20help%20running%20Import-TeamsUsers).

# Issues? Want a new feature?
If you're having problems with the script or have an idea for a new feature, please check [here](https://github.com/luketainton/Import-TeamsUsers/issues) to see if someone else is having the same problem, and open an issue if one doesn't already exist. If you can implement a fix or feature request, please file a pull request!
