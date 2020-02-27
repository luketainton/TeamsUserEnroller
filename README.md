# Import-TeamsUsers
A Powershell script that imports users from a CSV into a Microsoft Teams team.

# Setup
You'll need to do a few things before you can run the script:
1. Have a CSV file. This needs to be in the format `email,role`. You can copy the template if required.
2. Install the module. Run `Install-Module -Name MicrosoftTeams` in your Powershell terminal.
3. Import the module. Run `Import-Module -Name MicrosoftTeams` in your Powershell terminal.
4. Authenticate to the API. Run `Connect-MicrosoftTeams` in your Powershell terminal.

# How to use
Open `Import-TeamsUsers.ps1` and modify the `$GroupId` and `$File` variables at the bottom, then execute the script from the command line.

# Issues?
If you're having problems or have an idea for a new feature, please check [here](https://github.com/luketainton/Import-TeamsUsers/issues) to see if someone else is having the same problem, and open an issue if one doesn't already exist. If you can implement a fix or feature request, please file a pull request!
