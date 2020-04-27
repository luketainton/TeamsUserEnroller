# TeamsUserEnroller
A Powershell module that imports users from a CSV into a Microsoft Teams group.

# Setting up your device
This module uses PowerShell, which is pre-installed on Windows. If you're not on Windows, please [download **PowerShell Core**](https://github.com/PowerShell/PowerShell/releases).
1. Open PowerShell as an administrator.
1. Install this module by running `Install-Module -Name TeamsUserEnroller`. 

# Running the script
1. Create a CSV file containing your users and their desired roles. The first line must be the headers `email,role`, for example:
   ```csv
   email,role
   jbloggs@example.com,owner
   user@example.com,member
   ```
1. Run `Import-TeamsUsers -File <FILE>`, where `<FILE>` is the path to the CSV file.

<details>
  <summary>If you can't run non-signed scripts</summary>
  If your policy requires scripts to be digitally signed, run

  ```powershell
  Set-ExecutionPolicy Bypass -Scope Process
  ```
  then try running the command again. You may require administrative rights to change the Execution Policy.
</details>

# Need help?
If you need assistance, please try the following:
1. See the help documentation by running `Get-Help Import-TeamsUsers`.
1. Check closed issues [here](https://github.com/luketainton/Import-TeamsUsers/issues?q=is%3Aissue+sort%3Aupdated-desc+is%3Aclosed).
1. Open an issue [here](https://github.com/luketainton/Import-TeamsUsers/issues/new).
