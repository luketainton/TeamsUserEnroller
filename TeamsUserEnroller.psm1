Function Import-TeamsUsers {

<#
    .SYNOPSIS
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.

    .DESCRIPTION
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.

    .PARAMETER File
    The path to the CSV file that contains your users. Can either be an absolute path or relative path.

    .PARAMETER Create
    If specified, create a new Group first, then add the users from the CSV file.

    .PARAMETER Delimiter
    If specified, overrides the default CSV delimiter of ','.

    .PARAMETER Encoding
    If specified, manually sets the encoding of the CSV file.

    .EXAMPLE
    Import-TeamsUsers -File "users.csv"

    .EXAMPLE
    Import-TeamsUsers -Create -File "users.csv"
#>

    Param(
        [parameter(Mandatory=$true, position=1, ParameterSetName='Params', HelpMessage="Specify CSV file")]
        [string]$File,
        [parameter(Mandatory=$false, position=2, ParameterSetName='Params', HelpMessage="Create new Teams group")]
        [switch]$Create,
        [parameter(Mandatory=$false, position=3, ParameterSetName='Params', HelpMessage="Override default CSV delimiter")]
        [string]$Delimiter,
        [parameter(Mandatory=$false, position=4, ParameterSetName='Params', HelpMessage="Manually set CSV encoding")]
        [string]$Encoding
    )

    Begin {
        $ErrorActionPreference = 'Stop'

        ##### CHECK FOR NEW VERSION #####
        Try {
            # Get information from GitHub Releases
            $releases = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/luketainton/TeamsUserEnroller/releases";
            $rel = $releases[0];
            $latest_version = $rel.tag_name -replace 'v', '';
            $latest_version_changes = $rel.body;

            # Get currently installed version
            $current_version = (Get-Module TeamsUserEnroller | Select-Object Version).Version;

            # Compare versions and alert user if newer version available
            if ($current_version -lt $latest_version) {
                Write-Host -ForegroundColor Yellow "A new version of TeamsUserEnroller has been released!";
                Write-Host -ForegroundColor Yellow "Latest version: $latest_version";
                Write-Host -ForegroundColor Yellow "Installed version: $current_version";
                Write-Host -ForegroundColor Yellow "`n$latest_version_changes";
                $Consent = Read-Host -Prompt "`nWould you like to update now? [y/N]"
                If ($Consent -eq "y" -Or $Consent -eq "Y") {
                    Update-Module -Name TeamsUserEnroller -RequiredVersion "2.2.0";
                    $after_update_ver = (Get-Module TeamsUserEnroller | Select-Object Version).Version;
                    if ($after_update_ver -eq $latest_version) {
                        Write-Host -ForegroundColor Green "Update completed.";
                    } Else {
                        Write-Host -ForegroundColor Red "Update failed. Please update manually.";
                    }
                }
            }
        } Catch {
            Write-Host -ForegroundColor Red "An error occurred while checking for updates. Continuing.";
        }

        ##### IMPORT CSV FILE #####
        Try {
            $ImportCmd = "Import-CSV $File"
            If ($Delimiter) { $ImportCmd = $ImportCmd + " -Delimiter $Delimiter" }
            If ($Encoding) { $ImportCmd = $ImportCmd + " -Encoding $Encoding" }
            $Users = Invoke-Expression $ImportCmd
        } Catch {
            Write-Host -ForegroundColor Red "$File is not a valid CSV file."
            Exit
        }
        

        ##### CHECK MODULE IS INSTALLED AND IMPORTED #####
        if (Get-Module -ListAvailable -Name MicrosoftTeams) {
            try {
                Import-Module -Name MicrosoftTeams
                $Email = (Connect-MicrosoftTeams -Verbose:$false).Account
            } Catch {
                Write-Host -ForegroundColor Red "There was an error during authentication."
                Write-Host "If you're not on Windows and use Multi-Factor Authentication, please manually pass the MFA check in your browser, then try again."
                Exit
            }
        } else {
            Write-Host -ForegroundColor Red "Module MicrosoftTeams doesn't exist. Please run 'Install-Module -Name MicrosoftTeams' and retry."
            Exit
        }
    }

    Process {
        If ($Create) {
            ##### CREATE NEW TEAM #####
            Clear-Host
            $NewTeamName = Read-Host -Prompt "Name of the new group"
            $NewTeamDesc = Read-Host -Prompt "Group description"
            $NewTeamPriv = Read-Host -Prompt "P[u]blic or P[r]ivate?"
            If ($NewTeamPriv -Eq "u") {
                $NewTeamVis = "Public"
            } Elseif ($NewTeamPriv -Eq "r") {
                $NewTeamVis = "Private"
            }
            $NewTeam = New-Team -DisplayName $NewTeamName -MailNickName $NewTeamName -Description $NewTeamDesc -Visibility $NewTeamVis
            $GroupId = $NewTeam.GroupId
        } Else {
            ##### GET USER'S TEAMS #####
            Clear-Host
            Write-Host -ForegroundColor Green "Getting your teams - please wait"
            $EligibleTeams = @()
            Get-Team -User $Email -Verbose:$false | ForEach-Object {
                $CTeamId = $_.GroupId
                $CTeamName = $_.DisplayName
                If (Get-TeamUser -GroupId $CTeamId | Select-Object -Property User,Role | Where-Object {$_.User -eq $Email} | Where-Object {$_.Role -eq "owner"}) {
                    $EligibleTeams += @{GroupId = $CTeamId; DisplayName = $CTeamName}
                }
                Clear-Variable -Name CTeamId
                Clear-Variable -Name CTeamName
            }
            Clear-Host
            Write-Host "Teams that you own:"
            $EligibleTeams | ForEach-Object {[PSCustomObject]$_} | Format-Table 'GroupId', 'DisplayName' -AutoSize
            $GroupId = Read-Host -Prompt "GroupId of the desired group"
        }

        ##### ENROL USERS #####
        $global:UsersAdded = 0;
        $UserCount = $Users | Measure-Object | Select-Object -expand count
        $Consent = Read-Host -Prompt "You are about to add $UserCount users. Are you sure? [y/N]"
        If ($Consent -eq "y" -Or $Consent -eq "Y") {
            $Users | ForEach-Object {
                $User = $_.email
                $Role = $_.role
                Try {
                    Add-TeamUser -GroupId $GroupId -Role $Role -User $User
                    Write-Host "Added user $User with role $Role"
                    $global:UsersAdded++
                } Catch [Microsoft.TeamsCmdlets.PowerShell.Custom.ErrorHandling.ApiException] {
                    Write-Host -ForegroundColor Red "Error adding user $User with role $Role"
                }
                Clear-Variable -Name User
                Clear-Variable -Name Role
            }
            Write-Host -ForegroundColor Green "$global:UsersAdded users added successfully."
        } Else {
            Write-Host -ForegroundColor Red "Aborting."
        }
    }
    
    End {
        @('UserCount', 'UsersAdded', 'Consent', 'Users', 'GroupId') | ForEach-Object {Clear-Variable -Name $_}
        Disconnect-MicrosoftTeams
    }
}
