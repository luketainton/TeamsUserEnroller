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

    .EXAMPLE
    Import-TeamsUsers -File "users.csv"

    .EXAMPLE
    Import-TeamsUsers -Create -File "users.csv"
#>

    Param(
        [parameter(Mandatory=$true, position=1, ParameterSetName='Params', HelpMessage="Specify CSV file")]
        [string]$File,
        [parameter(Mandatory=$false, position=2, ParameterSetName='Params', HelpMessage="Create new Teams group")]
        [switch]$Create
    )

    Begin {
        $ErrorActionPreference = 'Stop'
        ##### IMPORT CSV FILE #####
        Try {
            $Users = Import-CSV $File
        } Catch {
            Write-Host -ForegroundColor Red "$File is not a valid CSV file."
        }
        

        ##### CHECK MODULE IS INSTALLED AND IMPORTED #####
        if (Get-Module -ListAvailable -Name MicrosoftTeams) {
            Import-Module -Name MicrosoftTeams
            $Email = (Connect-MicrosoftTeams -Verbose:$false).Account
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
