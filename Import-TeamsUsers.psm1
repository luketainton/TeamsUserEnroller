Function Import-TeamsUsers {

<#
    .SYNOPSIS
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.

    .DESCRIPTION
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.
    It has two required parameters (switches): -Email and -File. 

    .PARAMETER Email
    Your Office 365 email address. This is used to list your teams.

    .PARAMETER File
    The path to the CSV file that contains your users. Can either be an absolute path or relative path.

    .EXAMPLE
    Import-TeamsUsers -Email "user@domain.com" -File "users.csv"
#>

    Param(
        [parameter(Mandatory=$true, position=0, ParameterSetName='Params', HelpMessage="Specify your Office 365 email address")]
        [string]$Email,
        [parameter(Mandatory=$true, position=1, ParameterSetName='Params', HelpMessage="Specify CSV file")]
        [string]$File
    )

    Begin {
        $ErrorActionPreference = 'Stop'
        ##### IMPORT CSV FILE #####
        $Users = Import-CSV $File

        ##### CHECK MODULE IS INSTALLED AND IMPORTED #####
        if (Get-Module -ListAvailable -Name MicrosoftTeams) {
            Import-Module -Name MicrosoftTeams
            Connect-MicrosoftTeams
        } else {
            Write-Host -ForegroundColor Red "Module MicrosoftTeams doesn't exist. Please run 'Install-Module -Name MicrosoftTeams' and retry."
            Exit
        }
    }

    Process {
        ##### GET USER'S TEAMS #####
        Get-Team -User $Email | Select-Object -Property GroupId, DisplayName | Format-Table -AutoSize
        $GroupId = Read-Host -Prompt "Paste the GroupId of the desired group"

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

            }
        } Else {
            Write-Host -ForegroundColor Red "Aborting."
            Exit
        }
    }
    
    End {
        Write-Host -ForegroundColor Green "$global:UsersAdded users added successfully."
        Disconnect-MicrosoftTeams
    }
}
