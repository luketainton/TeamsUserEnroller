Function Import-TeamsUsers {

<#
    .SYNOPSIS
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.

    .DESCRIPTION
    Import-TeamsUsers is a Powershell function that will enrol users from a CSV file into a given Microsoft Teams group.
    It has two required parameters (switches): -GroupId and -File. 

    .PARAMETER GroupId
    Specifies the unique identified for the destination Team. This can be retrieved from Powershell by executing the Get-Team cmdlet.

    .PARAMETER File
    The path to the CSV file that contains your users. Can either be an absolute path or relative path.

    .EXAMPLE
    Import-TeamsUsers -GroupId "00000000-0000-0000-0000-000000000000" -File "users.csv"
#>

    Param(
        [parameter(Mandatory=$true, position=0, ParameterSetName='Params', HelpMessage="Specify Group ID")]
        [string]$GroupId,
        [parameter(Mandatory=$true, position=1, ParameterSetName='Params', HelpMessage="Specify CSV file")]
        [string]$File
    )

    Begin {
        ##### IMPORT CSV FILE #####
        If ($File) {
            $Users = Import-CSV $File
        } Else {
            Write-Host -ForegroundColor Red "CSV file not specified or does not exist."
            Exit
        }

        ##### CHECK MODULE IS INSTALLED AND IMPORTED #####
        if (Get-Module -ListAvailable -Name MicrosoftTeams) {
            Import-Module -Name MicrosoftTeams
            Connect-MicrosoftTeams
        } else {
            Write-Host -ForegroundColor Red "Module MicrosoftTeams doesn't exist. Please run 'Install-Module -Name MicrosoftTeams' and retry."
            Exit
        }

        ##### CHECK TEAM EXISTS #####
        Try {
            $Team = Get-Team -GroupId $GroupId
            If ($Team) {
                $TeamName = $Team.DisplayName
                Write-Host -ForegroundColor Green "Team $TeamName exists!"
            }
        } Catch [System.UnauthorizedAccessException] {
            # User is not authenticated or does not have access
            Write-Host -ForegroundColor Red "You do not have access to manage this team."
            Exit
        } Catch [System.Net.Http.HttpRequestException] {
            # Team does not exist
            Write-Host -ForegroundColor Red "Team with ID $GroupId does not exist."
            Exit
        }

    }
    Process {
        $UsersAdded = 0;
        $UserCount = $Users | Measure-Object | Select-Object -expand count
        $Consent = Read-Host -Prompt "You are about to add $UserCount users. Are you sure? [y/N]"
        If ($Consent -eq "y" -Or $Consent -eq "Y") {
            $Users | ForEach-Object {
                $User = $_.email
                $Role = $_.role
                Write-Host "Adding user $User with role $Role"
                If (Add-TeamUser -GroupId $GroupId -Role $Role -User $User -ErrorAction SilentlyContinue) {$global:UsersAdded++}
            }
        } Else {
            Write-Host -ForegroundColor Red "Aborting."
            Exit
        }
    }
    
    End {
        Write-Host -ForegroundColor Green "$UsersAdded added successfully."
    }
}
