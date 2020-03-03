function Import-TeamsUsers {
    Param(
        [parameter(Mandatory=$true,HelpMessage="Specify Group ID")]
        $GroupId,
        [parameter(Mandatory=$true,HelpMessage="Specify CSV file")]
        $File
    )

    # Import CSV file and required module
    $Users = Import-CSV $File
    $UserCount = $Users | Measure-Object | Select-Object -expand count
    Import-Module -Name MicrosoftTeams

    # Check Team exists
    Try {
        $Team = Get-Team -GroupId $GroupId
        If ($Team) {
            $TeamName = $Team.DisplayName
            Write-Host -ForegroundColor Green "Team $TeamName exists!"
        }
    } Catch [System.UnauthorizedAccessException] {
        # User is not authenticated
        Write-Host -ForegroundColor Red "You need to authenticate to Microsoft Teams before continuing. Please run 'Connect-MicrosoftTeams' and try again."
        Break
    } Catch [System.Net.Http.HttpRequestException] {
        # Team does not exist
        Write-Host -ForegroundColor Red "Team with ID $GroupId does not exist!"
        Break
    }

    $Consent = Read-Host -Prompt "You are about to add $UserCount users. Are you sure? [y/N]"
    If ($Consent -eq "y" -Or $Consent -eq "Y") {
        $Users | ForEach-Object {
            $User = $_.email
            $Role = $_.role
            Write-Host "Adding user $User with role $Role"
            Add-TeamUser -GroupId $GroupId -Role $Role -User $User -ErrorAction SilentlyContinue
        }
    } Else {
        Write-Host -ForegroundColor Red "Aborting."
    }

}


$GroupId = ""
$File = ""

If ($GroupId -And $File) {
	Import-TeamsUsers -GroupId $GroupId -File $File
} Else {
	Write-Host -ForegroundColor Red "`$GroupId and/or `$File missing."
}
