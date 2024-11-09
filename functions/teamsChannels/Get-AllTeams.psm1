[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"


function Get-AllTeams {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )

    Write-Verbose "Fetching all teams."

    $start = Get-Date

    $teamUri = "https://graph.microsoft.com/v1.0/teams"
    
    $allTeams = Invoke-Retry -Code {
        Invoke-RestMethod -Method Get -Uri $teamUri -Headers @{
            "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
        }
    }
    
    $allTeams.value

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get all users."
}