[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

function Get-User {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$userId
    )

    # Check the global user cache for the user ID
    if ($Global:UserCache.ContainsKey($userId)) {
        Write-Verbose "User cache hit."
        return $Global:UserCache[$userId]
    }
    else {
        Write-Verbose "User cache miss, fetching."

        $start = Get-Date

        $userUri = "https://graph.microsoft.com/v1.0/users/" + $userId
        
        # Fetch the user from the API
        $user = Invoke-Retry -Code {
            Invoke-RestMethod -Method Get -Uri $userUri -Headers @{
                "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
            }
        }

        Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get user."

        # Add the fetched user to the global user cache
        $Global:UserCache[$userId] = $user

        return $user
    }
}
