[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"


function Get-AllUsers {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )

    Write-Verbose "Fetching all users."

    $start = Get-Date

    $userUri = "https://graph.microsoft.com/v1.0/users"
    
    # Initialize the global user cache if it doesn't exist
    if (-not $Global:UserCache) {
        $Global:UserCache = @{}
    }

    try {
        $allUsers = Invoke-Retry -Code {
            Invoke-RestMethod -Method Get -Uri $userUri -Headers @{
                "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
            }
        }

        # Populate the global cache with user IDs and their corresponding objects
        foreach ($user in $allUsers.value) {
            $Global:UserCache[$user.id] = $user
        }

        Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get all users."

        return $allUsers.value  # Return the list of users if needed
    }
    catch {
        Write-Error "Failed to fetch users: $_"
    }
}