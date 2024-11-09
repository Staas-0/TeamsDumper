[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }

# used to get the initator of an event

function Get-Initiator ($identitySet) {
    if ($identitySet.user) {
        $userId = $identitySet.user.id

        # Check the cache for the display name
        if ($Global:UserCache.ContainsKey($userId)) {
            return $Global:UserCache[$userId].displayName
        }
        else {
            # If not found in cache, use Get-DisplayName
            return Get-DisplayName $userId
        }
    }
    elseif ($identitySet.application) {
        if ($null -ne $identitySet.application.displayName) {
            return $identitySet.application.displayName
        }
        else {
            return "An application"
        }
    }
    else {
        return "Unknown"
    }
}
