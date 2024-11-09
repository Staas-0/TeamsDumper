function Get-DisplayName {
    [CmdletBinding()]
    Param (
        [string]$userId
    )

    # Check the global cache for the user
    if ($Global:UserCache.ContainsKey($userId)) {
        return $Global:UserCache[$userId].displayName
    }

    Write-Verbose "User with ID $userId not found in cache."
    return "Unknown ($userId)"
}