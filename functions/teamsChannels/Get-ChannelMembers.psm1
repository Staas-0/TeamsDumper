function Get-ChannelMembers {
    [CmdletBinding()]
    param(
        [string]$teamId,
        [array]$channel,
        [string]$clientId,
        [string]$tenantId
    )

    $start = Get-Date

    $channelMembersUri = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + $channel.id + "/members"

    try {
        $channelMembers = Invoke-Retry -Code {
            Invoke-RestMethod -Method Get -Uri $channelMembersUri -Headers @{
                "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
            }
        }
    }
    catch {
        Write-Verbose "Failed to fetch channel members. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($channelMembers.value.count) members."

    $channelMembers.value
}
