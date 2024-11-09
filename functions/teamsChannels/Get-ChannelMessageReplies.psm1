function Get-ChannelMessageReplies {
    [cmdletbinding()]
    param (
        [string]$teamId,
        [string]$channelId,
        [string]$messageId,
        [string]$clientId,
        [string]$tenantId
    )

    $start = Get-Date
    
    # Construct the URI for getting replies
    $repliesUri = "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$channelId/messages/$messageId/replies"

    $replies = @()

    try {
        while ($null -ne $repliesUri) {
            
            $repliesResponse = Invoke-Retry -Code {
                Invoke-RestMethod -Method Get -Uri $repliesUri -Headers @{
                "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
                "Prefer"        = "include-unknown-enum-members"
            }
        }

            # Add the fetched replies to the array
            $replies += $repliesResponse.value
            
            # Check for the next page of replies
            $repliesUri = $repliesResponse."@odata.nextLink"

            # Rate limiting - wait for 1 second before the next request
            Start-Sleep -Milliseconds 500
        }
    }
    catch {
        Write-Verbose "Failed to fetch replies. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($replies.count) replies."

    return $replies | Sort-Object createdDateTime
}
