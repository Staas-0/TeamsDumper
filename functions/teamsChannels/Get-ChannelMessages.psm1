function Get-ChannelMessages {
    [cmdletbinding()]
    param (
        [string]$teamId,
        [array]$channel,
        [string]$clientId,
        [string]$tenantId
    )

    $start = Get-Date
    
    $channelMessagesUri = "https://graph.microsoft.com/v1.0/teams/$teamId/channels/" + $channel.id + "/messages/"

    $channelMessages = @()

    try {
        while ($null -ne $channelMessagesUri) {
            $channelMessagesToAdd = Invoke-Retry -Code {
                Invoke-RestMethod -Method Get -Uri $channelMessagesUri -Headers @{
                    "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
                    "Prefer"        = "include-unknown-enum-members"
                }
            }

            $channelMessages += $channelMessagesToAdd.value
            $channelMessagesUri = $channelMessagesToAdd."@odata.nextLink"

            # Rate limiting - wait for 1 second before the next request
            Start-Sleep -Milliseconds 500
        }
    }
    catch {
        Write-Verbose "Failed to fetch messages. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($channelMessages.count) messages."

    $channelMessages | Sort-Object createdDateTime
}