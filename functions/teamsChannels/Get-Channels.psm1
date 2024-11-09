function Get-channels {
    [CmdletBinding()]
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$teamId  # New parameter for the user ID
    )
    $link = "https://graph.microsoft.com/v1.0/teams/$teamId/channels/"
    $channels = @()

    $start = Get-Date

    try {
        while ($null -ne $link) {
            $channelsToAdd = Invoke-Retry -Code {
                Invoke-RestMethod -Method Get -Uri $link -Headers @{
                    "Authorization" = "Bearer $(Get-GraphAccessToken $clientId $tenantId)"
                }
            }
            
            $channels += $channelsToAdd.value
            $link = $channelsToAdd."@odata.nextLink"

            # Rate limiting - wait for 1 second before the next request
            Start-Sleep -Milliseconds 500
        }
    }
    catch {
        Write-Verbose "Failed to fetch Channels. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($channels.count) channels."

    $channels | Sort-Object createdDateTime -Descending
}