function Get-Chats ($clientId, $tenantId, $userId) {
    [CmdletBinding()]

    $link = "https://graph.microsoft.com/v1.0/users/$userId/chats/"
    $chats = @()

    $start = Get-Date

    try {
        while ($null -ne $link) {
            $chatsToAdd = Invoke-Retry -Code {
                Invoke-RestMethod -Method Get -Uri $link -Headers @{
                    "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
                }
            }
            
            $chats += $chatsToAdd.value
            $link = $chatsToAdd."@odata.nextLink"
            # Rate limiting - wait for 1 second before the next request
            Start-Sleep -Milliseconds 500
        }
    }
    catch {
        Write-Verbose "Failed to fetch chats. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($chats.count) chats."

    $chats | Sort-Object createdDateTime -Descending
}