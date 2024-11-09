function Get-Members ($chat, $clientId, $tenantId) {
    [CmdletBinding()]
    $start = Get-Date

    $membersUri = "https://graph.microsoft.com/v1.0/" + $userId + "/chats/" + $chat.id + "/members"

    try {
        $members = Invoke-Retry -Code {
            Invoke-RestMethod -Method Get -Uri $membersUri -Headers @{
                "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
            }
        }
    }
    catch {
        Write-Verbose "Failed to fetch members. Failing."
        throw $_
    }

    Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to get $($members.value.count) members."

    $members.value
}
