function ConvertTo-ChatName {
    [CmdletBinding()]
    param (
        [PSCustomObject]$chat,
        [array]$members,
        [string]$clientId,
        [string]$tenantId
    )
    
    # Initialize $name as empty
    $name = ""

    # Set $name to $chat.topic if it exists
    if ($chat.topic) {
        $name = $chat.topic
    } 
    else {
        # If $chat.topic does not exist, derive name from member display names
        $memberNames = $members | ForEach-Object -Process {
            # Check if the display name is in the cache
            if ($null -eq $_.displayName) {
                # Use the user ID to fetch from the cache or get the display name if not found
                if ($Global:UserCache.ContainsKey($_.userId)) {
                    $_.displayName = $Global:UserCache[$_.userId].displayName
                } else {
                    $_.displayName = Get-DisplayName $_.userId $clientId $tenantId
                }
            }
            $_.displayName -replace '\s+', ''
        } | Select-Object -Unique | Sort-Object 

        # Set $name to the concatenated member names, excluding the chat.displayName if present
        $name = ($memberNames | Where-Object { $_ -ne $chat.displayName }) -join ", "
    }

    return $name
}