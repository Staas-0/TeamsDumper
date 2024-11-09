function Get-UserOptions {
    [CmdletBinding()]
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )

    $users = @{}  # Initialize the hashtable
    $selectedUsers = @()

    # Retrieve the list of users
    $userList = Get-AllUsers -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
    
    # Check if userList is valid
    if (-not $userList) {
        Write-Host -ForegroundColor Red "User list is null or empty."
        return  # Exit the function if there's no user data
    }

    # Display the current list of users
    Write-Host -ForegroundColor Cyan "Current list of users:"
    foreach ($user in $userList) {
        if (-not $user.PSObject.Properties.Match('id') -or -not $user.PSObject.Properties.Match('displayName')) {
            Write-Host -ForegroundColor Red "User object does not have required properties."
            continue  # Skip this user if properties are missing
        }

        if (-not $users.ContainsKey($user.id)) {
            $users[$user.id] = $user.displayName
        }
        Write-Host $user.displayName
    }
    Write-Host  # Just adds a blank line for spacing

    # Prompt the user for input
    $userInput = Read-Host -Prompt "Select a list of users (comma-separated) or enter 'all' to export chats from all users"
    $userInput = $userInput.ToLower()

    # Process the user input
    if ($userInput -eq 'all') {
        $selectedUsers = $userList.userPrincipalName  # Assign all user IDs if 'all' is selected
        Write-Host -ForegroundColor Green "Exporting chats from all users..."
    } else {
        # Split the input into a list and trim whitespace
        $inputUsers = $userInput -split ',' | ForEach-Object { $_.Trim() }
        
        # Find matching users in the full list
        foreach ($inputUser in $inputUsers) {
            # Use the user cache to find the display name
            $matchedUser = $userList | Where-Object { 
                $_.displayName.ToLower() -eq $inputUser.ToLower()
            }
            if ($matchedUser) {
                $selectedUsers += $matchedUser.userPrincipalName  # Collect the IDs of matched users
            } else {
                Write-Host -ForegroundColor Yellow "User '$inputUser' not found."
            }
        }
        Write-Host -ForegroundColor Green "Exporting chats from the following users: $($inputUsers -join ', ')"
    }
    $selectedUsers
}