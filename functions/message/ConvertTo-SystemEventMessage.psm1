[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }

function ConvertTo-SystemEventMessage ($eventDetail, $clientId, $tenantId) {
    
    Write-Verbose "Processing event with type: $($eventDetail."@odata.type")"

    # https://learn.microsoft.com/en-us/graph/system-messages#supported-system-message-events
    switch ($eventDetail."@odata.type") {
        "#microsoft.graph.callEndedEventMessageDetail" {
            Write-Verbose "Call ended event detected, duration: $($eventDetail.callDuration)"
            return "Call ended after $($eventDetail.callDuration)."
        }
        "#microsoft.graph.callStartedEventMessageDetail" {
            Write-Verbose "Call started event detected."
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) started a call."
        }
        "#microsoft.graph.chatRenamedEventMessageDetail" {
            Write-Verbose "Chat renamed event detected. New chat name: $($eventDetail.chatDisplayName)"
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) changed the chat name to $($eventDetail.chatDisplayName)."
        }
        "#microsoft.graph.membersAddedEventMessageDetail" {
            Write-Verbose "Members added event detected. Processing added members."
            $addedMembers = $eventDetail.members | ForEach-Object {
                # Check the cache for display name, fall back if not found
                if ($Global:UserCache.ContainsKey($_.id)) {
                    Write-Verbose "Found display name in cache for member ID: $($_.id)"
                    $Global:UserCache[$_.id].displayName
                } else {
                    Write-Verbose "Display name not found in cache for member ID: $($_.id), fetching from API."
                    Get-DisplayName $_.id $clientId $tenantId
                }
            }
            Write-Verbose "Added members: $($addedMembers -join ', ')"
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) added $(($addedMembers -join ", "))."
        }
        "#microsoft.graph.membersDeletedEventMessageDetail" {
            Write-Verbose "Members deleted event detected."
            if (
                ($eventDetail.members.count -eq 1) -and
                ($null -ne $eventDetail.initiator.user) -and
                ($eventDetail.initiator.user.id -eq $eventDetail.members[0].id)
            ) {
                Write-Verbose "Single member left: $($eventDetail.members[0].id)"
                return "$(Get-DisplayName $eventDetail.members[0].id $clientId $tenantId) left."
            } else {
                $removedMembers = $eventDetail.members | ForEach-Object {
                    # Check the cache for display name, fall back if not found
                    if ($Global:UserCache.ContainsKey($_.id)) {
                        Write-Verbose "Found display name in cache for removed member ID: $($_.id)"
                        $Global:UserCache[$_.id].displayName
                    } else {
                        Write-Verbose "Display name not found in cache for removed member ID: $($_.id), fetching from API."
                        Get-DisplayName $_.id $clientId $tenantId
                    }
                }
                Write-Verbose "Removed members: $($removedMembers -join ', ')"
                return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) removed $(($removedMembers -join ", "))."
            }
        }
        "#microsoft.graph.channelAddedEventMessageDetail" {
            Write-Verbose "Channel added event detected. Channel: $($eventDetail.channelDisplayName)"
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) added $($eventDetail.channelDisplayName)"
        }
        "#microsoft.graph.messagePinnedEventMessageDetail" {
            Write-Verbose "Message pinned event detected."
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) pinned a message."
        }
        "#microsoft.graph.messageUnpinnedEventMessageDetail" {
            Write-Verbose "Message unpinned event detected."
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) unpinned a message."
        }
        "#microsoft.graph.teamsAppInstalledEventMessageDetail" {
            Write-Verbose "Teams app installed event detected. App: $($eventDetail.teamsAppDisplayName)"
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) added $($eventDetail.teamsAppDisplayName) here."
        }
        "#microsoft.graph.teamsAppRemovedEventMessageDetail" {
            Write-Verbose "Teams app removed event detected. App: $($eventDetail.teamsAppDisplayName)"
            return "$(Get-Initiator $eventDetail.initiator $clientId $tenantId) removed $($eventDetail.teamsAppDisplayName)."
        }
        Default {
            Write-Warning "Unhandled system event type: $($eventDetail."@odata.type")"
            Write-Verbose "Unhandled event details: $($eventDetail | ConvertTo-Json -Depth 5)"
            return "Unhandled system event type $($eventDetail."@odata.type"): $($eventDetail | ConvertTo-Json -Depth 5)"
        }
    }
}
