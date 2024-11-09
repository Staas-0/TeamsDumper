function Get-UserChats {
    [CmdletBinding()]
    param (
        [string]$exportFolder,
        [string]$clientId,
        [string]$clientSecret,
        [string]$tenantId,
        [array]$selectedUsers
    )

    ####################################
    ##   HTML  ##
    ####################################

    Write-Verbose "Loading HTML templates and CSS..."
    $chatHTMLTemplate = Get-Content -Raw ./assets/chat.html
    $messageHTMLTemplate = Get-Content -Raw ./assets/message.html
    $reactionHTMLTemplate = Get-Content -Raw ./assets/reaction.html
    $stylesheetCSS = Get-Content -Raw ./assets/stylesheet.css

    # Script
    $start = Get-Date
    Write-Verbose "Script start time recorded."

    $chatsFolder = Join-Path -Path $exportFolder -ChildPath "chats"
    if (-not(Test-Path -Path $chatsFolder)) {
        Write-Verbose "Creating chats folder at $chatsFolder."
        New-Item -ItemType Directory -Path $chatsFolder | Out-Null
    }

    # Define the shared assets folder
    $sharedAssetsFolder = Join-Path -Path $exportFolder -ChildPath "assets"
    if (-not(Test-Path -Path $sharedAssetsFolder)) { 
        New-Item -ItemType Directory -Path $sharedAssetsFolder | Out-Null 
        Write-Verbose "Created shared assets folder: $sharedAssetsFolder"
    }
    
    # Loop through each user ID in selectedUsers
    foreach ($userId in $selectedUsers) {
        Write-Host "Processing chats for user: $userId"
        Write-Verbose "Creating subfolder for user $userId."

        # Create a subfolder for the user
        $userFolderName = $userId -replace '[^\w\s-]', '_'  # Sanitize user ID for folder name
        $userFolder = Join-Path -Path $chatsFolder -ChildPath $userFolderName
        if (-not(Test-Path -Path $userFolder)) { New-Item -ItemType Directory -Path $userFolder | Out-Null }

        Write-Verbose "Retrieving chats for user $userId."
        $chats = Get-Chats -clientId $clientId -tenantId $tenantId -userId $userId
        Write-Host ("" + $chats.count + " possible chats found for user $userId.")
        
        $chatIndex = 0

        foreach ($chat in $chats) {
            Write-Progress -Activity "Exporting Chats" -Status "User: $userId - Chat $($chatIndex) of $($chats.count)" -PercentComplete $(($chatIndex / $chats.count) * 100)
            $chatIndex += 1

            Write-Verbose "Retrieving members for chat $chatIndex."
            $members = Get-Members $chat $clientId $tenantId
            $name = ConvertTo-ChatName -chat $chat -members $members -clientId $clientId -tenantId $tenantId
            Write-Host $name

            Write-Verbose "Retrieving messages for chat."
            $messages = Get-Messages $chat $clientId $tenantId

            $messagesHTML = $null

            if (($messages.count -gt 0) -and (-not([string]::isNullOrEmpty($name)))) {

                Write-Host -ForegroundColor White ("`r`n$name :: $($messages.count) messages.")

                Write-Host "Processing messages..."

                foreach ($message in $messages) {
                    Write-Verbose "Getting profile picture for message sender."
                    $profilePicture = Get-ProfilePicture $message.from.user.id $sharedAssetsFolder $clientId $tenantId
                    Write-Verbose "Profile picture retrieved."
                    $time = ((Get-Date $message.createdDateTime -Format "dd MMMM yyyy, hh:mm tt") +" UTC")

                    switch ($message.messageType) {
                        "message" {
                            Write-Verbose "Processing regular message type."
                            $messageBody = $message.body.content

                            # Replace custom emoji with their alt text
                            Write-Verbose "Converting emojis in message body."
                            $messageBody = Convert-Emoji $messageBody

                            # Process embedded images
                            Write-Verbose "Finding and downloading embedded images in message."
                            $imageTagMatches = [Regex]::Matches($messageBody, "<img.+?src=[\`"']https:\/\/graph.microsoft.com(.+?)[\`"'].*?>")
                            foreach ($imageTagMatch in $imageTagMatches) {
                                Write-Verbose "Downloading embedded image."
                                $imagePath = Get-Image $imageTagMatch $sharedAssetsFolder $clientId $tenantId
                                $messageBody = $messageBody.Replace($imageTagMatch.Groups[0], "<img src=`"$imagePath`" style=`"width: 100%;`" >")
                            }

                            Write-Verbose "Building message HTML."
                            $messageHTML = $messageHTMLTemplate
                            $messageHTML = $messageHTML.Replace("###ATTACHMENTS###", (ConvertTo-HTMLAttachments $message.attachments))
                            $messageHTML = $messageHTML.Replace("###CONVERSATION###", $messageBody)
                            $messageHTML = $messageHTML.Replace("###DATE###", $time)
                            $messageHTML = $messageHTML.Replace("###DELETED###", "$($null -ne $message.deletedDateTime)".ToLower())
                            $messageHTML = $messageHTML.Replace("###EDITED###", "$($null -ne $message.lastEditedDateTime)".ToLower())
                            $messageHTML = $messageHTML.Replace("###IMAGE###", $profilePicture)
                            $messageHTML = $messageHTML.Replace("###NAME###", (Get-Initiator $message.from clientId $tenantId))
                            $messageHTML = $messageHTML.Replace("###PRIORITY###", $message.importance)

                            Write-Verbose "Processing reactions."
                            $reactions = $message.reactions
                            $reactionHTML = ""  # Initialize for this message

                            foreach ($reaction in $reactions) {
                                Write-Verbose "Building reaction HTML."
                                $singleReactionHTML = $reactionHTMLTemplate
                                $reactionUser = $reaction.user
                                $reactionType = $reaction.reactionType
                                $reactionTime = ConvertTo-CleanDateTime $reaction.createdDateTime
                                
                                $singleReactionHTML = $singleReactionHTML.Replace("###REACTION_NAME###", ((Get-Initiator $reactionUser clientId $tenantId) + " reacted with " + "$reactionType"))
                                $singleReactionHTML = $singleReactionHTML.Replace("###REACTION_DATE###", $reactionTime)
                                $reactionHTML += $singleReactionHTML
                            }
                            $messageHTML = $messageHTML.Replace("###REACTION###", $reactionHTML)
                            $messageHTML = $messageHTML.Replace("###REPLY###", "")
                            $messagesHTML += $messageHTML
                                
                            Break
                        }
                        "systemEventMessage" {
                            $systemImage = "data:image/webp;base64,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"
                            Write-Verbose "Processing system event message type."
                            $messageHTML = $messageHTMLTemplate
                            $messageHTML = $messageHTML.Replace("###ATTACHMENTS###", $null)
                            $messageHTML = $messageHTML.Replace("###CONVERSATION###", (ConvertTo-SystemEventMessage $message.eventDetail $clientId $tenantId))
                            $messageHTML = $messageHTML.Replace("###DATE###", $time)
                            $messageHTML = $messageHTML.Replace("###DELETED###", $null)
                            $messageHTML = $messageHTML.Replace("###EDITED###", $null)
                            $messageHTML = $messageHTML.Replace("###IMAGE###", $systemImage)
                            $messageHTML = $messageHTML.Replace("###NAME###", "System Event")
                            $messageHTML = $messageHTML.Replace("###PRIORITY###", $message.importance)
                            $messageHTML = $messageHTML.Replace("###REACTION###", "")
                            $messageHTML = $messageHTML.Replace("###REPLY###", "")

                            $messagesHTML += $messageHTML
                            Break
                        }
                        Default {
                            Write-Warning "Unhandled message type: $($message.messageType)"
                        }
                    }
                }

                $chatHTML = $chatHTMLTemplate
                $chatHTML = $chatHTML.Replace("###MESSAGES###", $messagesHTML)
                $chatHTML = $chatHTML.Replace("###CHATNAME###", $name)
                $chatHTML = $chatHTML.Replace("###STYLE###", $stylesheetCSS)
                
                $name = $name.Split([IO.Path]::GetInvalidFileNameChars()) -join "_"
                if ($name.length -gt 64) { $name = $name.Substring(0, 64) }
                $file = Join-Path -Path $userFolder -ChildPath "$name.html"
                
                if ($chat.chatType -ne "oneOnOne") {
                    Write-Verbose "Adding hash to filename for non-one-on-one chat."
                    # Add truncated SHA256 of the chat ID in case there are duplicate chat names
                    $chatIdStream = [IO.MemoryStream]::new([byte[]][char[]]$chat.id)
                    $chatIdShortHash = (Get-FileHash -InputStream $chatIdStream -Algorithm SHA256).Hash.Substring(0,8)
                    $file = $file.Replace(".html", ( " ($chatIdShortHash).html"))
                }

                Write-Host -ForegroundColor Green "Exporting $file..."
                $chatHTML | Out-File -LiteralPath $file
            }
            else {
                Write-Host ("`r`n$name :: No messages found.")
                Write-Host -ForegroundColor Yellow "Skipping..."
            }

            Start-Sleep -Milliseconds 500
        }
    }
    
    Write-Host "Finished processing chats in $((Get-Date) - $start)."
}
Export-ModuleMember -Function Get-UserChats
