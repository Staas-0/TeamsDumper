<# 
    .SYNOPSIS
        This script exports Microsoft Teams data to HTML.

    .DESCRIPTION
        This script retrieves Microsoft Teams chat and message data and exports it to HTML files.

    .PARAMETER exportFolder
        Export location of where the HTML files will be saved.

    .PARAMETER clientId
        Enter Application (client) ID.

    .PARAMETER clientSecret
        Enter the Client Secret value.

    .PARAMETER tenantId
        Enter the Directory (tenant) ID.

    .PARAMETER startDate
        Enter start date in the format: YYYY-MM-DD.

    .PARAMETER endDate
        Enter end date in the format: YYYY-MM-DD.

    .EXAMPLE
        .\TeamsDumper.ps1 -exportFolder "C:\Exports" -clientId "YourClientId" -clientSecret "YourClientSecret" -tenantId "YourTenantId" -startDate "2024-01-01" -endDate "2024-12-31"
        Exports Teams data from the specified date range.

    .NOTES
        Pre-requisites: An app registration with the following permissions:
        User Chats: Chat.Read.All, ChatMember.Read.All, ChatMessage.Read.All, User.Read.All
        Teams Channels: Team.ReadBasic.All, TeamMember.Read.All, Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All, Files.Read.All, OnlineMeetingRecording.Read.All, OnlineMeetingTranscript.Read.All
#>

[cmdletbinding()]
Param(
    [Parameter(Mandatory = $false, HelpMessage = "Export location of where the HTML files will be saved.")][string] $exportFolder = "$PSScriptRoot\Exports",
    [Parameter(Mandatory = $false, HelpMessage = "Enter Application (client) ID: ")] [string] $clientId,
    [Parameter(Mandatory = $false, HelpMessage = "Enter the Client Secret value: ")] [string] $clientSecret,
    [Parameter(Mandatory = $false, HelpMessage = "Enter the Directory (tenant) ID: ")] [string] $tenantId,
    [Parameter(Mandatory = $false, HelpMessage = "Enter start date: ")] [string] $startDate,
    [Parameter(Mandatory = $false, HelpMessage = "Enter end date: ")] [string] $endDate
)

$verbose = $PSBoundParameters["Verbose"]
Set-Location $PSScriptRoot


    #################################
    ##   Import Modules  ##
    #################################

Get-ChildItem "$PSScriptRoot/functions/main/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/util/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/userChats/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/teamsChannels/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/message/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }



$start = Get-Date
Write-Host $exportFolder
Write-Host -ForegroundColor Cyan "Starting script..."
Write-Host

# Get app credentials
Write-Host -ForegroundColor Cyan "Getting App Registration Credentials..."
# Check if tenantId is already set, otherwise prompt for input
if (-not $clientId) {
    $clientId = Read-Host -Prompt "Enter Client ID"
}

# Check if clientId is already set, otherwise prompt for input
if (-not $clientSecret) {
    $clientSecret = Read-Host -Prompt "Enter Client Secret"
}

# Check if clientSecret is already set, otherwise prompt for input
if (-not $tenantId) {
    $tenantId = Read-Host -Prompt "Enter Tenant ID"
}


# Display a prompt with options for the user
Write-Host -ForegroundColor Cyan "Select an option to export: "
Write-Host "1) Export one or more user chats"
Write-Host "2) Export channels from one or more Teams"
Write-Host "3) Export all user chats and Teams channels"
Write-Host

# Capture the user's selection
$selectFunction = Read-Host -Prompt "Option"

# Process the selection
switch ($selectFunction) {
    '1' {
        Write-Host -ForegroundColor Green "You selected to export user chats."
        $selectedUsers = Get-UserOptions -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        Get-UserChats -exportFolder $exportFolder -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -selectedUsers $selectedUsers
    }
    '2' {
        Write-Host -ForegroundColor Green "You selected to export team channels."
        $selectedTeams = Get-TeamOptions -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        Get-TeamsChannels -exportFolder $exportFolder -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -selectedTeams $selectedTeams
    }
    '3' {
        Write-Host -ForegroundColor Green "You selected to export both user chats and channels."
        $selectedUsers = Get-UserOptions -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        $selectedTeams = Get-TeamOptions -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
        Get-UserChats -exportFolder $exportFolder -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -selectedUsers $selectedUsers
        Get-TeamsChannels -exportFolder $exportFolder -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -selectedTeams $selectedTeams
    }
    default {
        Write-Host -ForegroundColor Red "Invalid selection. Please enter 1, 2, or 3."
    }
}
Write-Progress -Completed
Write-Host -ForegroundColor Cyan "`r`nScript completed after $(((Get-Date) - $start).TotalSeconds)s... Bye!"