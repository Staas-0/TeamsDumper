[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

# Set verbose and progress preferences based on the input parameter
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

function Get-TeamOptions {
    param (
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )

    #Initialize UserCache
    Get-AllUsers -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret | Out-Null
    
    $selectedTeams = @()

    # Retrieve the list of teams
    $teamList = Get-AllTeams -clientId $clientId -tenantId $tenantId -clientSecret $clientSecret
    
    # Display the current list of teams
    Write-Host -ForegroundColor Cyan "Found the following teams:"
    $teamList | ForEach-Object { Write-Host $_.displayName }
    Write-Host  # Just adds a blank line for spacing

    # Prompt the user for input
    $teamInput = Read-Host -Prompt "Select a list of teams (comma-separated) or enter 'all' to export channels from all teams"
    $teamInput = $teamInput.ToLower()

    # Process the user input
    if ($teamInput -eq 'all') {
        $selectedTeams = @($teamList.id)  # Assign all Team IDs if 'all' is selected
        Write-Host -ForegroundColor Green "Exporting chats from all teams..."
    } else {
        # Split the input into a list and trim whitespace
        $inputTeams = $teamInput -split ',' | ForEach-Object { $_.Trim() }
        
        # Find matching teams in the full list
        foreach ($inputTeam in $inputTeams) {
            $matchedTeam = $teamList | Where-Object { $_.displayName -eq $inputTeam.ToLower() }
            if ($matchedTeam) {
                $selectedTeams += @($matchedTeam.id)  # Collect the IDs of matched teams
            } else {
                Write-Host -ForegroundColor Yellow "Team '$inputTeam' not found."
            }
        }
        Write-Host -ForegroundColor Green "Exporting channels from the following Teams: $($inputTeams -join ', ')"
    }
    $selectedTeams
}

