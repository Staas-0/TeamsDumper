[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

$defaultProfilePicture = ("data:image/jpeg;base64,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")

$attempted = @{}

function Get-ProfilePicture ($userId, $assetsFolderPath, $clientId, $tenantId) {
    $profilePictureFile = Join-Path -Path "$assetsFolderPath" -ChildPath "$userId.jpg"

    if (Test-Path $profilePictureFile) {
        # if available
        Write-Verbose "Profile picture cache hit."
        "$../../../../assets/$userId.jpg"
    }
    elseif (($null -eq $userId) -or ($attempted.ContainsKey($userId))) {
        Write-Verbose "Profile picture unavailable, using default."

        # if userId is null or failed to download profile picture
        $defaultProfilePicture
    }
    else {
        # if never attempted
        
        Write-Verbose "Profile picture cache miss, fetching."

        $attempted.Add($userId, $null)
        $profilePhotoUri = "https://graph.microsoft.com/v1.0/users/" + $userId + "/photo/`$value"

        try {
            $start = Get-Date

            Invoke-Retry -Code {
                Invoke-WebRequest -Uri $profilePhotoUri -Headers @{
                    "Authorization" = "Bearer $(Get-GraphAccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId)"
                } -OutFile $profilePictureFile
            }

            Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to download profile picture."

            "../../../../assets/$userId.jpg"
        }
        catch {
            Write-Verbose "Failed to fetch profile picture, using default."
            $defaultProfilePicture
        }
    }
}