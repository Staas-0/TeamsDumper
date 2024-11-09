[CmdletBinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

# Variables for access token
$global:accessToken = $null
$global:expires = $null

function Get-GraphAccessToken {
    param (
        [string]$clientId,
        [string]$clientSecret,
        [string]$tenantId
    )
    # If token is still valid for at least 10 minutes, return it
    if ($global:expires -ge (Get-Date).AddMinutes(10)) {
        Write-Verbose "Returning cached access token."
        return $accessToken
    }

    Write-Verbose "Fetching new access token with client credentials grant."

    # Create body for client credentials flow
    $tokenBody = @{
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    # Request an access token
    try {
        $authRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody -ErrorAction Stop
        $global:accessToken = $authRequest.access_token
        $global:expires = (Get-Date).AddSeconds($authRequest.expires_in)
        Write-Verbose "Access token obtained successfully."
    }
    catch {
        # Write out the full exception details for easier troubleshooting
        Write-Verbose "Failed to obtain token: $($_ | Out-String)"
        throw "Error obtaining access token. Please verify your clientId, clientSecret, and tenantId."
    }
    

    return $global:accessToken
}