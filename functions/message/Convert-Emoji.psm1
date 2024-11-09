function Convert-Emoji {
    param (
        [string]$htmlContent
    )

    # Use regex to find and replace <emoji> tags with their alt text
    $updatedContent = [Regex]::Replace($htmlContent, '<emoji id="[^"]+" alt="([^"]+)" title="[^"]+"></emoji>', '$1')
    return $updatedContent
}
