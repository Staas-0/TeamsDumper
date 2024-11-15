$fileAttachmentHTMLTemplate = Get-Content -Raw ./assets/fileAttachment.html

function ConvertTo-HTMLAttachments ($attachments) {
    [CmdletBinding()]
    $attachmentsHTML = ""

    # files
    $fileAttachments = $attachments | Where-Object { $_.contentType -eq "reference" }
        
    foreach ($attachment in $fileAttachments) {
        $attachmentsHTML += $fileAttachmentHTMLTemplate.Replace("###ATTACHMENTURL###", $attachment.contentURL).Replace("###ATTACHMENTNAME###", $attachment.name)
    }
    
    if ($attachmentsHTML.Length -ge 0) {
        $attachmentsHTML
    } else {
        $null
    }
}