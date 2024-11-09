function ConvertTo-CleanDateTime ($iso) {
    [CmdletBinding()]
    $time = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date ($iso)), (Get-TimeZone).Id)
    Get-Date $time -Format "dd MMMM yyyy, hh:mm tt"
}