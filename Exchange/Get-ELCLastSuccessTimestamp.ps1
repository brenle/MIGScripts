# Ensure you are connect to Exchange Online PowerShell
param (
    [Parameter(Mandatory = $true)][string]$targetUPN #enter UPN of user to check ELCLastSuccessTimestamp
)

$mailboxFound = $true
try {
    $targetMailbox = Get-Mailbox $targetUPN -ErrorAction Stop
} catch {
    write-host -ForegroundColor Red "Mailbox not found."
    $mailboxFound = $false
}
    if($mailboxFound){
    $diagLogs = Export-MailboxDiagnosticLogs $targetMailbox.primarysmtpaddress -ExtendedProperties
    $xmlProperties = [xml]($diagLogs.MailboxLog)
    $ELCLastSuccess = $xmlProperties.Properties.MailboxTable.property | ?{$_.Name -like "ELCLastSuccessTimestamp"}

    if($ELCLastSuccess -eq $null){
        write-Host "No ELC timestamp found."
    } else {
        $ELCLastSuccess
    }
}
