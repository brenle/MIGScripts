# Ensure you are connect to Exchange Online PowerShell
param (
    [Parameter(Mandatory = $true)][string]$targetUPN #enter UPN of user to check ELCLastSuccessTimestamp
)

try {
    Write-Host "Looking up $targetUPN..." -NoNewLine
    $targetMailbox = Get-Mailbox $targetUPN -ErrorAction Stop
    Write-Host -ForegroundColor Green "OK"
} catch {
    Write-Host -ForegroundColor Red "FAILED"
    Write-Host -ForegroundColor Red $error[0]
    exit
}

$diagLogs = Export-MailboxDiagnosticLogs $targetMailbox.primarysmtpaddress -ExtendedProperties
$xmlProperties = [xml]($diagLogs.MailboxLog)
$ELCLastSuccess = $xmlProperties.Properties.MailboxTable.property | ?{$_.Name -like "ELCLastSuccessTimestamp"}

if($ELCLastSuccess -eq $null){
    write-Host "No ELC timestamp found."
} else {
    $ELCLastSuccess
}
