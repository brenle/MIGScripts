param (
    [Parameter(Mandatory = $true)][string]$mailboxIdentity
)

function identifyPolicyOrHold ([string]$policy, [bool]$typeOnly)
{
    if($policy.substring(0,4) -eq "UniH")
    {
        #eDiscovery Hold
        if($typeOnly){
            return "eDiscovery"
        }
    } elseif ((($policy.substring(0,1) -eq "c") -or ($policy.substring(0,3) -eq "cld")) -and ($policy -ne "ComplianceTagHold")) {
        #inPlace Hold
        if($typeOnly){
            return "InPlace"
        }
    } elseif (($policy.substring(0,3) -eq "mbx") -or ($policy.substring(0,3) -eq "skp") -or ($policy.substring(0,4) -eq "-mbx") -or ($policy.substring(0,3) -eq "grp")) {
        #M365 retention policy
        if($typeOnly){
            return "Retention"
        } else {
            if(($policy.substring(0,3) -eq "mbx") -or ($policy.substring(0,3) -eq "skp") -or ($policy.substring(0,3) -eq "grp")){
                $policyGuid = $policy.trim($policy.Substring(0,3))
                $policyGuid = $policyGuid.trim($policyGuid.Substring($policyGuid.Length - 2))
                return $policyGuid
            } elseif ($policy.substring(0,4) -eq "-mbx"){
                $policyGuid = $policy.trim($policy.Substring(0,4))
                $policyGuid = $policyGuid.trim($policyGuid.Substring($policyGuid.Length - 2))
                return $policyGuid   
            } else {
                return
            }
        }
    } elseif ($policy -eq "LitigationHold"){
        #LitHold
        if($typeOnly){
            return "LitigationHold"
        }
    } elseif ($policy -eq "ComplianceTagHold"){
        #M365 label policy
        if($typeOnly){
            return "LabelHold"
        }
    } elseif ($policy -eq "DelayReleaseHold"){
        #Delay Release Hold
        if($typeOnly){
            return "DelayHoldApplied"
        }
    } else {
        #can't determine type
        return "UNKNOWN"
    }
}

function identifyRetentionPolicyAction ([string]$policy)
{
    $type = $policy.substring($policy.Length - 2)
    if ($type -eq ":1"){
        #DeleteOnly
        return "DeleteOnly"
    } elseif ($type -eq ":2"){
        #HoldOnly
        return "RetainOnly"
    } elseif ($type -eq ":3"){
        #Hold and Delete
        return "RetainThenDelete"
    } elseif ($policy.substring(0,1) -eq "-"){
        #exclusion
        return "Excluded"
    } else {
        #can't determine action
        return "UNKNOWN"
    }
}
###########################################
# Need to finish.
function stripGuid($type, $guid)
{
    $pen = "Get"
    return $pen
}

function identifyPolicyName ($type, $policyGuid, $policies)
{
    if($type -eq "Retention"){
        $policyName = ($policies | ?{$_.Guid -eq $policyGuid}).Name
        if($policyName -ne $null){
            return $policyName
        } else {
            return "UNKNOWN"
        }
    }
}
##########################################

#Declare variables
$connected = $true
$under10MB = $false
$elcNeverRun = $false

#### Verify connectivity
Write-Host -ForegroundColor gray -BackgroundColor black "Connectivity:"
## EXO
Write-Host " Exchange Online PowerShell: " -NoNewLine
try{
    $testCommand = Get-Command Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host -ForegroundColor Green "Connected"
} catch {
    Write-Host -ForegroundColor Red "Not Connected"
    Write-Host -ForegroundColor Red "You must be connected to Exchange Online PowerShell Module."
    try{
        $testCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Red ">> TIP: Run 'Connect-ExchangeOnline'"
    } catch {
        Write-host -ForegroundColor Red ">> It doesn't look like you have EXO PS Module installed!"
        Write-host -ForegroundColor Red ">> TIP: Run 'Install-Module ExchangeOnlineManagement'"
    }
    exit
}

## SCC
Write-Host " Security & Compliance Center PowerShell: " -NoNewLine
try{
    $testCommand = Get-Command Get-RetentionCompliancePolicy -ErrorAction Stop | Out-Null
    Write-Host -ForegroundColor Green "Connected"
} catch {
    Write-Host -ForegroundColor Red "Not Connected"
    Write-Host -ForegroundColor Red "You must be connected to Security & Compliance Center PowerShell Module."
    Write-host -ForegroundColor Red ">> TIP: Run 'Connect-IPPSSession'"
    exit
}

#gather data
Write-Host -ForegroundColor Gray -BackgroundColor Black "Initial Data:"

#test upn
try{
    Write-Host "Mailbox information: " -NoNewLine
    $targetMailbox = Get-Mailbox $mailboxIdentity -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red "Mailbox does not exist or you do not have proper permissions."
    exit
}

#get org config
Try{
    Write-Host "Organization Config: " -NoNewLine
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red "You may not have required permissions."
    exit
}

#get retention policies
Try{
    Write-Host "Retention Policies: " -NoNewLine
    $retentionPolicies = Get-RetentionCompliancePolicy -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red "You may not have required permissions."
    exit
}

Write-Host -ForegroundColor Gray -BackgroundColor Black -NoNewLine "Target Mailbox:"
Write-host " " $targetMailbox.DisplayName

#Test mailbox size
if ((Get-MailboxStatistics $targetMailbox.UserPrincipalName | Select-Object *, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}).TotalItemSizeMB -lt 10){
    Write-Host -ForegroundColor Yellow "WARNING: This mailbox is less than 10MB in size so the managed folder assistant (MFA) will not automatically run."
    $under10MB = $true
}

#Get ELC
$diagLogs = Export-MailboxDiagnosticLogs $targetMailbox.primarysmtpaddress -ExtendedProperties
$xmlProperties = [xml]($diagLogs.MailboxLog)
$ELCLastSuccess = $xmlProperties.Properties.MailboxTable.property | ?{$_.Name -like "ELCLastSuccessTimestamp"}

if($ELCLastSuccess -eq $null){
    write-Host -ForegroundColor Yellow "WARNING: No ELC timestamp found!"
    if($under10MB){
        Write-Host -ForegroundColor Yellow ">> NOTE: This is probably because the mailbox is under 10MB in size."
    }
    $elcNeverRun = $true
} else {
    Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "MFA last run time:"
    Write-Host " " $ELCLastSuccess.value
}

#Check Lithold/Duration
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Litigation Hold Enabled:"
Write-Host " " $targetMailbox.LitigationHoldEnabled
if($targetMailbox.LitigationHoldEnabled){
    Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "- Litigation Hold Duration:"
    Write-Host " " $targetMailbox.LitigationHoldDuration
}
# DelayHold / DelayReleaseHold
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Delay Hold Applied:"
Write-Host " " $targetMailbox.DelayHoldApplied
if($targetMailbox.DelayHoldApplied){
    Write-host -ForegroundColor Yellow ">> NOTE: The Delay Hold will expire after 30 days. Check the mailbox hold history below for an estimated date."
}
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Delay Release Hold Applied:"
Write-Host " " $targetMailbox.DelayReleaseHoldApplied
if($targetMailbox.DelayReleaseHoldApplied){
    Write-host -ForegroundColor yellow ">> NOTE: The Delay Release Hold will expire after 30 days. Check the substrate hold history below for an estimated date."
}

### InplacHolds

### Get Mailbox Hold History ###
$ht = Export-MailboxDiagnosticLogs $targetMailbox.UserPrincipalName -ComponentName HoldTracking
$ds = New-Object System.Data.DataSet

if($ht.MailboxLog.Length -gt 2){
    
    $logEntries = $ht.Mailboxlog | ConvertFrom-Json 
    $logEntries = $logEntries | Sort-Object -Property {$_.lsd}

    $holdLog = New-Object System.Data.DataTable
    $holdLog.TableName = "mailboxHoldHistory"
    
    $holdLog.Columns.Add("Applied") | Out-Null
    $holdLog.Columns.Add("PolicyName") | Out-Null
    $holdLog.Columns.Add("HoldType") | Out-Null
    $holdLog.Columns.Add("PolicyAction") | Out-Null
    $holdLog.Columns.Add("Removed") | Out-Null

    foreach ($logEntry in $logEntries)
    {   
        $row = $holdLog.NewRow()
        $row.Applied = $logEntry.lsd | Get-Date
        $row.HoldType = identifyPolicyOrHold $logEntry.hid $true
        $row.PolicyName = identifyPolicyOrHold $logEntry.hid $false
        #$row.PolicyName = identifyPolicyName $holdType $logEntry.hid $retentionPolicies
        if($holdType -eq "Retention"){
            $row.PolicyAction = identifyRetentionPolicyAction $logEntry.hid
        }
        if($logEntry.ed -ne "0001-01-01T00:00:00.0000000"){
            $row.Removed = $logEntry.ed | Get-Date
        } elseif ($holdType -eq "DelayHoldApplied"){
            $estimatedRemovalStart = ($logEntry.lsd | Get-Date).AddDays(30) | Get-Date -Format "MM/dd/yyyy"
            $estimatedRemovalEnd = ($logEntry.lsd | Get-Date).AddDays(37) | Get-Date -Format "MM/dd/yyyy"
            $row.Removed = "Estimated: ~$estimatedRemovalStart-$estimatedRemovalEnd"
        }
        $holdLog.Rows.Add($row)
    }
    
    $ds.Tables.Add($holdLog)

    #need to fix sort by date
    Write-Host -BackgroundColor black -ForegroundColor gray "Mailbox Hold History:"
    $ds.Tables["mailboxHoldHistory"] | Format-Table #| Sort-Object -Property {$_.Applied} -Descending
} else {
    Write-Host -ForegroundColor Yellow "WARNING: No hold history found for this mailbox!"
    if($elcNeverRun){
        Write-Host -ForegroundColor Yellow ">> NOTE: This is probably because MFA has never run."
    }
}

### Substrate hold history
$hts = Export-MailboxDiagnosticLogs $targetMailbox.UserPrincipalName -ComponentName SubstrateHoldTracking

if($hts.MailboxLog.Length -gt 2){

    $substrateLogEntries = $hts.Mailboxlog | ConvertFrom-Json
    $substrateLogEntries = $substrateLogEntries | Sort-Object -Property {$_.lsd}

    $substrateHoldLog = New-Object System.Data.DataTable
    $substrateHoldLog.TableName = "substrateHoldHistory"

    $substrateHoldLog.Columns.Add("Applied") | Out-Null
    $substrateHoldLog.Columns.Add("PolicyName") | Out-Null
    $substrateHoldLog.Columns.Add("HoldType") | Out-Null
    $substrateHoldLog.Columns.Add("PolicyAction") | Out-Null
    $substrateHoldLog.Columns.Add("Removed") | Out-Null

    foreach ($substrateLogEntry in $substrateLogEntries)
    {   
        $subRow = $substrateHoldLog.NewRow()
        $subRow.Applied = $substrateLogEntry.lsd | Get-Date
        $subRow.HoldType = identifyPolicyOrHold $logEntry.hid $true
        $subRow.PolicyName = identifyPolicyOrHold $logEntry.hid $false
        if($holdType -eq "Retention"){
            $subRow.PolicyAction = identifyRetentionPolicyAction $substrateLogEntry.hid
        }
        if($substrateLogEntry.ed -ne "0001-01-01T00:00:00.0000000"){
            $subRow.Removed = $substrateLogEntry.ed | Get-Date
        } elseif ($holdType -eq "DelayHoldApplied"){
            $estimatedRemovalStart = ($substrateLogEntry.lsd | Get-Date).AddDays(30) | Get-Date -Format "MM/dd/yyyy"
            $estimatedRemovalEnd = ($substrateLogEntry.lsd | Get-Date).AddDays(37) | Get-Date -Format "MM/dd/yyyy"
            $subRow.Removed = "Estimated: ~$estimatedRemovalStart-$estimatedRemovalEnd"
        }
        $substrateHoldLog.Rows.Add($subRow)
    }

    $ds.Tables.Add($substrateHoldLog)

    #need to fix sort by date
    Write-Host -BackgroundColor black -ForegroundColor gray "Substrate Hold History:"
    $ds.Tables["substrateHoldHistory"] | Format-Table # | Sort-Object -Property {$_.Applied} -Descending
} else {
    Write-Host -ForegroundColor Yellow "WARNING: No substrate hold history found for this mailbox!"
    if($elcNeverRun){
        Write-Host -ForegroundColor Yellow ">> NOTE: This is probably because MFA has never run."
    }
}
