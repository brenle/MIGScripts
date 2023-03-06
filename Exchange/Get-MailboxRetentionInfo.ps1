param (
    [Parameter(Mandatory = $true)][string]$mailboxIdentity,
    [switch]$enableDebug
)

function identifyPolicyOrHold ([string]$policy, [bool]$typeOnly, [bool]$substrate)
{
    Write-Debug "********"
    Write-Debug "ENTER FUNCTION: identifyPolicyOrHold"
    Write-Debug $policy
    Write-Debug "TypeOnly=$typeOnly"
    Write-Debug "Substrate=$substrate"

    if($policy.substring(0,4) -eq "UniH")
    {
        #eDiscovery Hold
        if($typeOnly){
            return "eDiscovery"
        } else {
            if ($policy.substring(0,4) -eq "UniH"){
                return $policy.trimstart($policy.Substring(0,4))
            } else {
                return
            }
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
                $policyGuid = $policy.trimstart($policy.Substring(0,3))
                Write-Debug "Remove Prefix Result: $policyGuid"
                $policyGuid = $policyGuid.trimend($policyGuid.Substring($policyGuid.Length - 2))
                Write-Debug "Remove Suffix Result: $policyGuid"
                return $policyGuid
            } elseif ($policy.substring(0,4) -eq "-mbx"){
                $policyGuid = $policy.trimstart($policy.Substring(0,4))
                Write-Debug "Remove Prefix Result: $policyGuid"
                #$policyGuid = $policyGuid.trim($policyGuid.Substring($policyGuid.Length - 2))
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
    } elseif ($substrate){
        #substrate
        if($typeOnly){
            return "Retention"
        } else {
            return $policy
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

function identifyPolicyName ($type, $policyGuid, $policies)
{
    if($type -eq "Retention"){

        $policyName = ($policies | ?{$_.Guid -eq $policyGuid}).Name
        if($policyName -ne $null){
            #policy found
            return $policyName
        } else {
            #policy not found - possibly permissions
            return $policyGuid
        }
    } elseif($type -eq "eDiscovery"){
        
        $caseHold = Get-CaseHoldPolicy $policyGuid
        $caseName = ($policies | ?{$_.Identity -eq $caseHold.CaseId}).Name
        #$caseType = ($policies | ?{$_.Identity -eq $caseHold.CaseId}).CaseType
        if($caseName -ne $null){
            #case could be found
            return $caseName
        } else {
            #case not found (probably no permissions)
            return $policyGuid
        }
    }
}

function coreOrAdvanced ($caseName, $policies){
    $type = ($policies | ?{$_.Name -eq $caseName}).CaseType
    return $type
}

function invokeCmdlet ([string]$cmdLet){
    try{
        return Invoke-Expression $cmdLet -ErrorAction Stop
    } catch {
        return Write-Host "ERROR: $($error[0])" -ForegroundColor Red
    }
}

#Declare variables
$under10MB = $false
$elcNeverRun = $false
$sccConnected = $false
$gotOrgConfig = $false
$gotRetentionPolicies = $false
$gotAppRetentionPolicies = $false
$gotLegalCases = $false
$gotAeDLegalCases = $false
$eDiscoveryCases = @()
$retentionPolicies = @()
$excludedPolicies = @()

#enable debugging
write-host $enableDebug
if($enableDebug){
    $debugPref = $DebugPreference
    $DebugPreference = "Continue"
    Write-Debug "Current DebugPreference: $debugPref"
    Write-Debug "Debug enabled"
}

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
    $sccConnected = $true
} catch {
    Write-Host -ForegroundColor Red "Not Connected"
    #Write-Host -ForegroundColor Yellow ">> NOTE: We will proceed but cannot resolve policy names."
    Write-host -ForegroundColor Red ">> TIP: To connect, run 'Connect-IPPSSession'"
    $sccConected = $false
    exit
}

#gather data
Write-Host -ForegroundColor Gray -BackgroundColor Black "Initial Data:"

#test upn
try{
    Write-Host "Mailbox information: " -NoNewLine
    $targetMailbox = Get-Mailbox $mailboxIdentity -IncludeInactiveMailbox -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red $error[0]
    exit
}

#get org config
Try{
    Write-Host "Organization Config: " -NoNewLine
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
    $gotOrgConfig = $true
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red $error[0]
    $gotOrgConfig = $false
    exit
}

#get retention policies
Try{
    Write-Host "Retention Policies: " -NoNewLine
    $sccRetentionPolicies = Get-RetentionCompliancePolicy -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
    $gotRetentionPolicies = $true
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red $error[0]
    $gotRetentionPolicies = $false
    exit
}

if($gotRetentionPolicies -and ($sccRetentionPolicies -ne $null)){
    foreach ($sccRetentionPolicy in $sccRetentionPolicies){
        $retentionPolicies += [pscustomobject] @{
            Guid   = $sccRetentionPolicy.Guid
            Name = $sccRetentionPolicy.Name
        }        
    }
}

#get app retention policies
Try{
    Write-Host "App Retention Policies: " -NoNewLine
    $appRetentionPolicies = Get-AppRetentionCompliancePolicy -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
    $gotAppRetentionPolicies = $true
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Red $error[0]
    $gotAppRetentionPolicies = $false
    exit
}

if($gotAppRetentionPolicies -and ($appRetentionPolicies -ne $null)){
    foreach ($appRetentionPolicy in $appRetentionPolicies){
        $retentionPolicies += [pscustomobject] @{
            Guid   = $appRetentionPolicy.Guid
            Name = $appRetentionPolicy.Name
        }        
    }
}

#get cases
Try{
    Write-Host "eDiscovery Cases: " -NoNewLine
    $eDiscoveryCoreCases = Get-ComplianceCase -CaseType eDiscovery -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
    $gotLegalCases = $true
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-host -ForegroundColor Yellow "You may not have required permissions.  We will continue but will not map case names."
    $gotLegalCases = $false
    #exit
}

if($gotLegalCases -and ($eDiscoveryCoreCases -ne $null)){
    foreach ($ediscoveryCoreCase in $eDiscoveryCoreCases){
        $eDiscoveryCases += [pscustomobject] @{
            Identity   = $eDiscoveryCoreCase.Identity
            Name = $eDiscoveryCoreCase.Name
            CaseType  = $eDiscoveryCoreCase.CaseType
        }
    }
}

#get AED cases
Try{
    Write-Host "Advanced eDiscovery Cases: " -NoNewLine
    $AeDCases = Get-ComplianceCase -CaseType AdvancedEdiscovery -ErrorAction Stop
    Write-host -ForegroundColor Green "OK"
    $gotAeDLegalCases = $true
} catch {
    write-Host -ForegroundColor Red "ERROR"
    write-Host $error[0]
    write-host -ForegroundColor Yellow "You may not have required permissions.  We will continue but will not map case names."
    $gotAeDLegalCases = $false
    #exit
}

if($gotAeDLegalCases -and ($AeDCases -ne $null)){
    foreach($AeDCase in $AeDCases){
        $eDiscoveryCases += [pscustomobject] @{
            Identity   = $AeDCase.Identity
            Name = $AeDCase.Name
            CaseType  = $AeDCase.CaseType
        }
    }
}

Write-Host -ForegroundColor Gray -BackgroundColor Black -NoNewLine "Target Mailbox:"
Write-host " " $targetMailbox.DisplayName

#Test mailbox size
if ((Get-MailboxStatistics $targetMailbox.UserPrincipalName -IncludeSoftDeletedRecipients | Select-Object *, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}).TotalItemSizeMB -lt 10){
    Write-Host -ForegroundColor Yellow "WARNING: This mailbox is less than 10MB in size so the managed folder assistant (MFA) will not automatically run."
    $under10MB = $true
}

#Get ELC
$diagLogs = Export-MailboxDiagnosticLogs $targetMailbox.primarysmtpaddress -ExtendedProperties -IncludeInactiveMailboxes
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
$cmdLet = "Write-Host ' ' $($targetMailbox.LitigationHoldEnabled)"
if($targetMailbox.LitigationHoldEnabled){
    $cmdLet += " -ForegroundColor Yellow"
    $cmdLet += "; Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine '- Litigation Hold Duration:'"
    $cmdLet += "; Write-Host ' ' $($targetMailbox.LitigationHoldDuration)"
}
invokeCmdlet $cmdLet

# DelayHold
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Delay Hold Applied:"
$cmdLet = "Write-Host ' ' $($targetMailbox.DelayHoldApplied)"
if($targetMailbox.DelayHoldApplied){
    $cmdLet += " -ForegroundColor Yellow"
    $cmdLet += "; Write-host -ForegroundColor Yellow '>> NOTE: The Delay Hold will expire after 30 days. Check the mailbox hold history below for an estimated date.'"
}
invokeCmdlet $cmdLet

#DelayReleaseHold
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Delay Release Hold Applied:"
$cmdLet = "Write-Host ' ' $($targetMailbox.DelayReleaseHoldApplied)"
if($targetMailbox.DelayReleaseHoldApplied){
    $cmdLet += " -ForegroundColor Yellow"
    $cmdLet += "; Write-host -ForegroundColor yellow {>> NOTE: The Delay Release Hold will expire after 30 days. Check the substrate hold history below for an estimated date.}"
}
invokeCmdlet $cmdLet

#Check ComplianceTag
Write-Host -BackgroundColor black -ForegroundColor gray -NoNewLine "Retention Label Hold Enabled:"
$cmdLet = "Write-Host ' ' $($targetMailbox.ComplianceTagHoldApplied)"
if($targetMailbox.ComplianceTagHoldApplied){
    $cmdLet += " -ForegroundColor Yellow"
}
invokeCmdlet $cmdLet

### initialize dataset
$ds = New-Object System.Data.DataSet

### In-place holds
Write-Host -BackgroundColor black -ForegroundColor gray "Current In-place Holds:"
if(($targetMailbox.InPlaceHolds | Measure-Object).Count -gt 0){
    $iphLog = New-Object System.Data.DataTable
    $iphLog.TableName = "inplaceHolds"

    $iphLog.Columns.Add("NameOrGuid") | Out-Null
    $iphLog.Columns.Add("HoldType") | Out-Null
    $iphLog.Columns.Add("PolicyAction") | Out-Null

    foreach($inPlaceHold in $targetMailbox.InPlaceHolds){
        $holdType = identifyPolicyOrHold $inPlaceHold $true $false
        $holdGuid = identifyPolicyOrHold $inPlaceHold $false $false
       
        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $policySet = $eDiscoveryCases
        } else {
            $policySet = $retentionPolicies
        }

        $policy = identifyPolicyName $holdType $holdGuid $policySet

        $iphRow = $iphLog.NewRow()
        $iphRow.NameOrGuid = $policy   
        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $iphRow.HoldType = coreOrAdvanced $policy $policySet
        } else {
            $iphRow.HoldType = $holdType
        }
        if($holdType -eq "Retention"){
            $action = identifyRetentionPolicyAction $inPlaceHold
            $iphRow.PolicyAction = $action
            if($action -eq "Excluded"){
                $excludedPolicies += $holdGuid
            }
        }

        $iphLog.Rows.Add($iphRow)
    }

    $ds.Tables.Add($iphLog)
    $ds.Tables["inplaceHolds"] | Format-Table

} else {
    write-Host "There are no in-place holds applied to this mailbox."
}

### orgwide holds
Write-Host -BackgroundColor black -ForegroundColor gray "Organization-wide Holds:"
if(($orgConfig.InPlaceHolds | Measure-Object).Count -gt 0){
    $owhLog = New-Object System.Data.DataTable
    $owhLog.TableName = "orgwideHolds"

    $owhLog.Columns.Add("NameOrGuid") | Out-Null
    $owhLog.Columns.Add("HoldType") | Out-Null
    $owhLog.Columns.Add("PolicyAction") | Out-Null
    $owhLog.Columns.Add('AppliesToThisMailbox') | Out-Null

    foreach($orgWideHold in $orgConfig.InPlaceHolds){
        $holdType = identifyPolicyOrHold $orgWideHold $true $false
        $holdGuid = identifyPolicyOrHold $orgWideHold $false $false
       
        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $policySet = $eDiscoveryCases
        } else {
            $policySet = $retentionPolicies
        }

        $policy = identifyPolicyName $holdType $holdGuid $policySet

        $owhRow = $owhLog.NewRow()
        $owhRow.NameOrGuid = $policy   
        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $owhRow.HoldType = coreOrAdvanced $policy $policySet
        } else {
            $owhRow.HoldType = $holdType
        }
        if($holdType -eq "Retention"){
            $owhRow.PolicyAction = identifyRetentionPolicyAction $orgWideHold 
            
            if($excludedPolicies -contains $holdGuid){
                $owhRow.AppliesToThisMailbox = $false
            } else {
                $owhRow.AppliesToThisMailbox = $true
            }
        }
        
        $owhLog.Rows.Add($owhRow)
    }

    $ds.Tables.Add($owhLog)
    $ds.Tables["orgwideHolds"] | Format-Table

} else {
    write-Host "There are no organization wide holds."
}

### Get Mailbox Hold History ###
$ht = Export-MailboxDiagnosticLogs -IncludeInactiveMailboxes $targetMailbox.UserPrincipalName -ComponentName HoldTracking

if($ht.MailboxLog.Length -gt 2){
    
    $logEntries = $ht.Mailboxlog | ConvertFrom-Json 
    $logEntries = $logEntries | Sort-Object -Property {$_.lsd}

    $holdLog = New-Object System.Data.DataTable
    $holdLog.TableName = "mailboxHoldHistory"
    
    $holdLog.Columns.Add("Applied") | Out-Null
    $holdLog.Columns.Add("NameOrGuid") | Out-Null
    $holdLog.Columns.Add("HoldType") | Out-Null
    $holdLog.Columns.Add("PolicyAction") | Out-Null
    $holdLog.Columns.Add("Removed") | Out-Null

    foreach ($logEntry in $logEntries)
    {   
        $holdType = identifyPolicyOrHold $logEntry.hid $true $false
        $holdGuid = identifyPolicyOrHold $logEntry.hid $false $false

        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $policySet = $eDiscoveryCases
        } else {
            $policySet = $retentionPolicies
        }

        $row = $holdLog.NewRow()
        $row.Applied = $logEntry.lsd | Get-Date
        #$row.HoldType = $holdType
        $policy = identifyPolicyName $holdType $holdGuid $policySet
        $row.NameOrGuid = $policy
        if($holdType -eq "Retention"){
            $row.PolicyAction = identifyRetentionPolicyAction $logEntry.hid
        }
        
        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $row.HoldType = coreOrAdvanced $policy $policySet
        } else {
            $row.HoldType = $holdType
        }
        
        if($logEntry.ed -ne "0001-01-01T00:00:00.0000000"){
            $row.Removed = $logEntry.ed | Get-Date
        } elseif ($holdType -eq "DelayHoldApplied"){
            $estimatedRemovalStart = ($logEntry.lsd | Get-Date).AddDays(30) | Get-Date -Format "MM/dd/yyyy"
            $estimatedRemovalEnd = ($logEntry.lsd | Get-Date).AddDays(37) | Get-Date -Format "MM/dd/yyyy"
            $row.Removed = "ETA: ~$estimatedRemovalStart-$estimatedRemovalEnd"
        }
        $holdLog.Rows.Add($row)
    }
    
    $ds.Tables.Add($holdLog)

    #need to fix sort by date
    Write-Host -BackgroundColor black -ForegroundColor gray "Mailbox Hold History:"
    $ds.Tables["mailboxHoldHistory"] | Format-Table
} else {
    Write-Host -ForegroundColor Yellow "WARNING: No hold history found for this mailbox!"
    if($elcNeverRun){
        Write-Host -ForegroundColor Yellow ">> NOTE: This is probably because MFA has never run."
    }
}

### Substrate hold history
$hts = Export-MailboxDiagnosticLogs $targetMailbox.UserPrincipalName -IncludeInactiveMailboxes -ComponentName SubstrateHoldTracking

if($hts.MailboxLog.Length -gt 2){

    $substrateLogEntries = $hts.Mailboxlog | ConvertFrom-Json
    $substrateLogEntries = $substrateLogEntries | Sort-Object -Property {$_.lsd}

    $substrateHoldLog = New-Object System.Data.DataTable
    $substrateHoldLog.TableName = "substrateHoldHistory"

    $substrateHoldLog.Columns.Add("Applied") | Out-Null
    $substrateHoldLog.Columns.Add("NameOrGuid") | Out-Null
    $substrateHoldLog.Columns.Add("HoldType") | Out-Null
    $substrateHoldLog.Columns.Add("Removed") | Out-Null

    foreach ($substrateLogEntry in $substrateLogEntries)
    {   
        $holdType = identifyPolicyOrHold $substrateLogEntry.hid $true $true
        $holdGuid = identifyPolicyOrHold $SubstrateLogEntry.hid $false $true

        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $policySet = $eDiscoveryCases
        } else {
            $policySet = $retentionPolicies
        }

        $subRow = $substrateHoldLog.NewRow()
        $subRow.Applied = $substrateLogEntry.lsd | Get-Date
        #$subRow.HoldType = $holdType
        $policy = identifyPolicyName $holdType $holdGuid $policySet
        $subRow.NameOrGuid = $policy

        if(($holdType -eq "eDiscovery") -and ($gotLegalCases -eq $true)){
            $subRow.HoldType = coreOrAdvanced $policy $policySet
        } else {
            $subRow.HoldType = $holdType
        }
        
        if($substrateLogEntry.ed -ne "0001-01-01T00:00:00.0000000"){
            $subRow.Removed = $substrateLogEntry.ed | Get-Date
        } elseif ($holdType -eq "DelayHoldApplied"){
            $estimatedRemovalStart = ($substrateLogEntry.lsd | Get-Date).AddDays(30) | Get-Date -Format "MM/dd/yyyy"
            $estimatedRemovalEnd = ($substrateLogEntry.lsd | Get-Date).AddDays(37) | Get-Date -Format "MM/dd/yyyy"
            $subRow.Removed = "ETA: ~$estimatedRemovalStart-$estimatedRemovalEnd"
        }
        $substrateHoldLog.Rows.Add($subRow)
    }

    $ds.Tables.Add($substrateHoldLog)

    #need to fix sort by date
    Write-Host -BackgroundColor black -ForegroundColor gray "Substrate Hold History:"
    $ds.Tables["substrateHoldHistory"] | Format-Table
} else {
    Write-Host -ForegroundColor Yellow "WARNING: No substrate hold history found for this mailbox!"
    if($elcNeverRun){
        Write-Host -ForegroundColor Yellow ">> NOTE: This is probably because MFA has never run."
    }
}

#revert debugging preference if changed
if($enableDebug)
{
    Write-Debug "DebugPreference set back to $DebugPreference"
    $DebugPreference = $debugPref
    
}
