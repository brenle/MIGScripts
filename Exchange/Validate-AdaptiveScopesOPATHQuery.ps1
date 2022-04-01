# Written by Brendon Lee (brenle@microsoft.com)
# Please note that this script is provided only as an example script and with no support.
# Full documentation located at: https://brenle.github.io/MIGScripts/exo/validate-adaptivescopesopathquery/
# ----------------------------------------------
# NOTE: You must be connected to Exchange Online PowerShell with permissions to run Get-Mailbox as a minimum to run this script
#
#### USAGE:
#
# To run the script and enter an OPATH query using a GUI:
#     .\Validate-AdaptiveScopesOPATHQuery.ps1
#
# To run the script and extract an OPATH query from an existing scope:
#     .\Validate-AdaptiveScopesOPATHQuery.ps1 -adaptiveScopeName [name of scope]
#
#     NOTE: This option will require being connected to SCC PowerShell with permissions to run Get-AdaptiveScope
#
# To run the script and supply a query via parameter:
#
#     .\Validate-AdaptiveScopesOPATHQuery.ps1 -rawQuery {[OPATH query]} -scopeType [User | Group]
#
#     NOTE: You must include -scopeType when using -rawQuery
#
#### OPTIONAL PARAMETERS:
# -exportCsv = Exports full output of objects that match OPATH query to CSV file. No value is required with this parameter.
# -csvPath [path] = Path to export Csv.  Default value is c:\temp\.
# ----------------------------------------------
param (
    # You can only provide no value, a RawQuery or an AdaptiveScopeName - not combined
    [ValidateScript({-not ($rawQuery)})][string]$adaptiveScopeName,
    [ValidateScript({-not ($adaptiveScopeName)})][string]$rawQuery,
    [string]$scopeType,
    [switch]$exportCsv = $false,
    [string]$csvPath = "c:\temp\",
    [switch]$skipQuickValidation = $false,
    [switch]$skipMixedPropertyDetection = $true
)

function quickValidation($query){
    
    #normalize query
    $query = $query.ToLower()

    #detect surrounding quotes
    if($query.StartsWith('"') -or ($query.StartsWith("'"))){
        Write-Host -ForegroundColor Red "FAIL"
        Write-host -ForegroundColor Red ">> ERROR: OPATH Query cannot be enclosed in quotes."
        exit
    }

    #detect boolean operator 
    if($query.Contains("true")){
        $location = $query.IndexOf("true")
        if($query[$location -1] -eq "$"){
            write-host -ForegroundColor Red "FAIL"
            Write-Host -ForegroundColor Red ">> ERROR: OPATH query cannot contain boolean operators.  Instead use the boolean value, such as 'True'".
            exit
        }
    }
    if($query.Contains("false")){
        $location = $query.IndexOf("false")
        if($query[$location -1] -eq "$"){
            write-host -ForegroundColor Red "FAIL"
            Write-Host -ForegroundColor Red ">> ERROR: OPATH query cannot contain boolean operators.  Instead use the boolean value, such as 'False'".
            exit
        }
    }

    #detect mixed properties
    $getmailbox = @(
        "AcceptMessagesOnlyFrom"
        "AcceptMessagesOnlyFromDLMembers"
        "AdministrativeUnits"
        "AggregatedMailboxGuids"
        "AggregatedMailboxGuidsRaw"
        "AltSecurityIdentities"
        "ArbitrationMailbox"
        "ArchiveName"
        "ArchiveQuota"
        "ArchiveWarningQuota"
        "AuditAdminFlags"
        "AuditDelegateAdminFlags"
        "AuditDelegateFlags"
        "AuditEnabled"
        "AuditLogAgeLimit"
        "AuditOwnerFlags"
        "BypassModerationFrom"
        "BypassModerationFromDLMembers"
        "CalendarLoggingQuota"
        "CalendarRepairDisabled"
        "Certificate"
        "ComplianceTagHoldApplied"
        "DataEncryptionPolicy"
        "DefaultPublicFolderMailbox"
        "DelayHoldApplied"
        "DelayReleaseHoldApplied"
        "DeletedItemFlags"
        "DeliverToMailboxAndForward"
        "DisabledArchiveDatabase"
        "DisabledArchiveGuid"
        "ElcExpirationSuspensionEndDate"
        "ElcExpirationSuspensionStartDate"
        "EnforcedTimestamps"
        "ExchangeSecurityDescriptorRaw"
        "ExchangeUserAccountControl"
        "ExternalOofOptions"
        "ForwardingAddress"
        "ForwardingSmtpAddress"
        "GeneratedOfflineAddressBooks"
        "GenericForwardingAddress"
        "GrantSendOnBehalfTo"
        "ImmutableId"
        "IncludeInGarbageCollection"
        "InPlaceHolds"
        "InPlaceHoldsRaw"
        "IsExcludedFromServingHierarchy"
        "IsHierarchyReady"
        "IsHierarchySyncEnabled"
        "IsInactiveMailbox"
        "IsLinked"
        "IsMachineToPersonTextMessagingEnabled"
        "IsMailboxEnabled"
        "IsMonitoringMailbox"
        "IsPersonToPersonTextMessagingEnabled"
        "IsResource"
        "IsShared"
        "IsSoftDeletedByDisable"
        "IsSoftDeletedByRemove"
        "IssueWarningQuota"
        "JournalArchiveAddress"
        "LanguagesRaw"
        "LastExchangeChangedTime"
        "LegacyExchangeDN"
        "LitigationHoldDate"
        "LitigationHoldOwner"
        "MailboxContainerGuid"
        "MailboxDatabasesRaw"
        "MailboxGuidsRaw"
        "MailboxLocationsRaw"
        "MailboxPlan"
        "MailTipTranslations"
        "MaxBlockedSenders"
        "MaxReceiveSize"
        "MaxSafeSenders"
        "MaxSendSize"
        "MessageHygieneFlags"
        "ModeratedBy"
        "ModerationEnabled"
        "ModerationFlags"
        "NetID"
        "NonCompliantDevices"
        "OfflineAddressBook"
        "OriginalNetID"
        "PasswordLastSetRaw"
        "PersistedCapabilities"
        "PitrEnabled"
        "PreviousDatabase"
        "ProhibitSendQuota"
        "ProhibitSendReceiveQuota"
        "ProtocolSettings"
        "QueryBaseDN"
        "RawIssueWarningQuota"
        "RawProhibitSendQuota"
        "RawProhibitSendReceiveQuota"
        "RawRecoverableItemsQuota"
        "RawRecoverableItemsWarningQuota"
        "RecipientLimits"
        "RecipientSoftDeletedStatus"
        "RecoverableItemsQuota"
        "RecoverableItemsWarningQuota"
        "RejectMessagesFrom"
        "RejectMessagesFromDLMembers"
        "RemoteAccountPolicy"
        "RemotePowerShellEnabled"
        "RequireAllSendersAreAuthenticated"
        "ResourceCapacity"
        "ResourceCustom"
        "ResourcePropertiesDisplay"
        "RetainDeletedItemsFor"
        "RetentionComment"
        "RetentionUrl"
        "RoleAssignmentPolicy"
        "RulesQuota"
        "SCLDeleteThresholdInt"
        "SCLJunkThresholdInt"
        "SCLQuarantineThresholdInt"
        "SCLRejectThresholdInt"
        "SecurityProtocol"
        "SharedEmailDomainStateLastModified"
        "SharedEmailDomainTenant"
        "SharedWithTargetSmtpAddress"
        "SimpleDisplayName"
        "SingleItemRecoveryEnabled"
        "SMimeCertificate"
        "SourceAnchor"
        "StsRefreshTokensValidFrom"
        "TextMessagingState"
        "ThrottlingPolicy"
        "ThumbnailPhoto"
        "TransportSettingFlags"
        "UCSImListMigrationCompleted"
        "UMDtmfMap"
        "UMSpokenName"
        "UnifiedMailboxAccount"
        "UseDatabaseQuotaDefaults"
        "UsnCreated"
        "WasInactiveMailbox"
        "WindowsEmailAddress"
    )
    $getrecipient = @(
        "ActiveSyncMailboxPolicy"
        "BlockedSendersHash"
        "C"
        "City"
        "Co"
        "CoManagedBy"
        "Company"
        "CountryCode"
        "CountryOrRegion"
        "Department"
        "DirSyncAuthorityMetadata"
        "ExpansionServer"
        "ExternalEmailAddress"
        "FirstName"
        "HasActiveSyncDevicePartnership"
        "InformationBarrierSegments"
        "LastName"
        "ManagedBy"
        "Manager"
        "Members"
        "MobileMailboxFlags"
        "Notes"
        "OwaMailboxPolicy"
        "Phone"
        "PostalCode"
        "RawExternalEmailAddress"
        "RawManagedBy"
        "RecipientDisplayTypeRaw"
        "SafeRecipientsHash"
        "SafeSendersHash"
        "ShadowC"
        "ShadowCity"
        "ShadowCo"
        "ShadowCompany"
        "ShadowCountryCode"
        "ShadowDepartment"
        "ShadowFirstName"
        "ShadowLastName"
        "ShadowManager"
        "ShadowNotes"
        "ShadowPhone"
        "ShadowPostalCode"
        "ShadowStateOrProvince"
        "ShadowTitle"
        "SidRaw"
        "StateOrProvince"
        "Title"
        "UMMailboxPolicy"
        "UMRecipientDialPlanId"
        "WhenIBSegmentChanged"
    )
    
    if(!$skipMixedPropertyDetection){
        $GetMailboxCmdlet = $false
        $GetRecipientCmdlet = $false
        #$mixedProperties = @()
        foreach($cmdlet in $getmailbox){
            if($query -match $cmdlet){
                $GetMailboxCmdlet = $true
            }
        }
        foreach($cmdlet in $getrecipient){
            if($query -match $cmdlet){
                $GetRecipientCmdlet = $true
            }
        }
        if($GetMailboxCmdlet -and $GetRecipientCmdlet)
        {
            write-host -ForegroundColor Red "FAIL"
            Write-Host -ForegroundColor Red ">> ERROR: Mixed query detected."
            Write-host -ForegroundColor Red ">> Review: https://brenle.github.io/MIGScripts/exo/validate-adaptivescopesopathquery/#known-limitations"
            exit
        }
    }
}

function determineElapsedTime($start,$end){
    $totalTime = $end - $start
    return $totalTime
}

function formatElapsedTime($totalTime){
    if($totalTime.Hours -ne 0){
        return "$($totalTime.hours):$($totalTime.minutes):$($totalTime.seconds)"
    } elseif ($totalTime.minutes -ne 0){
        return "$($totalTime.minutes):$($totalTime.seconds)"
    } elseif ($totalTime.seconds -ne 0){
        return "$($totalTime.seconds) seconds"
    } else {
        return "$($totalTime.milliseconds) ms"
    }
}

function getCsvFilepath([string]$path,[bool]$cloud){
    
    # path should end with \
    if (!$path.EndsWith("\"))
    {
        $path += "\"
    }

    # path should not be on root drive
    if ($path.EndsWith(":\"))
    {
        $path += "temp\"
    }
    
    # generate file name
    if($cloud){
        $filename = "OPATHQueryResults-Cloud-" + (Get-Date -Format "MMddyyyyHHmmss") + ".csv"
    } else {
        $filename = "OPATHQueryResults-OnPrem-" + (Get-Date -Format "MMddyyyyHHmmss") + ".csv"
    }

    # verify folder exists, if not try to create it
    if (!(Test-Path($path)))
    {
        try
        {
            New-Item -ItemType "directory" -Path $path -ErrorAction Stop | Out-Null
        } catch {
            write-host -ForegroundColor Red "FAILED"
            Write-Host -ForegroundColor Red ">> ERROR: The directory '$path' could not be created."
            Write-Host -ForegroundColor Red $error[0]
            exit
        }
    }

    return $path + $filename
}

$rawQueryGetMailboxPassed = $false
$rawQueryGetRecipientPassed = $false
$rawQueryGetOnPremObjectsPassed = $false
$inactiveMailboxes = 0
$sharedMailboxes = 0
$resourceMailboxes = 0
$userMailboxes = 0
$wrongLicense = 0

Write-host -ForegroundColor Yellow "NOTE: This script is provided only as an example script and with no support."
Write-host ""
#verify EXO connectivity
write-Host -BackgroundColor White -ForegroundColor Black ".:| Verifying Required Connectivity |:."
Write-Host ""

Write-Host -ForegroundColor Cyan "- Exchange Online PowerShell: " -NoNewLine
try{
    $testCommand = Get-Command Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Host -ForegroundColor Green "Connected"
} catch {
    Write-Host -ForegroundColor Red "Not Connected"
    Write-Host -ForegroundColor Red ">> ERROR: You must be connected to Exchange Online PowerShell Module."
    Write-host -ForegroundColor Red ">> You are either not connected or have insufficient permissions."
    try{
        $testCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Red ">> TIP: Run 'Connect-ExchangeOnline'"
    } catch {
        Write-host -ForegroundColor Red ">> It doesn't look like you have EXO PS Module installed!"
        Write-host -ForegroundColor Red ">> TIP: Run 'Install-Module ExchangeOnlineManagement'"
    }
    exit
}

if($adaptiveScopeName){
    
    Write-Host -ForegroundColor Cyan "- Verifying Security & Compliance Center PowerShell Connectivity: " -NoNewLine
    try{
        $testCommand = Get-Command Get-AdaptiveScope -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green "Connected"
    } catch {
        Write-Host -ForegroundColor Red "Not Connected"
        Write-host -ForegroundColor Red ">> ERROR: You must be connected to SCC PowerShell because you used -AdaptiveScopeName"
        Write-host -ForegroundColor Red ">> You are either not connected or have insufficient permissions."
        Write-host -ForegroundColor Red ">> TIP: To connect, run 'Connect-IPPSSession'"
        exit
    }

    Write-host ""
    Write-Host -BackgroundColor White -ForegroundColor Black ".:| Validating OPATH Query using Adaptive Scope |:."
    Write-host ""
    
    try {
        Write-Host "- Looking up Adaptive Scope '" -ForegroundColor Cyan -NoNewLine
        Write-Host -ForegroundColor Gray $adaptiveScopeName -NoNewline
        Write-host -ForegroundColor Cyan "'..." -NoNewline
        $adaptiveScope = Get-AdaptiveScope $adaptiveScopeName -ErrorAction Stop
    } catch {
        Write-Host -ForegroundColor Red "FAILED"
        Write-Host -ForegroundColor Red $error[0]
        exit
    }
    
    if($adaptiveScope.LocationType -eq "Site"){
        Write-host -ForegroundColor Red "FAILED"
        Write-Host -ForegroundColor Red ">> ERROR: Site scopes do not support OPATH queries so cannot be used with this script."
        exit
    } else {
        $scopeType = $adaptiveScope.LocationType
    }

    if($adaptiveScope.RawQuery -eq ""){
        Write-Host -ForegroundColor Red "No Advanced Query Found"
        Write-Host -ForegroundColor Red ">> ERROR: This script cannot be used to test queries created in the simple query builder."
        exit
    } else {
        $queryToTest = $adaptiveScope.RawQuery
        Write-Host -ForegroundColor Green "OK"
    }

} elseif ($rawQuery){
    
    #normalize Scope Type
    $TextInfo = (Get-Culture).TextInfo
    $scopeType = $TextInfo.ToTitleCase($scopeType)
    #I'm not sure if this works in other cultures, so just in case...
    $scopeUser = $TextInfo.ToTitleCase("User")
    $scopeGroup = $TextInfo.ToTitleCase("Group")

    if(($scopeType -ne $scopeUser) -and ($scopeType -ne $scopeGroup)){
        Write-host -ForegroundColor Red ">> ERROR: When using -rawQuery, you MUST provide a valid scope type. For example:"
        Write-host -ForegroundColor Yellow ".\$($MyInvocation.MyCommand.Name) -rawQuery " -NoNewline
        Write-Host -ForegroundColor Gray "[OPATH Query] " -NoNewline
        Write-Host -ForegroundColor Yellow "-scopeType " -NoNewline
        Write-Host -ForegroundColor Gray "User"
        Write-Host -ForegroundColor Cyan "-- or --"   
        Write-host -ForegroundColor Yellow ".\$($MyInvocation.MyCommand.Name) -rawQuery " -NoNewline
        Write-Host -ForegroundColor Gray "[OPATH Query] " -NoNewline
        Write-Host -ForegroundColor Yellow "-scopeType " -NoNewline
        Write-Host -ForegroundColor Gray "Group"
        exit
    }
    Write-host ""
    Write-Host -BackgroundColor White -ForegroundColor Black ".:| Validating Raw OPATH Query |:."
    Write-host ""

    $queryToTest = $rawQuery
} else {
    Write-host ""
    Write-Host -BackgroundColor White -ForegroundColor Black ".:| Validating OPATH Query - Enter a OPATH Query |:."
    Write-host ""

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $queryInputForm = New-Object System.Windows.Forms.Form
    $queryInputForm.Text = "Enter OPATH Query to Validate"
    $queryInputForm.Size = New-Object System.Drawing.Size(415,190)
    $queryInputForm.StartPosition = "CenterScreen"
    
    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Location = New-Object System.Drawing.Point(225,90)
    $validateButton.Text = "Validate"
    $validateButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $validateButton.TabIndex = 3
    
    $queryInputForm.AcceptButton = $validateButton
    $queryInputForm.Controls.Add($validateButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(305,90)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.TabIndex = 4

    $queryInputForm.AcceptButton = $cancelButton
    $queryInputForm.Controls.Add($cancelButton)

    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Location = New-Object System.Drawing.Point(20,20)
    $instructionLabel.Size = New-Object System.Drawing.Size(360,40)
    $instructionLabel.Text = "Please enter the OPATH query EXACTLY as you would write it in the Advanced Query Builder:"
    
    $queryInputForm.Controls.Add($instructionLabel)

    $queryInputBox = New-Object System.Windows.Forms.TextBox
    $queryInputBox.Location = New-Object System.Drawing.Point(20,60)
    $queryInputBox.Size = New-Object System.Drawing.Size(360,20)
    $queryInputBox.TabIndex = 1

    $queryInputForm.Controls.Add($queryInputBox)

    $selectScopeLabel = New-Object System.Windows.Forms.Label
    $selectScopeLabel.Location = New-Object System.Drawing.Point(20,93)
    $selectScopeLabel.Size = New-Object System.Drawing.Size(70,20)
    $selectScopeLabel.Text = "Scope Type: "

    $queryInputForm.Controls.Add($selectScopeLabel)

    $scopeSelection = New-Object System.Windows.Forms.ComboBox
    $scopeSelection.Location = New-Object System.Drawing.Point(90,90)
    $scopeSelection.Size = New-Object System.Drawing.Size(50,20)
    $scopeSelection.DropDownStyle = "Dropdownlist"
    $scopeSelection.Width = 80
    $scopeSelection.TabIndex = 2

    $scopeSelection.Items.Add('User') | Out-Null
    $scopeSelection.Items.Add('Group') | Out-Null
    $scopeSelection.SelectedIndex = 0
    $queryInputForm.Controls.Add($scopeSelection)

    $opathSupportedProperties = New-Object System.Windows.Forms.LinkLabel
    $opathSupportedProperties.Location = New-Object System.Drawing.Point(20,130)
    $opathSupportedProperties.Size = New-Object System.Drawing.Size(160,20)
    $opathSupportedProperties.Text = "Supported OPATH Properties"
    $opathSupportedProperties.add_Click({[system.Diagnostics.Process]::start("https://docs.microsoft.com/en-us/powershell/exchange/filter-properties?view=exchange-ps#filterable-properties")})
    $queryInputForm.Controls.Add($opathSupportedProperties)

    $opathSyntax = New-Object System.Windows.Forms.LinkLabel
    $opathSyntax.Location = New-Object System.Drawing.Point(180,130)
    $opathSyntax.Size = New-Object System.Drawing.Size(85,20)
    $opathSyntax.Text = "OPATH Syntax"
    $opathSyntax.add_Click({[system.Diagnostics.Process]::start("https://docs.microsoft.com/en-us/powershell/exchange/recipient-filters?view=exchange-ps#additional-opath-syntax-information")})
    $queryInputForm.Controls.Add($opathSyntax)

    $opathSyntax = New-Object System.Windows.Forms.LinkLabel
    $opathSyntax.Location = New-Object System.Drawing.Point(270,130)
    $opathSyntax.Size = New-Object System.Drawing.Size(115,20)
    $opathSyntax.Text = "Validating Queries"
    $opathSyntax.Text = "How to use this script"
    $opathSyntax.add_Click({[system.Diagnostics.Process]::start("https://brenle.github.io/MIGScripts/exo/validate-adaptivescopesopathquery/")})
    $queryInputForm.Controls.Add($opathSyntax)

    $queryInputForm.Topmost = $true
    $queryInputForm.AcceptButton = $validateButton
    $queryInputForm.Add_Shown({$queryInputBox.Select()})

    $result = $queryInputForm.ShowDialog()

    if($result -eq [System.Windows.Forms.DialogResult]::OK){
        
        $queryToTest = $queryInputBox.Text
        if(!$queryToTest){
            Write-host -ForegroundColor Red ">> ERROR: You must enter an OPATH query to test! Or, run the script with ONE of the following switches:"
            
            Write-host -ForegroundColor Yellow ".\$($MyInvocation.MyCommand.Name) -adaptiveScopeName " -NoNewline
            Write-Host -ForegroundColor Gray "[Name of Existing Adaptive Scope]"
            Write-Host -ForegroundColor Cyan "-- or --"   
            Write-host -ForegroundColor Yellow ".\$($MyInvocation.MyCommand.Name) -rawQuery " -NoNewline
            Write-host -ForegroundColor Gray "[OPATH Query]"
            exit
        } else {
            if($scopeSelection.SelectedIndex -eq 0)
            {
                $scopeType = "User"
            } else {
                $scopeType = "Group"
            }
        }

    } else {
        # User clicked cancel
        Write-host -ForegroundColor Red "Validation Aborted"
        exit
    }
}

Write-Host -ForegroundColor Cyan "- Query to Validate: " -NoNewline
Write-host -ForegroundColor Gray $queryToTest
Write-Host -ForegroundColor Cyan "- Scope Type: " -NoNewline
WRite-host -ForegroundColor Gray $scopeType

write-host -ForegroundColor Cyan "- Validating RawQuery (Quick)..." -NoNewline
if(!$skipQuickValidation){
    quickValidation $queryToTest
    Write-host -ForegroundColor Green "PASSED"
} else {
    Write-host -ForegroundColor Gray "SKIPPED"
}

Write-Host -ForegroundColor Cyan "- Validating RawQuery (Full)..." -NoNewLine

try{
    $queryStart = Get-Date
    if($scopeType -eq "User"){
        $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeInactiveMailbox -ErrorAction Stop
    } else {
        #must be group
        $mailboxes = Get-Mailbox -GroupMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeInactiveMailbox -ErrorAction Stop
    }
    $rawQueryGetMailboxPassed = $true
    $queryStop = Get-Date
} catch {
    $rawQueryGetMailboxPassed = $false
}

if($rawQueryGetMailboxPassed -eq $false){
    try{
        $queryStart = Get-Date
        if($scopeType -eq "User"){
            $mailboxes = Get-Recipient -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeSoftDeletedRecipients -ErrorAction Stop
        } else {
            #must be group
            $mailboxes = Get-Recipient -RecipientTypeDetails GroupMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeSoftDeletedRecipients -ErrorAction Stop
        }
        $rawQueryGetRecipientPassed = $true
        $queryStop = Get-Date
    } catch {
        $rawQueryGetRecipientPassed = $false
    }
}


if($rawQueryGetMailboxPassed -or $rawQueryGetRecipientPassed){
    
    # 1/11/22 - adding support for onprem mailboxes (mailusers)
    if($scopeType -eq "User"){
        try{
            $onpremQueryStart = Get-Date
            $onPremMailboxes = Get-MailUser -Filter $queryToTest -ResultSize Unlimited -ErrorAction Stop
            $rawQueryGetOnPremObjectsPassed = $true
            $onpremQueryStop = Get-Date
        } catch {
            try{
                $onpremQueryStart = Get-Date
                $onPremMailboxes = Get-Recipient -RecipientTypeDetails MailUser -Filter $queryToTest -ResultSize Unlimited -ErrorAction Stop
                $rawQueryGetOnPremObjectsPassed = $true
                $onpremQueryStop = Get-Date
            } catch {
                #query didn't work for MailUser objects
                $rawQueryGetOnPremObjectsPassed = $false
            }
        }
    }
    
    Write-host -ForegroundColor Green "PASSED"
    
    $matchingObjects = ($mailboxes | Measure-Object).Count
    Write-host -ForegroundColor Cyan "- Cloud objects matching query: " -NoNewLine
    
    if($matchingObjects -eq 0){
        Write-host -ForegroundColor Yellow $matchingObjects
        Write-Host -ForegroundColor Yellow ">> NOTE: The query was valid, but returned no cloud results."
    } else {
        Write-Host -ForegroundColor Green $matchingObjects
    }

    # 1/11/22 - adding support for onprem mailboxes (mailusers)
    Write-Host -ForegroundColor Cyan "- On-premises objects matching query: " -NoNewline
    if($rawQueryGetOnPremObjectsPassed){
        $matchingOnPremObjects = ($onPremMailboxes | Measure-Object).Count
        if($matchingOnPremObjects -eq 0){
            Write-Host -ForegroundColor Yellow $matchingOnPremObjects
            Write-Host -ForegroundColor Yellow ">> NOTE: The query was valid, but returned no on-prem results."
        } else {
            Write-Host -ForegroundColor Green $matchingOnPremObjects
        }
    } else {
        Write-Host -ForegroundColor Yellow "FAILED"
        Write-Host -ForegroundColor Yellow ">> NOTE: The OPATH syntax was not compatible with Get-Recipient which is needed to look for MailUser objects.  This does not mean the query is invalid or will not identify on-prem objects."
    }

    if($matchingObjects -gt 0 -and $matchingOnPremObjects -gt 0){
        Write-Host -ForegroundColor Cyan "- Total objects matching query: " -NoNewline
        Write-host -ForegroundColor Green ($matchingObjects + $matchingOnPremObjects)
    }
    
    # recalculate query time
    if($rawQueryGetOnPremObjectsPassed)
    {
        $totalQueryTime = (determineElapsedTime $queryStart $queryStop) + (determineElapsedTime $onpremQueryStart $onpremQueryStop)
    } else {
        $totalQueryTime = (determineElapsedTime $queryStart $queryStop)
    }

    Write-Host -ForegroundColor Cyan "- Total Query Time: " -NoNewline
    Write-Host -ForegroundColor Green (formatElapsedTime $totalQueryTime)

    #no need to go further if no results
    if($matchingObjects -eq 0 -and $matchingOnPremObjects -eq 0){
        exit
    }

    # 1/11/22 - don't run if there are no resources
    if($matchingObjects -gt 0){
        Write-host ""
        Write-Host -BackgroundColor White -ForegroundColor Black ".:| Checking cloud object properties |:."
        Write-host ""
        # 1/5/22 - adding support for detecting shared/resource mailboxes in addition to inactive mailboxes and collapsing license check
        $i = 1
        foreach($mailbox in $mailboxes){
            Write-Progress -Activity "Analyzing $($mailbox.Identity)..." -Status "Object $i of $matchingObjects" -PercentComplete (($i/$matchingObjects) * 100)

            if($rawQueryGetMailboxPassed){
                if($mailbox.IsInactiveMailbox){
                    $inactiveMailboxes++
                }
                if($mailbox.persistedCapabilities -notcontains "BPOS_S_InformationBarriers"){
                    $wrongLicense++
                }
            } else {        
                if($mailbox.WhenSoftDeleted -ne $null){
                    $inactiveMailboxes++
                }
                if($mailbox.Capabilities -notcontains "BPOS_S_InformationBarriers"){
                    $wrongLicense++
                }
            }
            
            Switch ($mailbox.RecipientTypeDetails)
            {
                "SharedMailbox" {$sharedMailboxes++}
                "RoomMailbox" {$resourceMailboxes++}
                "EquipmentMailbox" {$resourceMailboxes++}
                "UserMailbox" {$userMailboxes++}
            }
            $i++
        }

        Write-host -ForegroundColor Cyan "- Query Matches Cloud User Mailboxes: " -NoNewline
        if($userMailboxes -gt 0){
            Write-host -ForegroundColor Yellow "YES ($userMailboxes)"
        } else {
            Write-host -ForegroundColor Yellow "NO"
        }

        Write-host -ForegroundColor Cyan "- Query Matches Cloud Shared Mailboxes: " -NoNewline
        if($sharedMailboxes -gt 0){
            Write-host -ForegroundColor Yellow "YES ($sharedMailboxes)"
            Write-host -ForegroundColor Magenta ">> TIP: Use 'IsShared' to include/exclude."
        } else {
            Write-host -ForegroundColor Yellow "NO"
        }

        Write-host -ForegroundColor Cyan "- Query Matches Cloud Resource Mailboxes: " -NoNewline
        if($resourceMailboxes -gt 0){
            Write-host -ForegroundColor Yellow "YES ($resourceMailboxes)"
            Write-host -ForegroundColor Magenta ">> TIP: Use 'IsResource' to include/exclude."
        } else {
            Write-host -ForegroundColor Yellow "NO"
        }

        Write-host -ForegroundColor Cyan "- Query Matches Cloud Inactive Mailboxes: " -NoNewline
        if($inactiveMailboxes -gt 0){
            Write-host -ForegroundColor Yellow "YES ($inactiveMailboxes)"
            Write-host -ForegroundColor Magenta ">> TIP: Use 'IsInactiveMailbox' to include/exclude."
        } else {
            Write-host -ForegroundColor Yellow "NO"
        }
        if(!$rawQueryGetMailboxPassed){
            Write-Host -ForegroundColor Yellow ">> WARNING: Get-Recipient was used to verify the query."
            Write-Host -ForeGroundColor Yellow ">> Get-Recipient can only identify recent inactive mailboxes so this may be inaccurate."
        }

        Write-host -ForegroundColor Cyan "- Experimental - Query Matches Incorectly Licensed Users (E/A/G1 or E/A/G3): " -NoNewline
        if($wrongLicense -gt 0){
            Write-host -ForegroundColor Yellow "YES ($wrongLicense)"
        } else {
            Write-host -ForegroundColor Green "NO"
        }
    }
   
    Write-host ""
    Write-Host -BackgroundColor White -ForegroundColor Black ".:| Output |:."
    Write-host ""

    ### Output Sample Data  
    if($matchingObjects -gt 0){
        write-host -ForegroundColor Cyan "- Here is a sampling of the cloud-based results (max 10):"
        $mailboxes | Select-Object -First 10 | ft -a DisplayName, Alias, Identity, PrimarySmtpAddress
    }
    if($matchingOnPremObjects -gt 0){
        write-host -ForegroundColor Cyan "- Here is a sampling of the on-prem results (max 10):"
        $onPremMailboxes | Select-Object -First 10 | ft -a DisplayName, Alias, Identity, PrimarySmtpAddress
    }

    if(!$exportCsv){
        Write-Host -ForegroundColor Yellow "NOTE: Run the script with -ExportCSV if you want to export all objects that matched the query."
    } else {
        if($matchingObjects -gt 0){
            Write-host -ForegroundColor Cyan "Exporting all matching cloud objects to CSV..." -NoNewline
            $csvFile = getCsvFilepath $csvPath $true
            try{
                $mailboxes | Export-Csv -Path $csvFile -NoTypeInformation -ErrorAction Stop
                Write-host -ForegroundColor Green "OK"
                Write-host -ForegroundColor Cyan ">> File location: " -NoNewline
                Write-Host -ForegroundColor Magenta $csvFile
            } catch {
                write-host -ForegroundColor Red "FAILED"
                Write-host -ForegroundColor Red ">> ERROR: Unable to export to '$csvFile'"
                Write-Host -ForegroundColor Red $error[0]
                exit
            }
        }
        if($matchingOnPremObjects -gt 0){
            Write-host -ForegroundColor Cyan "Exporting all matching on-prem objects to CSV..." -NoNewline
            $csvFile = getCsvFilepath $csvPath $false
            try{
                $onPremMailboxes| Export-Csv -Path $csvFile -NoTypeInformation -ErrorAction Stop
                Write-host -ForegroundColor Green "OK"
                Write-host -ForegroundColor Cyan ">> File location: " -NoNewline
                Write-Host -ForegroundColor Magenta $csvFile
            } catch {
                write-host -ForegroundColor Red "FAILED"
                Write-host -ForegroundColor Red ">> ERROR: Unable to export to '$csvFile'"
                Write-Host -ForegroundColor Red $error[0]
                exit
            }
        }
    }
} else {
    Write-Host -ForegroundColor Red "FAILED"
    Write-Host -ForegroundColor Red $error[0]
    exit
}