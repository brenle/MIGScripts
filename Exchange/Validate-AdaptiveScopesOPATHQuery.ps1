param (
    # You can only provide no value, a RawQuery or an AdaptiveScopeName - not combined
    [ValidateScript({-not ($rawQuery)})][string]$adaptiveScopeName,
    [ValidateScript({-not ($adaptiveScopeName)})][string]$rawQuery,
    [switch]$exportCsv = $false

    ##TODO:  Need to figure out Group/User switch
)

function determineElapsedTime($start, $end){
    $totalTime = $end - $start
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

$rawQueryGetMailboxPassed = $false
$rawQueryGetRecipientPassed = $false
$inactiveMailboxesFound = $false

#verify EXO connectivity
write-Host -ForegroundColor Cyan ".:| Verifying Required Connectivity |:."
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
    Write-Host -ForegroundColor Cyan ".:| Validating OPATH Query using Adaptive Scope |:."
    Write-host ""
    
    try {
        Write-Host "- Looking up Adaptive Scope '" -ForegroundColor Cyan -NoNewLine
        Write-Host -ForegroundColor Gray $adaptiveScopeName -NoNewline
        Write-host -ForegroundColor Cyan "'..." -NoNewline
        $adaptiveScope = Get-AdaptiveScope $adaptiveScopeName -ErrorAction Stop
        Write-Host -ForegroundColor Green "OK"
    } catch {
        Write-Host -ForegroundColor Red "FAILED"
        Write-Host -ForegroundColor Red $error[0]
        exit
    }

    Write-Host -ForegroundColor Cyan "- Scope Type: " -NoNewline
    $scopeType = $adaptiveScope.LocationType
    if($scopeType -eq "Site"){
        Write-host -ForegroundColor Red $scopeType
        Write-Host -ForegroundColor Red ">> ERROR: Site scopes do not support OPATH queries so cannot be used with this script."
        exit
    } else {
        Write-host -ForegroundColor Green $scopeType
    }

    Write-Host -ForegroundColor Cyan "- OPATH Query from '" -NoNewline
    Write-Host -ForegroundColor Gray $adaptiveScopeName -NoNewline
    Write-Host -ForegroundColor Cyan "': " -NoNewline
    if($adaptiveScope.RawQuery -eq ""){
        Write-Host -ForegroundColor Red "No Advanced Query Found"
        Write-Host -ForegroundColor Red ">> ERROR: This script cannot be used to test queries created in the simple query builder."
        exit
    } else {
        Write-Host -ForegroundColor Green $adaptiveScope.RawQuery
        $queryToTest = $adaptiveScope.RawQuery
    }

} elseif ($rawQuery){
    Write-host ""
    Write-Host -ForegroundColor Cyan ".:| Validating Raw OPATH Query |:."
    Write-host ""

    $queryToTest = $rawQuery
    Write-Host -ForegroundColor Cyan "- Query to Validate: " -NoNewline
    Write-host -ForegroundColor Gray $queryToTest
} else {
    Write-host ""
    Write-Host -ForegroundColor Cyan ".:| Validating OPATH Query - Enter a OPATH Query |:."
    Write-host ""

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $queryInputForm = New-Object System.Windows.Forms.Form
    $queryInputForm.Text = "Enter OPATH Query to Validate"
    $queryInputForm.Size = New-Object System.Drawing.Size(400,175)
    $queryInputForm.StartPosition = "CenterScreen"
    
    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Location = New-Object System.Drawing.Point(20,90)
    $validateButton.Text = "Validate"
    $validateButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $queryInputForm.AcceptButton = $validateButton
    $queryInputForm.Controls.Add($validateButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(100,90)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $queryInputForm.AcceptButton = $cancelButton
    $queryInputForm.Controls.Add($cancelButton)

    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Location = New-Object System.Drawing.Point(20,20)
    $instructionLabel.Size = New-Object System.Drawing.Size(350,40)
    $instructionLabel.Text = "Please enter the OPATH query EXACTLY as you would write it in the Advanced Query Builder:"
    
    $queryInputForm.Controls.Add($instructionLabel)

    $queryInputBox = New-Object System.Windows.Forms.TextBox
    $queryInputBox.Location = New-Object System.Drawing.Point(20,60)
    $queryInputBox.Size = New-Object System.Drawing.Size(350,20)

    $queryInputForm.Controls.Add($queryInputBox)

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
        }

    } else {
        # User clicked cancel
        Write-host -ForegroundColor Red "Validation Aborted"
        exit
    }
}

write-host -ForegroundColor Cyan "- Validating RawQuery (Quick)..." -NoNewline
#call function to look for common mistakes
Write-host -ForegroundColor Green "PASSED"

Write-Host -ForegroundColor Cyan "- Validating RawQuery (Full)..." -NoNewLine

try{
    $queryStart = Get-Date
    if($scopeType -eq "User"){
        $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeInactiveMailbox -ErrorAction Stop
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
            $mailboxes = Get-Recipient -RecipientTypeDetails UserMailbox -Filter $queryToTest -ResultSize Unlimited -IncludeSoftDeletedRecipients -ErrorAction Stop
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
    Write-host -ForegroundColor Green "PASSED"
    
    $matchingObjects = ($mailboxes | Measure-Object).Count
    Write-host -ForegroundColor Cyan "- Objects matching query: " -NoNewLine
    
    if($matchingObjects -eq 0){
        Write-host -ForegroundColor Yellow $matchingObjects
        Write-Host -ForegroundColor Yellow ">> NOTE: The query was valid, but returned no results."
        exit
    } else {
        Write-Host -ForegroundColor Green $matchingObjects
    }
    Write-Host -ForegroundColor Cyan "- Total Query Time: " -NoNewline
    Write-Host -ForegroundColor Green (determineElapsedTime $queryStart $queryStop)
    Write-host -ForegroundColor Cyan "- Query Matches Inactive Mailboxes: " -NoNewline
    foreach($mailbox in $mailboxes){
        ##TODO: add progress bar
        if($rawQueryGetMailboxPassed){
            if($mailbox.IsInactiveMailbox){
                $inactiveMailboxesFound = $true
                break
            }
        } else {        
            if((Get-mailbox -IncludeInactiveMailbox $mailbox.UserPrincipalName).IsInactiveMailbox){
                $inactiveMailboxesFound = $true
                break
            }
        }
    }

    if($inactiveMailboxesFound){
        Write-host -ForegroundColor Yellow "YES"
    } else {
        Write-host -ForegroundColor Yellow "NO"
    }

    ### Output Sample Data
    write-host -ForegroundColor Cyan "- Here is a sampling of the result (max 10):"

    $mailboxes | Select-Object -First 10 | ft -a DisplayName, Alias, Identity, PrimarySmtpAddress

    if(!$exportCsv){
        Write-Host -ForegroundColor Yellow "NOTE: Run the script with -ExportCSV if you want to export all results that matched the query."
    } else {
        #export
    }
} else {
    Write-Host -ForegroundColor Red "FAILED"
    Write-Host -ForegroundColor Red $error[0]
    exit
}