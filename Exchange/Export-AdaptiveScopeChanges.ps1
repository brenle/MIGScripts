param (
    [Parameter(Mandatory = $true)][string]$AdaptiveScopeName # must be user scope right now
)

$allRecords = @()
$sessionId = "$($AdaptiveScope.Guid)-$(Get-Random)"
$allRecordsFetched = $false
$csvFilePath = "c:\temp\"
$supportedScopeTypes = @('User')

function assembleArrayList($allRecords, $records){
    foreach ($record in $records){
        $allRecords += $record
    }
    return $allRecords
}

# Test connnectivity - EXO
try{
    $testCommand = Get-Command Get-OrganizationConfig -ErrorAction Stop | Out-Null
} catch {
    try{
        $testCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Red "You must be connected to Exchange Online PowerShell Module."
    } catch {
        Write-Host -ForegroundColor Red "You must have the Exchange Online PowerShell Module installed & connected."
    }
    exit
}

## Test connectivity - SCC
try{
    $testCommand = Get-Command Get-RetentionCompliancePolicy -ErrorAction Stop | Out-Null
} catch {
    Write-Host -ForegroundColor Red "You must be connected to SCC PowerShell Module."
    exit
}

try{
    $adaptiveScope = Get-AdaptiveScope $AdaptiveScopeName
    if(!($supportedScopeTypes -contains $adaptiveScope.LocationType)){
        Write-host -ForegroundColor Red "Only the following scope types are currently supported:"
        $supportedScopeTypes
        exit
    }
    $createdDate = $adaptiveScope.WhenCreatedUTC.Date
    $todayDate = (Get-Date).Date
    $csvFileName = "$($AdaptiveScope.Name)-LOG_$(Get-Date -Format "yyyyMMdd-HHmmss").csv"
    if($createdDate -le (Get-Date).AddDays(-365)){
        # scope too old
        write-host -ForegroundColor Red "The scope was created more than 1 year ago."
        exit
    }
} catch {
    #failed to find scope
    write-host -ForegroundColor Red $error.Exception.Message
    exit
}

$a = 1
do{
    Write-Host "Collecting any mactching records from $a-$($a + 99)..." -NoNewline
    $tempRecords = Search-UnifiedAuditLog -StartDate $createdDate -EndDate $todayDate -Operations ApplicableAdaptiveScopeChange -SessionId $sessionId -SessionCommand ReturnLargeSet
    if(($tempRecords | Measure-Object).Count -gt 0){
        Write-Host "$(($tempRecords | measure-object).Count) found."
        $allRecords = assembleArrayList $allRecords $tempRecords
    } else {
        Write-Host "0 found, moving on."
        $allRecordsFetched = $true
    }
    $a += 100
}until($allRecordsFetched)

$totalRecords = ($allRecords | Measure-Object).count
Write-Host "Gathered a total of $totalRecords records."
if($totalRecords -ge 50000){
    Write-Host -ForegroundColor Yellow "WARNING: The maximum number of records has been exceeded.  The results may not be complete."
}

$ds = New-Object System.Data.DataSet

$log = New-Object System.Data.DataTable
$log.TableName = "ScopeLog"
$log.Columns.Add('Id') | Out-Null
$log.Columns.Add("DateTime") | Out-Null
$log.Columns.Add("User") | Out-Null
$log.Columns.Add("State") | Out-Null

$i = 0
foreach($record in $allRecords){
    Write-Progress -Activity "Processing changes for $($AdaptiveScope.Name)" -Status "Analyzing $($record.Id)..." -PercentComplete (($i / $totalRecords) * 100) 
    $recordData = $record.AuditData | ConvertFrom-Json

    foreach($exprop in $recordData.ExtendedProperties){
        
        #only care about the referenced scope
        if($exprop.Value -match $AdaptiveScope.Guid){
            $logRow = $log.NewRow()
            if($exprop.Name -eq "AssociatedAdaptiveScopeIds"){
                $logRow.Id = $recordData.Id #can use this to remove dupes later
                $logRow.DateTime = $recordData.CreationTime
                $logRow.User = $recordData.ObjectId
                $logRow.State = "Added"
            } elseif($exprop.Name -eq "DissociatedAdaptiveScopeIds"){
                $logRow.Id = $recordData.Id #can use this to remove dupes later
                $logRow.DateTime = $recordData.CreationTime
                $logRow.User = $recordData.ObjectId
                $logRow.State = "Removed"
            }
            $log.Rows.Add($logRow)
        }
    }
    $i++
}

$ds.Tables.Add($log)

$totalLogEntries = ($ds.Tables['ScopeLog'] | measure-Object).Count

if($totalLogEntries -gt 0){
    try{
        $path = $csvFilePath + $csvFileName
        $ds.Tables['ScopeLog'] | Sort-Object DateTime | Export-Csv -NoTypeInformation -Path $path
        write-host "Exported $totalLogEntires log entries to $path"
    } catch {
        Write-Host -ForegroundColor Red $error.Exception.Message
    }
} else {
    Write-Host "There were no log entries to export."
}
#$ds.Tables['ScopeLog'] | Format-Table

### Future - export list of current users in scope.
<#
$sortedLog = $ds.Tables['ScopeLog'] | Sort-Object DateTime

$members = New-Object System.Data.DataTable
$members.TableName = "ScopeMembership"
$members.Columns.Add("Added") | Out-Null
$members.Columns.Add("User") | Out-Null

foreach($entry in $sortedLog){
    if($entry.State -eq "Added"){
        $memberRow = $members.NewRow()
        $memberRow.Added = $entry.DateTime
        $memberRow.User = $entry.User
    } else {
        #need to figure out how to now remove
    }
}
#>