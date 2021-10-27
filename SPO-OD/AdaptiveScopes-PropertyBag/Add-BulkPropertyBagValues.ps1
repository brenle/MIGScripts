param (
    [Parameter(Mandatory = $true)][string]$customKeyToAdd,
    [Parameter(Mandatory = $true)][string]$csvFile,
    [Parameter(Mandatory = $true)][string]$storedCredential, # must first setup - https://pnp.github.io/powershell/articles/authentication.html#authenticating-with-pre-stored-credentials-using-the-windows-credential-manager-windows-only
    [switch]$overwrite = $false # enable if want to overwrite any existing prop bag values with same key
)

#static variables
$failedSites = 0
$completedSites = 0
$skippedSites = 0
$totalSites = 0
$pnpStillUnlocked = @{}

#verify the CSV provided is valid, then create a log file in the same location
function verifyCsvLocation([string]$csvFilePath)
{
    $tempFilepath = $csvFilePath.ToLower()
    if($tempFilepath.EndsWith(".csv"))
    {
        $pathExists = Test-Path($csvFilePath)
        if ($pathExists)
        {   
            $path = $csvFilePath | Split-Path
            $datetime = Get-Date -Format FileDateTime
            $logFile = $path + "\Add-BulkPropertyBagValuesLog-$datetime.csv"
            try {
                Add-Content -Path $logFile -Value '"DateTime","Url","Succeeded","FailureReason"' -ErrorAction stop
            } catch {
                Write-Host -ForegroundColor Red "Could not create log file"
                write-Host -ForegroundColor Red "Error: $($error[0].exception.message)"
                exit
            }
            return $logFile
        }
    } else {
        Write-Host -ForegroundColor Red "CsvFile string should end in .csv"
        exit
    }
}

#verify pnp online is installed, then verify credential was stored correctly
function verifyModule([string]$cred){

    #first verify pnp is installed
    try{
        $pnpModule = Get-Command Connect-PnPOnline -ErrorAction Stop | Out-Null
    } catch {
        write-host -ForegroundColor Red "PnP Online module not installed."
        exit
    }

    #then verify credential was specified correctly in switch
    $checkIfCredIsStored = Get-PnPStoredCredential -Name $cred
    if(!$checkIfCredIsStored){
        write-host -ForegroundColor red "Credential was not stored.  Store credential using Add-PnPStoredCredential"
        exit
    }
    
}

function logWrite([string]$url, [bool]$result, [string]$reason, $log)
{
    Add-Content -Path $log -Value "$(Get-Date),$url,$result,$reason"
}

#initialization
$LogCsv = verifyCsvLocation $csvFile
verifyModule $storedCredential

#import CSV
try{
    $sites = Import-Csv $csvFile
} catch {
    Write-Host -ForegroundColor Red "Could not import sites from $csvFile"
    write-Host -ForegroundColor Red "Error: $($error[0].exception.message)"
    exit
}

#determine totalSites for progress bar/log
$totalSites = $sites.count
$i = 0

#will cycle through each site in the imported csv and attempt to set the new key value pair
foreach ($site in $sites){
    $success = $true #keeps track of failures
    $siteSkipped = $false #keeps track of any skipped sites
    $pbUnlocked = $false #IMPORTANT: Keeps track of whether NoScriptSite is enabled or disabled on a site
    $keyValue = $site.$customKeyToAdd #value to add
    $i++

    Write-Progress -Activity "Processing site $i : $($site.Url)" -Status "Total Sites: $totalSites; Completed: $completedSites; Failed: $failedSites; Skipped: $skippedSites" -PercentComplete (($i/$totalSites)*100)
    
    #Connect to PnP Online. If failure, note as such.
    if($keyValue -ne ""){
        try {
            Connect-PnPOnline -Url $site.Url -Credentials $storedCredential -ErrorAction Stop  
        } catch {
            logWrite $site.url $false "Could not connect using PnP.  Incorrect stored credential or possibly a site collection permissions issue.  Error: $($error[0].exception.message)" $logCsv
            $success = $false
        }
        #if no failures, capture current property bag to verify if key already exists
        if($success){
            $propertyBag = Get-PnPPropertyBag -key $customKeyToAdd
            if (($propertyBag -eq "") -or ($overwrite -eq $true)){
                # key doesn't exist OR we allow overwrite
                try {
                    # try to unlock property bag
                    Set-PnPSite -Url $site.Url -NoScriptSite $false -ErrorAction Stop
                } catch {
                    # could't unlock property bag
                    logWrite $site.url $false "Could not unlock the property bag.  Error: $($error[0].exception.message)" $logCsv
                    $success = $false
                }
                if($success){
                    # property bag is unlocked
                    $pbUnlocked = $true
                    try {
                        # set key:value pair
                        Set-PnPPropertyBagValue -Key $customKeyToAdd -Value $keyValue -Indexed -ErrorAction Stop
                    } catch {
                        # failed adding key:value pair
                        logWrite $site.url $false "Could not write the key:value pair: $customKeyToAdd : $keyValue. Error: $($error[0].exception.message)" $logCsv
                        $success = $false
                    }
                }
            } else {
                # key already exists - overwrite disabled
                logWrite $site.url $false "A key:value pair already exists and overwrite is disabled: $customKeyToAdd : $($site.$customKeyToAdd)" $logCsv
                $success = $false
            }
            if($pbUnlocked){
                #property bag is still unlocked
                try {
                    Set-PnPSite -Url $site.Url -NoScriptSite $true -ErrorAction Stop
                } catch {
                    # unable to lock property bag
                    logWrite $site.url $false "The key:value pair was written but the property bag could not be locked. Error: $($error[0].exception.message)" $logCsv
                    $pnpStillUnlocked.Add($site.Url,"Error:$($error[0].exception.message)") #add for warning to display after script is complete as this could be security concern
                    $success = $false
                }
            }
            Disconnect-PnPOnline
        }
    } else {
        #skip site if no value is specified in csv (empty cell)
        logWrite $site.Url $false "NOTE: Skipped because no value was specified for this site in the CSV" $logCsv
        $skippedSites++
        $siteSkipped = $true
    }
    if(!$siteSkipped){
        if($success){
            #if not skipped & success still = true, then note as success in log
            logWrite $site.url $true "" $logCsv
            $completedSites++
        } else {
            #else if not skipped & success does not = true, we add to number of failed sites for reporting
            $failedSites++
        }
    }
}
#output info
write-Host "Total Sites: $totalSites"
write-Host "Completed Sites: $completedSites"
Write-host "Failed Sites: $failedSites"
Write-Host "Skipped Sites: $skippedSites"

#note that there were failures or skipped sites
if(($failedSites -gt 0) -and ($failedSites -ne $skippedSites)){
    Write-Host -ForegroundColor Red "There were failures and/or skipped sites.  Check $logCsv for more info."
} elseif (($failedSites -eq 0) -and ($skippedSite -gt 0)){
    Write-Host -ForegroundColor Yellow "There were skipped sites.  Check $logCsv for more info."
}

#make sure we warn user if some sites didn't re-enable NoScript
if($pnpStillUnlocked.Count -gt 0){
    write-host -ForegroundColor red "NOTE: RE-ENABLING NoScriptSite FAILED ON THE FOLLOWING SITES. ENSURE THIS IS REMEDIATED AS SOON AS POSSIBLE AS THIS MAY INTRODUCE A SECURITY RISK!"
    $pnpStillUnlocked
}

