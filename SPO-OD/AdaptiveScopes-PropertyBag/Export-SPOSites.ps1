param (
    [Parameter(Mandatory = $true)][string]$customKeyToAdd,
    [string]$customValueToAdd,
    [string]$csvExportPath = "c:\temp\",
    [string]$csvExportFileName = "SPOSitesExport.csv",
    [switch]$identifyTeamsConnectedGroups = $false, # when enabled, will connect to EXO to identify which M365 groups are teams.
    [switch]$outputAllAvailableSPOSiteProperties = $false, # when enabled, will not limit output columns. when disabled, only select columns are output.
    
    # the following switches cannot be combined
    [ValidateScript({-not ($TeamsConnectedGroupsOnly -or $M365GroupsOnly)})][switch]$SPOSitesOnly = $false, #outputs ONLY SPO sites (no M365 groups/Teams)
    [ValidateScript({-not ($SPOSitesOnly -or $M365GroupsOnly)})][switch]$TeamsConnectedGroupsOnly = $false, #outputs ONLY M365 groups which are Teams connected
    [ValidateScript({-not ($TeamsConnectedGroupsOnly -or $SPOSitesOnly)})][switch]$M365GroupsOnly = $false #outputs ONLY M365 groups. If identifyTeamsConnectedGroups is enabled, it will output non-Teams M365 groups
)
#initialize variables for later
$totalSites = 0
$SPOSites = 0
$M365GroupSites = 0
$TeamsSites = 0

function verifyExportLocation([string]$path,[string]$filename){
    
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
    
    # filename should end in .csv
    $tempFilename = $filename.ToLower()
    if (!$tempFilename.EndsWith(".csv"))
    {
        $filename += ".csv"
    }

    # verify folder exists, if not try to create it
    if (!(Test-Path($path)))
    {
        try
        {
            New-Item -ItemType "directory" -Path $path -ErrorAction Stop | Out-Null
        } catch {
            write-host -ForegroundColor Red "Directory doesn't exist and could not be created.  Exiting."
            exit
        }
    }

    return $path + $filename
}

function verifyConnectivity([bool]$willNeedExo){
    
    #first verify SPO module is installed
    try{
        $spoTenant = Get-Command Get-SpoTenant -ErrorAction Stop | Out-Null
    } catch {
        write-host -ForegroundColor Red "SharePoint Online PowerShell module is not installed."
        exit
    }

    #then verify we are connected
    try{
        $spoTenant = Get-SpoTenant -ErrorAction Stop | Out-Null
    } catch {
        write-host -ForegroundColor Red "Not connected to SharePoint Online.  Connect using Connect-SpoService."
        exit
    }

    #if we need EXO (if Teams info is needed, we also need to connect to EXO)
    if($willNeedExo){
        #first verify EXO module is installed
        try{
            $exoTenant = Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        } catch {
            write-host -ForegroundColor Red "Exchange Online PowerShell module is not installed."
            exit
        }
    
        #then verify we are connected
        try{
            $exoTenant = Get-Command Get-Mailbox -ErrorAction Stop | Out-Null
        } catch {
            write-host -ForegroundColor Red "Not connected to Exchange Online, which is required for identifying Teams connected groups.  Connect using Connect-ExchangeOnline."
            exit
        }
    }
}

# if getting teams groups only, teams identification needs to be enabled
if ($TeamsConnectedGroupsOnly -and !$identifyTeamsConnectedGroups){
    $identifyTeamsConnectedGroups = $true
}

#verify file path and file name for export
$csvExportFile = verifyExportLocation $csvExportPath $csvExportFileName

#verify required modules are installed & connected
verifyConnectivity $identifyTeamsConnectedGroups

#grab full list of SPO sites
$siteList = Get-SpoSite -Limit All
$totalSites = $siteList.Count

#add column for custom key
$siteList = $siteList | Select-Object -Property *, @{label = $customKeyToAdd;expression = {}}

#add column for isM365Group
$siteList = $siteList | Select-Object -Property *, @{label = 'isM365Group';expression = {$false}}

#add column for TeamsConnected if true
if(($identifyTeamsConnectedGroups) -and (!$SPOSitesOnly)){
    $siteList = $siteList | Select-Object -Property *, @{label = 'isTeamsConnected';expression = {$false}}
}

$i = 0
#idenitfy groups & teams
foreach($site in $siteList){
    $i++
    Write-Progress -Activity "Identifying M365 Groups. Processing site $i : $($site.Url)" -Status "Total Sites: $totalSites; SPO Sites:$SPOSites; M365 Group Sites: $M365GroupSites; Teams Sites: $TeamsSites" -PercentComplete (($i/$totalSites)*100)
    if($site.GroupId.Guid -ne "00000000-0000-0000-0000-000000000000"){
        $site.isM365Group = $true
        if(($identifyTeamsConnectedGroups) -and (!$SPOSitesOnly)){
            #see if team
            $M365Group = Get-UnifiedGroup -Identity $site.GroupId.Guid -ErrorAction SilentlyContinue
            if($M365Group.resourceProvisioningOptions -contains "Team"){
                $site.isTeamsConnected = $true
                $TeamsSites++
            } else {
                $M365GroupSites++
            }
        }
    }
    $SPOSites++
}

#remove all but classic SPO sites, if requested
if ($SPOSitesOnly){
    $siteList = $siteList | ?{$_.isM365Group -eq $false}
}

#remove all but M365 groups sites, if requested.  If Teams Identification is enabled, remove Teams as well.
if ($M365GroupsOnly){
    $siteList = $siteList | ?{$_.isM365Group -eq $true}
    if($identifyTeamsConnectedGroups){
        $siteList = $siteList | ?{$_.isTeamsConnected -eq $false}
    }
}

#remove all but Teams sites, if requested, including non-Teams connected M365 groups
if ($TeamsConnectedGroupsOnly){
    $siteList = $siteList | ?{$_.isTeamsConnected -eq $true}
}

#if value is to be added now, then do it now that we have final list
if($customValueToAdd -ne $null){
    foreach ($site in $siteList){
        $site.$customKeyToAdd = $customValueToAdd
    }
}

#export to csv

if ($outputAllAvailableSPOSiteProperties){
    try{
        $siteList | Export-Csv $csvExportFile -NoTypeInformation
    } catch {
        Write-Host -ForegroundColor Red "Unable to export to $csvExportFile!"
        write-Host -ForegroundColor Red "Error: $($error[0].exception.message)"
        exit
    }
} else {
    if(($identifyTeamsConnectedGroups) -and (!$SPOSitesOnly)){
        try{
            $siteList | Select Title, URL, $customKeyToAdd, Status, LastContentModifiedDate, Owner, isM365Group,isTeamsConnected | Export-Csv $csvExportFile -NoTypeInformation -ErrorAction Stop
        } catch {
            Write-Host -ForegroundColor Red "Unable to export to $csvExportFile!"
            write-Host -ForegroundColor Red "Error: $($error[0].exception.message)"
            exit
        }
    } else {
        try{
            $siteList | Select Title, URL, $customKeyToAdd, Status, LastContentModifiedDate, Owner, isM365Group | Export-Csv $csvExportFile -NoTypeInformation
        } catch {
            Write-Host -ForegroundColor Red "Unable to export to $csvExportFile!"
            write-Host -ForegroundColor Red "Error: $($error[0].exception.message)"
            exit
        }
    }
}

#output details
write-Host "Total SPO Sites Found: $totalSites"
write-host "Total SPO Sites Exported: $($siteList.Count)"
write-Host "Exported file located at $csvExportFile"
Write-host ""
Write-Host "Next step: Add/Verify custom values for $customKeyToAdd in $csvExportFile."
Write-host "Once done, run the next script to update the property bags using the following:"
write-host ""
write-host "Add-BulkPropertyBagValues -csvFile $csvExportFile -customKeyToAdd '$customKeyToAdd'"