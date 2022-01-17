# Adaptive Scopes Property Bag Scripts

These scripts are to be used as examples showing how you can use [SharePoint Patterns & Practices (PnP)](https://pnp.github.io/powershell/index.html) to add custom properties to a large number of existing sites in SharePoint Online.

## Requirements

- Ensure you've read the [disclaimer](https://brenle.github.io/MIGScripts/#disclaimer) and [running the scripts](https://brenle.github.io/MIGScripts/#running-the-scripts) sections of this documentation.
- For the **first script** ([Export-SPOSites.ps1](#download)) you will be required to [connect to SharePoint Online using the SharePoint Online PowerShell Module](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online) and will need permissions that allow you to run ```Get-SPOSite```:
- For the **second script** ([Add-BulkPropertyBagValues.ps1](#download)) you will be required to use [PnP.PowerShell module](https://pnp.github.io/powershell/index.html) and will need to be a site collection administrator for each site you want to add custom properties to.
- If this is the first time you are using PnP PowerShell, you will need to first log in interactively and allow permissions:

    ``` powershell
    Connect-PnPOnline -Url https://{tenantName}.sharepoint.com/sites/{sitename} -Interactive
    ```

<figure>
    <img src="../img/pnp-aad-permissions.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 1: When first using PnP PowerShell you must accept the AAD permissions by logging in with the -Interactive switch.</figcaption>
</figure>

<br/>

- Since the purpose of these scripts are to update many existing SharePoint Online sites, you must save your credentials (at least temporarily) in the credential manager.  Follow [these instructions](https://pnp.github.io/powershell/articles/authentication.html#authenticating-with-pre-stored-credentials-using-the-windows-credential-manager-windows-only) to do so.

    !!! note
        This method only works with Windows.  There are other methods available, but you will need to update the scripts to use them.

- Some [optional parameters](#optional-parameters) may require connectivity to the [Exchange Online PowerShell module](https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps) and will require permissions to run ```Get-UnifiedGroup```.

## Usage

### Step 1: Export the existing sites

1. Log in to SharePoint Online PowerShell

    ``` powershell
    Connect-SPOService -Url https://{tenantName}-admin.sharepoint.com
    ```

1. Use [Export-SPOSites.ps1](#download) to export all SPO sites to a CSV file.  Use ```-customKeyToAdd``` to provide the name of the custom property you will be adding to all sites.

    ``` powershell
    .\Export-SPOSites.ps1 -customKeyToAdd customKeyName
    ```

    !!! note
        Replace ```customKeyName``` with whatever custom key name you want to use, such as ```customDepartment``` or ```customProjectName```.

1. All sites that the script exports will be stored in a CSV file. The location will be output once the script is completed.

    !!! note
        The default location and name of the CSV file will be ```c:\temp\SPOSitesExport.csv```.  You can use [optional parameters](#optional-parameters) to change the default location and name.

<figure>
    <img src="../img/step1.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 2: Using Export-SPOSites.</figcaption>
</figure>

<br/>

### Step 2: Add a property bag value for each site in the CSV file

1. Open the CSV file that was created.

1. A column for the custom property you specified with the ```-customKeyToAdd``` switch has been added.

1. Add a value in this column for each site that you want to add the custom property to, then save the CSV file.

    !!! note
        Any site that you add a value for will be processed.  Any site that you do not add a value for will be skipped.  In this example, 4 sites have values set so only 4 sites will be updated.

<figure>
    <img src="../img/step2.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 3: Specifying the custom property values.</figcaption>
</figure>

<br/>

### Step 3: Update the property bag for each site

1. Execute [Add-BulkPropertyBagValues.ps1](#download) using ```-customKeyToAdd``` to specify the name of the custom property added in the previous steps, ```-csvFile``` to provide the path to the CSV file updated in the previous steps, and ```-storedCredential``` to provide the [credential stored in the credential manager](https://pnp.github.io/powershell/articles/authentication.html#authenticating-with-pre-stored-credentials-using-the-windows-credential-manager-windows-only).  This script will automatically connect to each site and add the new custom properties.

    !!! note
        Unless [optional parameters](#optional-parameters) are specified, the script will default to **not** overwriting custom property values if the properties already exist.

    ``` powershell
    .\Add-BulkPropertyBagValues.ps1 -customKeyToAdd "customKeyName" -csvFile c:\temp\SPOSitesExport.csv -storedCredential PropertyBagExample
    ```

<figure>
    <img src="../img/step3a.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 4: Add-BulkPropertyBagValues.ps1 will give a status bar as it is running giving an indication as to how many sites were completed, skipped, and failed.</figcaption>
</figure>

<br/>

<figure>
    <img src="../img/step3b.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 5: Add-BulkPropertyBagValues.ps1 will give a status report after running indicating how many sites were completed, skipped, and failed.</figcaption>
</figure>

<br/>

## Optional parameters

### Export-SPOSites.ps1

- ```customValueToAdd```: You can alternatively have the script automatically populate custom values in the CSV. Keep in mind this will apply to ALL exported sites.
- ```csvExportPath```: Path to export CSV.  Default is c:\temp\.
- ```csvExportFileName```: Name of CSV file.  Default is SPOSitesExport.csv.
- ```identifyTeamsConnectedGroups```: When enabled, will connect to EXO to identify which M365 groups are teams. Default is disabled (```$false```).  Enabling this parameter will require connection to [Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps).
- ```outputAllAvailableSPOSiteProperties```: when enabled, the script will not limit output columns. when disabled, only select columns are output. Default is disabled (```$false```).

#### The following parameters are optional but cannot be combined with each other

- ```TeamsConnectedGroupsOnly```: Outputs **only** M365 groups which are Teams connected. Enabling this parameter will require connection to [Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps). Default is disabled (```$false```).
- ```M365GroupsOnly```: Outputs **only** M365 groups. If ```identifyTeamsConnectedGroups``` is enabled, it will output non-Teams M365 groups.  Default is disabled (```$false```).
- ```SPOSitesOnly```:  Outputs **only** SPO Sites (non-Group connected). Default is disabled (```$false```).

### Update-BulkPropertyBagValues.ps1

- ```overwrite```: If enabled, the script will overwrite any existing property bag values that match the custom property being added.  Default is disabled (```$false```).

## Download

- [Export-SPOSites.ps1](https://github.com/brenle/MIGScripts/blob/main/SPO-OD/AdaptiveScopes-PropertyBag/Export-SPOSites.ps1)
- [Add-BulkPropertyBagValues.ps1](https://github.com/brenle/MIGScripts/blob/main/SPO-OD/AdaptiveScopes-PropertyBag/Add-BulkPropertyBagValues.ps1)

## Changelog

### Export-SPOSites.ps1

##### October 27, 2021 [(0586751)](https://github.com/brenle/MIGScripts/commit/0586751ae8d6cf9934388b7cb0b7b465a73aaec2#diff-ee6b6f5372a92328619ec9cda43aa93c9ba202773bdf0f6a65541d8829825ebe)

- Initial release

### Update-BulkPropertyBagValues.ps1

##### January 14, 2022 [(8394ccd)](https://github.com/brenle/MIGScripts/commit/8394ccdb2e001f9e14960c6ae1adbc0de23a655f)

- Updated with new PnP cmdlet and module version check

##### October 27, 2021 [(0586751)](https://github.com/brenle/MIGScripts/commit/0586751ae8d6cf9934388b7cb0b7b465a73aaec2#diff-ee6b6f5372a92328619ec9cda43aa93c9ba202773bdf0f6a65541d8829825ebe)

- Initial release