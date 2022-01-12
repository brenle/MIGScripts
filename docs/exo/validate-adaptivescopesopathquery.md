# Validate-AdaptiveScopesOPATHQuery.ps1

This script can be used to validate advanced adaptive scopes queries written in OPATH.

## Requirements

- Ensure you've read the [disclaimer](https://brenle.github.io/MIGScripts/#disclaimer) and [running the scripts](https://brenle.github.io/MIGScripts/#running-the-scripts) sections of this documentation.
- To run this script, you must have the [Exchange Online PowerShell module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module) installed.
- You will be required to at least connect to Exchange Online, and will need permissions that allow you to run ```Get-Mailbox``` and ```Get-Recipient```.
- To connect to Exchange Online using the Exchange Online PowerShell module, run:

``` powershell
Connect-ExchangeOnline
```

- If you use ```-adaptiveScopeName``` you will also need to connect to Security and Compliance Center PowerShell, and will need permissions that allow ou to run ```Get-AdaptiveScope```.
- To connect to the Security and Compliance Center PowerShell module, run:

``` powershell
Connect-IPPSSession
```

## Usage

##### To run the script and enter an OPATH query using a GUI

``` powershell
.\Validate-AdaptiveScopesOPATHQuery.ps1
```

##### To run the script and extract an OPATH query from an existing scope

``` powershell
.\Validate-AdaptiveScopesOPATHQuery.ps1 -adaptiveScopeName [name of scope]
```

##### To run the script and supply a query via parameter

``` powershell
.\Validate-AdaptiveScopesOPATHQuery.ps1 -rawQuery [OPATH query] -scopeType [User | Group]
```
!!! note
    You must include ```-scopeType``` when using ```-rawQuery```

### Optional parameters

- ```-exportCSV```: Exports full output of objects that match OPATH query to CSV file. No value is required with this parameter.
- ```-csvPath [path]```: Path to export Csv.  Default value is c:\temp\

## Known Limitations

- Some properties exist for ```Get-Mailbox``` and some for ```Get-Recipient```.  The script attempts to see if the query works with ```Get-Mailbox``` first, then attempts to use ```Get-Recipient```.  However, if properties are mixed (one that works only with ```Get-Mailbox``` and one that works only with ```Get-Recipient```), the script will not be able to validate the query although mixing properties is supported with adaptive policy scopes.  Review which cmdlet each property works with [here](https://aka.ms/opath-filter).

## Screenshots

<figure>
    <img src="../img/validation-script-no-params.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 1: Executing Validate-AdaptiveScopesOPATHQuery.ps1 with no parameters.</figcaption>
</figure>

<br/>

<figure>
    <img src="../img/validation-script-result.png"/> 
    <figcaption style="font-style: italic; text-align:center;">Figure 2: Validate-AdaptiveScopesOPATHQuery.ps1 results.</figcaption>
</figure>

## Download

Access the script [here](https://github.com/brenle/MIGScripts/blob/main/Exchange/Validate-AdaptiveScopesOPATHQuery.ps1)

## Changelog

##### January 11th, 2022 [(47823d2)](https://github.com/brenle/MIGScripts/commit/47823d2a7238fc6636324aa2c22bdc58fb87c6c4#diff-d66eb284e212b1b6f1b77953502640bbff7c6cf68a740538c399132a83193360)

- Added support for on-prem users in hybrid environment (MailUser)
- Added warning for inactive mailboxes discovered by Get-Recipient
- Added quick validation for mixed properties

##### January 5th, 2022 [(39ad9d4)](https://github.com/brenle/MIGScripts/commit/39ad9d4f80599c69a99318b28aa01ad421d87482#diff-d66eb284e212b1b6f1b77953502640bbff7c6cf68a740538c399132a83193360)

- Added support for SharedMailbox, EquipmentMailbox and RoomMailbox recipient types
- Rewrote analysis to provide stats for number of shared/resource mailboxes in addition to inactive/incorrectly licensed

##### November 7th, 2021 [(6d829d5)](https://github.com/brenle/MIGScripts/commit/6d829d5acf12f5b3a8e43383089106ff2c3b4d51#diff-d66eb284e212b1b6f1b77953502640bbff7c6cf68a740538c399132a83193360)

- Updated documentation link
- Improved detection of inactive mailboxes
- Added total number of inactive mailboxes in query because of improvements
- Added detection of improperly licensed users (**Experimental!** *This may incorrectly report depending on the license or add-on*)

##### November 4th, 2021 [(6c5e7c0)](https://github.com/brenle/MIGScripts/commit/6c5e7c01c9815d189eda8b81e3ee5a0933477c8d#diff-d66eb284e212b1b6f1b77953502640bbff7c6cf68a740538c399132a83193360)


- Initial release