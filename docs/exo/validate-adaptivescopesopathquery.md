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

## Download

Access the script [here](https://github.com/brenle/MIGScripts/blob/6c5e7c01c9815d189eda8b81e3ee5a0933477c8d/Exchange/Validate-AdaptiveScopesOPATHQuery.ps1)

## Changelog

##### January 5th, 2022
- Added support for SharedMailbox, EquipmentMailbox and RoomMailbox recipient types
- Rewrote analysis to provide stats for number of shared/resource mailboxes in addition to inactive/incorrectly licensed

##### November 7th, 2021
- Updated documentation link
- Improved detection of inactive mailboxes
- Added total number of inactive mailboxes in query because of improvements
- Added detection of improperly licensed users (**Experimental!** *This may incorrectly report depending on the license or add-on*)

##### November 4th, 2021
- Initial release