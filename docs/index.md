# Welcome to MIGScripts

An assortment of scripts which I've created to help with Microsoft Information Governance (MIG) and Records Management (RM) scenarios.

## Disclaimer

- These scripts are provided as *working (usually) examples*, **as-is, with no support**.  However, if you do find an issue with any of theses scripts, feel free to [submit an issue](https://github.com/brenle/MIGScripts/issues) or fork and submit a pull request with any updates.
- Use these scripts **at your own risk**. 
- Ensure you review the entire script and understand **exactly what it does** before using.  
- Test the scripts within a demo tenant or test environment **before** using in production.  
!!! warning
    I am not responsible for any negative result from using these scripts.  It is your responsibility to fully review and understand what the script will do befor running in your environment.

## Running the scripts

- Given the scripts connect to resources in Microsoft 365, most of these scripts will require some prerequisite modules, such as the [Exchange Online PowerShell module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module).  I will note what is required and/or optional for each script.
- Since the scripts are meant to be simple examples and not production scripts, I do not sign my scripts.  In order to run these scripts, you will need to set the execution policy to unrestricted within PowerShell on your machine:

``` powershell
Set-ExecutionPolicy Unrestricted
```

- Additionally, if you download any script from the internet and attempt to run it, you will first need to unblock the file first.

``` powershell
Unblock-File <filename>
```