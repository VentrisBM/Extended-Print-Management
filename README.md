## Overview

This module is created out of necessity for managing advanced features on drivers for large print servers.
Module outputs 13 new cmdlets and every cmdlet is used for specific task. Every cmdlet supports running on remote computer so that you can manage all print servers from one powershell window without PS sessions or RDP.

## Requirements

To run this module you have to run PowerShell 3.0 or later.
Minimum OS Windows 7 SP1 / Windows Server 2008 SP2

To check your PowerShell version run one of these commands
```powershell
PS C:\WINDOWS\system32> get-host

Name             : ConsoleHost
Version          : 5.1.15063.502
.
.
```

```powershell
PS C:\WINDOWS\system32> $PSVersionTable

Name                           Value
----                           -----
PSVersion                      5.1.15063.502
PSEdition                      Desktop
.
.
```

To upgrade PowerShell install _Windows Management Framework_ version 3.0 or greater.
You can find out more about WMF [here](https://docs.microsoft.com/en-us/powershell/wmf/readme).

## Cmdlet Details

All cmdlets accept objects and object attributes from pipe and none of them show any output by default, if you want to see cmdlet output use parameter _-ShowProgress_. Every cmdlet in this module have defined _CmdletBinding_ which means you can use [Common Parameters](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-6) for every cmdlet.

### Set-AdvancedPrintingFeatures

This cmdlet enables you to change state of _“Enable advanced printing features”_ option under _Advanced_ tab of printer driver.
Please note that his cmdlet is not compatible with "Set-PrintDirectly" cmdlet.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-AdvancedPrintingFeatures -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would activate _advanced printing features_ on "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-AdvancedPrintingFeatures -SetActive "True" -ShowProgress
```
This command would activate _advanced printing features_ on every print driver where printer name starts with “HP” on local computer.

```powershell
Set-AdvancedPrintingFeatures -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would activate _advanced printing features_ on every print driver on remote computer.

### Set-EnableBidiSupport

This cmdlet enables you to change state of _“Enable bidirectional support”_ option under _Ports_ tab of printer driver.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-EnableBidiSupport -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would activate _bidirectional support_ on "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-EnableBidiSupport -SetActive "True" -ShowProgress
```
This command would activate _bidirectional support_ on every print driver where printer name starts with “HP” on local computer.

```powershell
Set-EnableBidiSupport -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would activate _bidirectional support_ on every print driver on remote computer.

### Set-KeepDocuments

This cmdlet enables you to change state of _“Keep printed documents”_ option under _Advanced_ tab of printer driver.
Please note that his cmdlet is not compatible with "Set-PrintDirectly" cmdlet.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-KeepDocuments -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would enable _keeping printed documents_ in print queue for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-KeepDocuments -SetActive "True" -ShowProgress
```
This command would enable _keeping printed documents_ in print queue for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-KeepDocuments -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would enable _keeping printed documents_ in print queue for every print driver on remote computer.

### Set-ListInDirectory

This cmdlet enables you to change state of _“List in the directory”_ option under _Sharing_ tab of printer driver.
Please note that this option is only available if print queue is already shared!!

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-ListInDirectory -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would list printer "Test Printer" in active directory.

```powershell
(Get-Printer -Name HP*).Name | Set-ListInDirectory -SetActive "True" -ShowProgress
```
This command would list every printer where printer name starts with “HP” in active directory.

```powershell
Set-ListInDirectory -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would list every printer on remote computer in active directory.

### Set-PrintDirectly

This cmdlet enables you to change state of _“Print directly to the printer”_ option under _Advanced_ tab of printer driver.
Please note that his cmdlet is not compatible with "Set-SpoolPrintDocuments" cmdlet.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-PrintDirectly -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would disable document spooling for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-PrintDirectly -SetActive "True" -ShowProgress
```
This command would disable document spooling for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-PrintDirectly -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would disable document spooling for every print driver on remote computer.

### Set-PrinterAvailability

This cmdlet enables you to set _“Printer Availability”_ option under _Advanced_ tab of printer driver.
Please note that time must be entered in "military" format(ISO 8601 standard).

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-StartTime**
    Specifies availability start time value. Valid range 0000-2359.

    **-UntilTime**
    Specifies availability end time value. Valid range 0000-2359.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-PrinterAvailability -PrinterName "Test Printer" -StartTime 0800 -UntilTime 2000 -ShowProgress
```
This command would set availability for "Test Printer" on the local computer. Printer would be available from 8:00AM to 8:00PM.

```powershell
(Get-Printer -Name HP*).Name | Set-PrinterAvailability -StartTime 1000 -UntilTime 1200 -ShowProgress
```
This command would set availability for every print driver where printer name starts with “HP” on local computer. Printers would be available from 10:00AM to 12:00PM.

```powershell
Set-PrinterAvailability -PrinterName * -StartTime 0800 -UntilTime 1630 -ComputerName "Remote_Computer" -ShowProgress
```
This command would set availability for every print driver on remote computer. Printers would be available from 8:00AM to 4:30PM.

### Set-PrinterComment

This cmdlet enables you to set _“Comment”_ under _General_ tab of printer driver.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-Comment**
    Specifies the string that will be set as comment.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-PrinterComment -PrinterName "Test Printer" -Comment "This is a test printer" -ShowProgress
```
This command would set comment for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-PrinterComment -Comment "This is a test printer" -ShowProgress
```
This command would set comment for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-PrinterComment -PrinterName * -Comment "This is a test printer on remote computer" -ComputerName "Remote_Computer" -ShowProgress
```
This command would set comment for every print driver on remote computer.

### Set-PrinterLocation

This cmdlet enables you to set _“Location”_ under _General_ tab of printer driver.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-Location**
    Specifies the string that will be set as location.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-PrinterLocation -PrinterName "Test Printer" -Location "Location 1" -ShowProgress
```
This command would set location for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-PrinterLocation -Location "Location 2" -ShowProgress
```
This command would set location for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-PrinterLocation -PrinterName * -Location "Location 3" -ComputerName "Remote_Computer" -ShowProgress
```
This command would set location for every print driver on remote computer.

### Set-PrinterPriority

This cmdlet enables you to set _“Print Priority”_ under _Advanced_ tab of printer driver.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-Priority**
    Specifies the integer value that will be used to set priority.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
Set-PrinterPriority -PrinterName "Test Printer" -Priority 5 -ShowProgress
```
This command would set print priority to 5 for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-PrinterPriority -Priority 5 -ShowProgress
```
This command would set print priority to 5 for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-PrinterPriority -PrinterName * -Priority 5 -ComputerName "Remote_Computer" -ShowProgress
```
This command would set print priority to 5 for every print driver on remote computer.

### Set-PrintSpooledFirst

This cmdlet enables you to set _“Print spooled documents first”_ under _Advanced_ tab of printer driver.
Please note that his cmdlet is not compatible with "Set-PrintDirectly" cmdlet.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SetActive**
    Specifies desired state of print driver option.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
 Set-PrintSpooledFirst -PrinterName "Test Printer" -SetActive "True" -ShowProgress
```
This command would activate _“Print spooled documents first”_ option for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-PrintSpooledFirst -SetActive "True" -ShowProgress
```
This command would activate _“Print spooled documents first”_ option for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-PrintSpooledFirst -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress
```
This command would activate _“Print spooled documents first”_ option for every print driver on remote computer.

### Set-SpoolPrintDocuments

This cmdlet enables you to define behaviour of spooled files under _Advanced_ tab of printer driver. Avaliable states are _"Start printing after last page is spooled"_ or _"Start printing immediately"_.
Please note that his cmdlet is not compatible with "Set-PrintDirectly" cmdlet.

#### Parameters

    **-PrinterName**
    Specifies the name of the printer on which to set information.

    **-SelectOption**
    Specifies desired state of printer option. Valid values are "Start_Printing_After_Last_Page" and "Start_Printing_Immediately"

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
 Set-SpoolPrintDocuments -PrinterName "Test Printer" -SelectOption Start_Printing_After_Last_Page -ShowProgress
```
This command would activate "Start printing after last page is spooled" option for "Test Printer" on the local computer.

```powershell
(Get-Printer -Name HP*).Name | Set-SpoolPrintDocuments -SelectOption Start_Printing_Immediately -ShowProgress
```
This command would activate "Start printing immediately" option for every print driver where printer name starts with “HP” on local computer.

```powershell
Set-SpoolPrintDocuments -PrinterName * -SelectOption Start_Printing_Immediately -ComputerName "Remote_Computer" -ShowProgress
```
This command would activate "Start printing immediately" option for every print driver on remote computer.

### Copy-PrinterSettings

This cmdlet enables you to copy printer driver settings that are covered with this module. 

#### Parameters

    **-SourcePrinter**
    Specifies the name of the printer from which to take settings.

    **-TargetPrinter**
    Specifies the name of the printer on which settings will be set.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
 Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "Lexmark" -ShowProgress
```
This command would copy settings from printer queue "HP" to printer queue "Lexmark" on local computer.

```powershell
Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "*" -ShowProgress
```
This command would copy settings from printer queue "HP" to all other printer queues on local computer.

```powershell
Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "*" -ComputerName "Remote_Computer" -ShowProgress
```
This command would copy settings from printer queue "HP" to all other printer queues on remote computer.

### Add-NetworkPrinter

This cmdlet enables you to create new printers manually or by using printer list in CSV format.
Please note that columns in CSV file become parameter for script and because of that they need to be named exactly like it is listed below.
To ensure right column naming copy next line in first row of you CSV file

        PrinterIPAddress,PrinterName,DriverUsed,SharedName,Location,Comment


    Mandatory Parameters in CSV file:
    1. PrinterIPAddress
    2. PrinterName
    3. DriverUsed

    Optional Parameters in CSV file:
    4. SharedName
    5. Location
    6. Comment

#### Parameters

    **-PrinterIPAddress**
    Specifies the IP address for creating TCP/IP printer port.

    **-PrinterName**
    Specifies the name for printer which will be created.

    **-SharedName**
    Specifies shared name for printer which will be created.

    **-DriverUsed**
    Specifies print driver name that will be used for creating printer.

    **-Location**
    Specifies location that will be set on created printer.

    **-Comment**
    Specifies comment that will be set on created printer.

    **-CsvPath**
    Specifies path to CSV file for creating printers.

    **-ComputerName**
    Specifies the target computer for the management operation.

    **-ShowProgress**
    Shows progress of cmdlet to console output.

#### Examples

```powershell
 Add-NetworkPrinter -PrinterIPAddress 0.0.0.0 -PrinterName "Test Printer" -SharedName "TestP" -DriverUsed "Microsoft Print To PDF" -ShowProgress
```
This command would create network printer with name "Test Printer" on local computer.

```powershell
Add-NetworkPrinter -PrinterIPAddress 0.0.0.0 -PrinterName "Test Printer" -SharedName "TestP" -DriverUsed "Microsoft Print To PDF" -ComputerName "Remote_Computer" -ShowProgress
```
This command would create network printer with name "Test Printer" on remote computer and printer would be shared with name "TestP".

```powershell
Add-NetworkPrinter -CsvPath "C:\PrinterList.csv" -ShowProgress
```
This command would create all printers structured in CSV file on local computer.

```powershell
Add-NetworkPrinter -CsvPath "C:\PrinterList.csv" -ComputerName "Remote_Computer" -ShowProgress
```
This command would create all printers structured in CSV file on remote computer.

## Change Log

Version 1.0
    Initial release