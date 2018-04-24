---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# Get-EWSService

## SYNOPSIS
Function that returns active EWS connections.

## SYNTAX

```
Get-EWSService [<CommonParameters>]
```

## DESCRIPTION
EWS module requires Exchange Web Service connection before performing any operation.
This function allows user to see active connections to EWS.

## EXAMPLES

### Example 1
```powershell
Get-EWSService
```
```
No Active Session Found
```
Before creating connection, function returns information about it.

### Example 2
```powershell
$null = Connect-EWSService @connectionParameters
Get-EWService | Select-Object -Property Url
```
```
Url
---
https://outlook.office365.com/EWS/Exchange.asmx
```
Once connection is created, function retrieves it from the module scope.

## PARAMETERS

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.ExchangeService


## NOTES

## RELATED LINKS
