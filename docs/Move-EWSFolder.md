---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Move-EWSFolder

## SYNOPSIS
Moves folder to a different location.

## SYNTAX

```
Move-EWSFolder [-DestinationPath] <String> [-Folder] <Folder> [[-Service] <ExchangeService>] [-WhatIf]
 [-Confirm]
```

## DESCRIPTION
Function used to move folder from current location to the new location.
Destination folder has to exist for move to succeed.


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Get-EWSFolder -Path Inbox\Test | Move-EWSFolder -DestinationPath Inbox\Parent
```

Folder 'Test' located inside Inbox is moved as a sub-folder to Inbox\Parent.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> $folder = Get-EWSFolder -Path Inbox\Parent\Test 
PS C:\> Move-EWSFolder -DestinationPath Inbox -Folder $folder
```

Folder 'Test' located inside Inbox\Parent is moved as a sub-folder to Inbox.

## PARAMETERS

### -Confirm
Prompts you for confirmation before running the function.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationPath
Path to the folder where the folder should be moved.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Folder
Folder object that should be moved (retrieved using Get-EWSFolder).

```yaml
Type: Folder
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Service
Service object that will be used to move folder.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the function runs.
The function is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### Microsoft.Exchange.WebServices.Data.Folder
Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS

