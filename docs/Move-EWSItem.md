---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# Move-EWSItem

## SYNOPSIS
Moves item to a different location.

## SYNTAX

```
Move-EWSItem [-DestinationPath] <String> [-Item] <Item> [[-Service] <ExchangeService>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Function used to move item from current location to the new location.
Destination folder has to exist for move to succeed.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> Get-EWSItem -Name Inbox -Filter subject:test -Limit 1 | Move-EWSItem -DestinationPath Inbox\Moved
```

First item found in Inbox with subject that contains word 'test' is moved to Inbox\Moved.

### EXAMPLE 2
```
PS C:\> $item = Get-EWSFolder -Path Inbox\Moved | Get-EWSItem -Filter subject:test -Limit 1
PS C:\> Move-EWSItem -DestinationPath Inbox -Folder $item
```

First item found in Inbox\Moved with subject that contains word 'test' is moved to Inbox.

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
Path to the folder where the item should be moved.

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

### -Item
Item object that should be moved (retrieved using Get-EWSItem).

```yaml
Type: Item
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Service
Service object that will be used to move item.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Microsoft.Exchange.WebServices.Data.Item
Microsoft.Exchange.WebServices.Data.ExchangeService

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
