---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Remove-EWSItem

## SYNOPSIS
Removes item in a specified delete mode.

## SYNTAX

```
Remove-EWSItem [[-DeleteMode] <DeleteMode>] [[-Service] <ExchangeService>] [-Item] <Item> [-WhatIf] [-Confirm]
```

## DESCRIPTION
Function used to remove item in one of three ways:
- HardDelete 
- SoftDelete
- MoveToDeletedItems.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Get-EWSFolder -Path Inbox\ToDelete | Get-EWSItem -Filter IsRead:true | Remove-EWSItem
```

Prompts user to confirm deletion of read (IsRead:true) items in Inbox\ToDelete folder.
Once confirmed delete items using MoveToDeletedItems (default) DeleteMode.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> $item = Get-EWSFolder -Path Inbox\ToDelete | Get-EWSItem -Filter IsRead:true -Limit 1
PS C:\> Remove-EWSItem -Folder $item -DeleteMode HardDelete -Confirm:$false
```

Deletes first read (IsRead:true) item from Inbox\ToDelete folder using HardDelete mode.

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

### -DeleteMode
Mode used to delete item. If not specified, MoveToDeletedItems mode is used.

```yaml
Type: DeleteMode
Parameter Sets: (All)
Aliases: 
Accepted values: HardDelete, SoftDelete, MoveToDeletedItems

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Item
Item object that should be removed (retrieved using Get-EWSItem).

```yaml
Type: Item
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Service
Service object that will be used to remove item.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
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

### Microsoft.Exchange.WebServices.Data.ExchangeService
Microsoft.Exchange.WebServices.Data.Item


## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS

