---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Clear-EWSFolder

## SYNOPSIS
Clears content of a folder.

## SYNTAX

```
Clear-EWSFolder [[-DeleteMode] <DeleteMode>] [[-Service] <ExchangeService>] [-Folder] <Folder>
 [-IncludeSubfolders] [-WhatIf] [-Confirm]
```

## DESCRIPTION
Function used to clear content of select folder.
By default it will only wipe items in the folder specified.
When used with the flag 'Include Sub-folders' it will wipe both items and folders in selected folder.

User can specify mode that will be used to delete items/folders:
- MoveToDeletedItems (default)
- SoftDelete
- HardDelete.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Get-EWSFolder Inbox\ToDelete | Clear-EWSFolder
```

Clears content of Inbox\ToDelete folder, leaving sub-folders intact.
Will prompt the user before performing the operation.
Uses MoveToDeletedItems mode to remove items (default).

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> Get-EWSFolder Inbox\ToDelete | Clear-EWSFolder -IncludeSubfolders -Confirm:$false -DeleteMode HardDelete
```

Clears content of Inbox\ToDelete folder, including sub-folders and items found in them.
Will not prompt the user before performing the operation.
Uses HardDelete mode to remove items/folders.

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
Mode that should be used when removing items/folders.
Will use MoveToDeletedItems mode when nothing is specified.

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

### -Folder
Folder object that should be moved (retrieved using Get-EWSFolder).

```yaml
Type: Folder
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeSubfolders
Flag to specify if sub-folders and items in them should be removed too.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Service object that will be used to move folder.

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
Microsoft.Exchange.WebServices.Data.Folder


## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS

