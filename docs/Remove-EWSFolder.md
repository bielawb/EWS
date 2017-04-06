---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Remove-EWSFolder

## SYNOPSIS
Removes folder in a specified delete mode.

## SYNTAX

```
Remove-EWSFolder [[-DeleteMode] <DeleteMode>] [[-Service] <ExchangeService>] [-Folder] <Folder> [-WhatIf]
 [-Confirm]
```

## DESCRIPTION
Function used to remove folder in one of three ways:
- HardDelete 
- SoftDelete
- MoveToDeletedItems.
All items in the folder are deleted too.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Get-EWSFolder -Path Inbox\ToDelete | Remove-EWSFolder
```

Prompts user to confirm deletion of Inbox\ToDelete folder.
Once confirmed delete folder using MoveToDeletedItems (default) DeleteMode.

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> $folder = Get-EWSFolder -Path Inbox\ToDelete
PS C:\> Remove-EWSFolder -Folder $folder -DeleteMode HardDelete -Confirm:$false
```

Deletes Inbox\ToDelete folder using HardDelete mode.

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
Mode used to delete folder. If not specified, MoveToDeletedItems mode is used.

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
Folder object that should be removed (retrieved using Get-EWSFolder).

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

### -Service
Service object that will be used to remove folder.

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

