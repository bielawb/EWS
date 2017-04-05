---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Remove-EWSItem

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Remove-EWSItem [[-DeleteMode] <DeleteMode>] [[-Service] <ExchangeService>] [-Item] <Item> [-WhatIf] [-Confirm]
```

## DESCRIPTION
{{Fill in the Description}}

## EXAMPLES

### Example 1
```
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Confirm
Prompts you for confirmation before running the cmdlet.

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
{{Fill DeleteMode Description}}

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
{{Fill Item Description}}

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
{{Fill Service Description}}

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
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

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

