---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# Get-EWSAttachment

## SYNOPSIS
Gets attachments from selected item.

## SYNTAX

```
Get-EWSAttachment [[-Item] <Item[]>] [[-Service] <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Function that allows user to get attachment from item or items.
Can be used in combination with Save-EWSAttachment.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> $item = Get-EWSItem -Name Inbox -Filter 'from:Sebastian and hasAttachment:true' -Limit 1
PS C:\> Get-EWSAttachment -Item $item
```

Gets first e-mail sent from 'Sebastian' that has attachments.
List any attachment find in this e-mail.

## PARAMETERS

### -Item
Item object that contains attachment (retrieved using Get-EWSItem).

```yaml
Type: Item[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Service
Service object that will be used to get item/attachment.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Microsoft.Exchange.WebServices.Data.Item[]
Microsoft.Exchange.WebServices.Data.ExchangeService

## OUTPUTS

### Microsoft.Exchange.WebServices.Data.FileAttachment

## NOTES

## RELATED LINKS
