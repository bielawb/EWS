---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSAttachment

## SYNOPSIS
Gets attachments from selected item.

## SYNTAX

```
Get-EWSAttachment [[-Item] <Item[]>] [[-Service] <ExchangeService>]
```

## DESCRIPTION
Function that allows user to get attachment from item or items.
Can be used in combination with Save-EWSAttachment.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
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

## INPUTS

### Microsoft.Exchange.WebServices.Data.Item[]
Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.FileAttachment


## NOTES

## RELATED LINKS

