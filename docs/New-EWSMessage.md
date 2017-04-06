---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# New-EWSMessage

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

### inline (Default)
```
New-EWSMessage -To <String[]> [-Cc <String[]>] [-Bcc <String[]>] -Subject <String> -Body <String>
 [-BodyType <BodyType>] [-Attachment <String[]>] [-Service <ExchangeService>]
```

### pipe
```
New-EWSMessage -To <String[]> [-Cc <String[]>] [-Bcc <String[]>] -Subject <String> [-BodyType <BodyType>]
 [-Attachment <String[]>] -InputObject <Object> [-IsHtml] [-Service <ExchangeService>]
```

## DESCRIPTION
{{Fill in the Description}}

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Attachment
{{Fill Attachment Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bcc
{{Fill Bcc Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Body
{{Fill Body Description}}

```yaml
Type: String
Parameter Sets: inline
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyType
{{Fill BodyType Description}}

```yaml
Type: BodyType
Parameter Sets: (All)
Aliases: 
Accepted values: HTML, Text

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Cc
{{Fill Cc Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
{{Fill InputObject Description}}

```yaml
Type: Object
Parameter Sets: pipe
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IsHtml
{{Fill IsHtml Description}}

```yaml
Type: SwitchParameter
Parameter Sets: pipe
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
{{Fill Service Description}}

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Subject
{{Fill Subject Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
{{Fill To Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### System.Object
Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.EmailMessage


## NOTES

## RELATED LINKS

