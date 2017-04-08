---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# New-EWSMessage

## SYNOPSIS
Creates e-mail message.

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
Function that can be used to create e-mails.
Required parameters:
- To
- Subject
- Body

Latter can be piped in the function.

Optional parameters:
- Cc/Bcc
- BodyType
- Attachment

E-mail is sent and saved to 'Sent Items' folder.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> New-EWSMessage -To MrBeans@gmail.com -Subject Test -Body 'Mic test: 1... 2... 3...'
```

Sends e-mail to MrBeans@gmail.com using last connected service with subject Test and specified body.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> $html = Get-ChildItem | Select-Object -Property Name, Length, LastWriteTime | ConvertTo-Html 
PS C:\> $html | New-EWSMessage -To MrBeans@gmail.com -Subject Test -Body 'Mic test: 1... 2... 3...' -IsHtml -BodyType HTML
```

Generates HTML document from the output of Get-ChildITem
Sends e-mail to MrBeans@gmail.com using last connected service with subject Test and $html as HTML body.


## PARAMETERS

### -Attachment
Paths to attachments that should be added to the message.

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
Background carbon copy - hidden recipients of the e-mail.

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
Body of the e-mail.

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
Type of the e-mail body. The default is Text.

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
Carbon copy - indirect recipients of the e-mail.

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
Input text that will be used as a message body.

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
Flag to specify that command input is already converted to HTML.

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
Service used to create/send message.

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
Subject of the created message.

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
Main recipients of the message.

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

