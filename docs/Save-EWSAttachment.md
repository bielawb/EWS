---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# Save-EWSAttachment

## SYNOPSIS
Saves attachment (or attachments) to disk.

## SYNTAX

```
Save-EWSAttachment [[-Path] <String>] [[-Attachment] <FileAttachment[]>] [[-Service] <ExchangeService>]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Function used to save attachments found in items to disk.
Need to be used in combination with Get-EWSAttachment.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> $item = Get-EWSItem -Name Inbox -Filter 'from:Sebastian and hasAttachment:true' -Limit 1
PS C:\> Get-EWSAttachment -Item $item | Save-EWSAttachment -Path c:\temp\
```

Saves attachment from e-mail to c:\temp\ folder.

### EXAMPLE 2
```
PS C:\> Get-EWSItem -Name Inbox -Filter 'hasAttachment:true' -Limit 1 | Get-EWSAttachment | Save-EWSAttachment -Path c:\temp\
```

Saves attachment from e-mail to c:\temp\ folder using direct pipeline input.

## PARAMETERS

### -Attachment
Attachment object that should be saved to disk (retrieved using Get-EWSAttachment).

```yaml
Type: FileAttachment[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

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

### -Path
Folder where attachments should be saved.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Service object that will be used to get/save attachment.

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

### Microsoft.Exchange.WebServices.Data.FileAttachment[]
Microsoft.Exchange.WebServices.Data.ExchangeService

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
