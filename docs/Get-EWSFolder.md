---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# Get-EWSFolder

## SYNOPSIS
Gets folder inside selected mailbox.

## SYNTAX

```
Get-EWSFolder [-Path] <String> [[-Service] <ExchangeService>] [[-Mailbox] <Mailbox>] [<CommonParameters>]
```

## DESCRIPTION
Function that is used to get folder in given path inside selected mailbox.
The path needs to start with one of the well known folders (such as Inbox, or Calendar).
Behavior is different depending on the completeness of the path:
- complete path specified - target folder returned
- partial path specified - parent and all matching child folders will be returned.

If service/mailbox are not specified, default values are used:
- for service - last service that user connected to in the current session
- for mailbox - matching mailbox from the last connected service.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> Get-EWSFolder -Path Inbox
```

Object representing Inbox in the last connected service is returned.

### EXAMPLE 2
```
PS C:\> Get-EWSFolder -Path Inbox\A
```

Result depends on the presence of folder A.
If folder 'A' doesn't exist, object representing Inbox in the last connected service is returned.
If there is any folder underneath the Inbox with name that matches A, it's returned too.
Finally if A exists, only that folder is returned.

### EXAMPLE 3
```
PS C:\> Get-EWSFolder -Path Calendar -Service $myService
```

Object representing Calendar in the $myService service is returned.

## PARAMETERS

### -Mailbox
Mailbox where folder is located.

```yaml
Type: Mailbox
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to folder inside the mailbox with two formats supported:
WellKnownFolder\Full\Match - returns selected folder
WellKnownFolder\Partial\MatchTo - returns parent folder and any folder matching to MatchTo.

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

### -Service
Service used to communicate with selected mailbox.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Folder

## NOTES

## RELATED LINKS
