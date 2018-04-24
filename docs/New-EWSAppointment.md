---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# New-EWSAppointment

## SYNOPSIS
Creates new appointment.

## SYNTAX

### inline (Default)
```
New-EWSAppointment [-Required <String[]>] [-Optional <String[]>] -Subject <String> -Body <String>
 [-BodyType <BodyType>] [-Location <String>] -Start <DateTime> -Duration <TimeSpan> [-Attachment <String[]>]
 [-Service <ExchangeService>] [<CommonParameters>]
```

### pipe
```
New-EWSAppointment [-Required <String[]>] [-Optional <String[]>] -Subject <String> [-BodyType <BodyType>]
 [-Location <String>] -Start <DateTime> -Duration <TimeSpan> [-Attachment <String[]>] -InputObject <Object>
 [-IsHtml] [-Service <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Function that can be used to create appointments.
I requires following fields:
- Subject
- Body
- Start
- Duration

Optional parameters that can be used:
- Required (required attendees)
- Optional (optional attendees)
- BodyType (type of the body)
- Location
- Attachment

You can pipe in the body of appointment.
If the body is in HTML format, use -IsHtml flag.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> New-EWSAppointment -Subject Dentist -Body 'Visit Dentist' -Start (Get-Date).AddDays(1) -Duration 0:30:0
```

Creates appointment in the calendar about Dentist visit planned for tomorrow that will take half of an hour.

## PARAMETERS

### -Attachment
Paths to attachments that should be added to the new appointment.

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
Body of the appointment.

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
Type of body of the appointment.
Text is used if not specified.

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

### -Duration
Duration of the appointment.

```yaml
Type: TimeSpan
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Content piped into the function.
It's used to form appointment's body.

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
Flag to specify that piped-in content is already formatted as HTML.

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

### -Location
Location of the appointment.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Optional
List of optional attendees.

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

### -Required
List of required attendees.

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

### -Service
Service used to create new appointment.

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

### -Start
Date/time of the appointment start.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Subject of the appointment.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Object
Microsoft.Exchange.WebServices.Data.ExchangeService

## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Appointment

## NOTES

## RELATED LINKS
