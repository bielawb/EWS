---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Connect-EWSService

## SYNOPSIS
Connects to Exchange Web Service.

## SYNTAX

```
Connect-EWSService [-Mailbox] <String> [[-ServiceUrl] <String>] [[-Version] <ExchangeVersion>]
 [[-Credential] <PSCredential>]
```

## DESCRIPTION
Function that needs to be used first in order to create connection to Exchange Web Service.
It supports both user-provided credentials and default credentials.
By default it will attempt to auto-discover service URL. User may decide to provide it manually.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Connect-EWSService -Mailbox bartek.bielawski@live.com
```
Connects to bartek.bielawski@live.com mailbox.
Uses default credentials and auto-discovers service URL.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> Connect-EWSService -Mailbox bartek.bielawski@live.com -Credential bartek.bielawski@live.com
```
Connects to bartek.bielawski@live.com mailbox.
Prompts for bartek.bielawski@live.com password. 
Auto-discovers service URL.

## PARAMETERS

### -Credential
Credentials used to authenticate to Exchange Web Service.

```yaml
Type: PSCredential
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mailbox
Mailbox that supports Exchange Web Services.

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

### -ServiceUrl
Optional URL to service (when specified, auto-discovery is not used).

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Version
Version of Exchange Web Service.

```yaml
Type: ExchangeVersion
Parameter Sets: (All)
Aliases: 
Accepted values: Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013, Exchange2013_SP1

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### None


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.ExchangeService


## NOTES

## RELATED LINKS

