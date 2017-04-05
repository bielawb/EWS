---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSItem

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

### byId (Default)
```
Get-EWSItem [-Filter] <String> -Id <FolderId> [-Service <ExchangeService>] [-PropertySet <BasePropertySet>]
 [-PageSize <Int32>] [-Limit <Int32>]
```

### byName
```
Get-EWSItem [-Filter] <String> -Name <WellKnownFolderName> [-Service <ExchangeService>]
 [-PropertySet <BasePropertySet>] [-PageSize <Int32>] [-Limit <Int32>]
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

### -Filter
{{Fill Filter Description}}

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

### -Id
{{Fill Id Description}}

```yaml
Type: FolderId
Parameter Sets: byId
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Limit
{{Fill Limit Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
{{Fill Name Description}}

```yaml
Type: WellKnownFolderName
Parameter Sets: byName
Aliases: 
Accepted values: Calendar, Contacts, DeletedItems, Drafts, Inbox, Journal, Notes, Outbox, SentItems, Tasks, MsgFolderRoot, PublicFoldersRoot, Root, JunkEmail, SearchFolders, VoiceMail, RecoverableItemsRoot, RecoverableItemsDeletions, RecoverableItemsVersions, RecoverableItemsPurges, ArchiveRoot, ArchiveMsgFolderRoot, ArchiveDeletedItems, ArchiveRecoverableItemsRoot, ArchiveRecoverableItemsDeletions, ArchiveRecoverableItemsVersions, ArchiveRecoverableItemsPurges, SyncIssues, Conflicts, LocalFailures, ServerFailures, RecipientCache, QuickContacts, ConversationHistory, ToDoSearch

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
{{Fill PageSize Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PropertySet
{{Fill PropertySet Description}}

```yaml
Type: BasePropertySet
Parameter Sets: (All)
Aliases: 
Accepted values: IdOnly, FirstClassProperties

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

## INPUTS

### Microsoft.Exchange.WebServices.Data.FolderId
Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Item


## NOTES

## RELATED LINKS

