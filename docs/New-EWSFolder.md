---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# New-EWSFolder

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

### byName (Default)
```
New-EWSFolder -DisplayName <String> [-FolderType <String>] -ParentName <WellKnownFolderName>
 [-Service <ExchangeService>] [-Mailbox <Mailbox>]
```

### byId
```
New-EWSFolder -DisplayName <String> [-FolderType <String>] -ParentId <FolderId> [-Service <ExchangeService>]
 [-Mailbox <Mailbox>]
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

### -DisplayName
{{Fill DisplayName Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -FolderType
{{Fill FolderType Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: 
Accepted values: Calendar, Contacts, Tasks, Mail

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Mailbox
{{Fill Mailbox Description}}

```yaml
Type: Mailbox
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ParentId
{{Fill ParentId Description}}

```yaml
Type: FolderId
Parameter Sets: byId
Aliases: Id

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -ParentName
{{Fill ParentName Description}}

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

### System.String
Microsoft.Exchange.WebServices.Data.FolderId
Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Folder


## NOTES

## RELATED LINKS

