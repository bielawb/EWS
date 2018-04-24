---
external help file: EWS-help.xml
Module Name: EWS
online version:
schema: 2.0.0
---

# New-EWSFolder

## SYNOPSIS
Creates sub-folder in the specified parent folder.

## SYNTAX

### byName (Default)
```
New-EWSFolder -DisplayName <String> [-FolderType <String>] -ParentName <WellKnownFolderName>
 [-Service <ExchangeService>] [-Mailbox <Mailbox>] [<CommonParameters>]
```

### byId
```
New-EWSFolder -DisplayName <String> [-FolderType <String>] -ParentId <FolderId> [-Service <ExchangeService>]
 [-Mailbox <Mailbox>] [<CommonParameters>]
```

## DESCRIPTION
Function used to create folders in the folder structure of a mailbox.
Parent folder can be specified by name (for well known folder names) or Id.

## EXAMPLES

### EXAMPLE 1
```
PS C:\> New-EWSFolder -ParentName Inbox -DisplayName FooBar
```

Create sub-folder FooBar in the folder Inbox.

### EXAMPLE 2
```
PS C:\> Get-EWSFolder -Path Contacts\Parent | New-EWSFolder -DisplayName Child -FolderType Contacts
```

Create sub-folder Child designed to store Contacts inside Contacts\Parent.

## PARAMETERS

### -DisplayName
Display name of the created folder.

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
Type of the folder to create. Supported types:
- Mail (default)
- Calendar
- Contacts
- Tasks

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
Mailbox where folder will be created.

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
Id of the Parent folder (useful for nested folders).

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
Well known name of the parent folder (e.g. Inbox, Calendar).

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
Service object that will be used to create a folder.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String
Microsoft.Exchange.WebServices.Data.FolderId
Microsoft.Exchange.WebServices.Data.ExchangeService

## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Folder

## NOTES

## RELATED LINKS
