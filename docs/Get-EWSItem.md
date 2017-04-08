---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSItem

## SYNOPSIS
Gets items inside selected mailbox and specified folder. 

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
Functions used to get items from the selected folder inside mailbox.
Filter is required and uses AQL syntax.
You can read more about it here: https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx.

Folders can be specified by Name (for well known folder names, like Inbox) or by Id (works for both well known and custom folders).
For custom folders it's recommended to use Get-EWSFolder command first, and pipe results from it into Get-EWSItem.

By default not all properties are returned - behavior that can be changed by specifying PropertySet.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> Get-EWSFolder Inbox\Urgent | Get-EWSItem -Filter subject:Call
```

Gets all items in sub-folder 'Urgent' in Inbox of last connected service/mailbox.
Items are filtered out and only items that have a word 'call' somewhere in the subject are returned.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> Get-EWSItem -Name Inbox -Filter body:Azure -Limit 5
```

Gets all items in Inbox of last connected service/mailbox.
Items are filtered out and only items that have a word 'Azure' somewhere in the body are returned.
Results are limited to first 5 results.

## PARAMETERS

### -Filter
Filter that will be used to get items needed.
Uses AQL syntax (https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx)

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
Id of the folder that items will be retrieved from.

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
Number of items that should be returned in total.

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
Name of the well known folder (e.g. Calendar, Inbox, Contacts)

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
Number of items tha should be returned per page of results.

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
Optional parameter used to change set of properties returned.

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
Service object that will be used to retrieve items.

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

