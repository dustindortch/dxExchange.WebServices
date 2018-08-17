---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Get-EwsWellKnownFolder

## SYNOPSIS
Get the object of a well known folder in an Exchange mailbox.

## SYNTAX

```
Get-EwsWellKnownFolder [-FolderName] <WellKnownFolderName> [-Service <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Get the object of a well known folder in an Exchange mailbox, such as the Inbox, Sent Items, Deleted Items, etc.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-EwsWellKnownFolder -FolderName Inbox
```
```
FolderClass    DisplayName                      TotalCount   UnreadCount
-----------    -----------                      ----------   -----------
IPF.Note       Inbox                            76496        60099
```

Returns the Inbox folder within the Exchange mailbox.

## PARAMETERS

### -FolderName
Supply a Folder object.  Can be tab completed as it uses the type accelerator.  Only valid well known folders can be passed.

```yaml
Type: WellKnownFolderName
Parameter Sets: (All)
Aliases: Folder
Accepted values: Calendar, Contacts, DeletedItems, Drafts, Inbox, Journal, Notes, Outbox, SentItems, Tasks, MsgFolderRoot, PublicFoldersRoot, Root, JunkEmail, SearchFolders, VoiceMail, RecoverableItemsRoot, RecoverableItemsDeletions, RecoverableItemsVersions, RecoverableItemsPurges, ArchiveRoot, ArchiveMsgFolderRoot, ArchiveDeletedItems, ArchiveRecoverableItemsRoot, ArchiveRecoverableItemsDeletions, ArchiveRecoverableItemsVersions, ArchiveRecoverableItemsPurges, SyncIssues, Conflicts, LocalFailures, ServerFailures, RecipientCache, QuickContacts, ConversationHistory, ToDoSearch

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Optionally, specifies an ExchangeService object to set the impersonation mailbox.

```yaml
Type: Microsoft.Exchange.WebServices.Data.ExchangeService
Parameter Sets: (All)
Aliases: None

Required: False
Position: Named
Default value: $Script:Service
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Folder

## NOTES

## RELATED LINKS

[Microsoft.Exchange.WebServices.Data.Folder](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.folder?view=exchange-ews-api)
