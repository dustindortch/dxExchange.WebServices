---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Get-EwsFolder

## SYNOPSIS
Returns a single folder object, given a folder or folder path

## SYNTAX

```
Get-EwsFolder [[-FolderPath] <Object>] [[-WellKnownFolder] <WellKnownFolderName>]
 [[-Service] <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Returns a single folder object, given a folder or folder path, that can be used to get folder items or perform other tasks.

## EXAMPLES

### Example 1
```powershell
PS C:\> $Folder = Get-EwsFolder -FolderPath Inbox\Subfolder
PS C:\> $Folder
```

Gets the ~Subfolder~ folder from within the ~Inbox~.

## PARAMETERS

### -FolderPath
The hierarchical path to a leaf folder

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Optionally, specifies an ExchangeService object.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WellKnownFolder
Provide a **WellKnownFolderName** to start as the root of the hierarchy.

```yaml
Type: WellKnownFolderName
Parameter Sets: (All)
Aliases:
Accepted values: Calendar, Contacts, DeletedItems, Drafts, Inbox, Journal, Notes, Outbox, SentItems, Tasks, MsgFolderRoot, PublicFoldersRoot, Root, JunkEmail, SearchFolders, VoiceMail, RecoverableItemsRoot, RecoverableItemsDeletions, RecoverableItemsVersions, RecoverableItemsPurges, ArchiveRoot, ArchiveMsgFolderRoot, ArchiveDeletedItems, ArchiveRecoverableItemsRoot, ArchiveRecoverableItemsDeletions, ArchiveRecoverableItemsVersions, ArchiveRecoverableItemsPurges, SyncIssues, Conflicts, LocalFailures, ServerFailures, RecipientCache, QuickContacts, ConversationHistory, ToDoSearch

Required: False
Position: 1
Default value: None
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