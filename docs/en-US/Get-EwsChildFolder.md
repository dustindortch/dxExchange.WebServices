---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Get-EwsChildFolder

## SYNOPSIS
Returns a the child folders of the specified folder object.

## SYNTAX

```
Get-EwsChildFolder [[-Folder] <Object>] [[-FolderView] <FolderView>] [[-Service] <ExchangeService>]
 [<CommonParameters>]
```

## DESCRIPTION
Returns the child folders of the specified folder object.  It defaults to 1000 total child folders, but that can be modified with the FolderView parameter.  Utilizes the FindFolders method of the ExchangeService and only uses the Shallow depth of the Traversal property, as opposed to the Deep depth which would grab all subfolders.

This is depended on by the **Get-EwsFolder** cmdlet.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-EwsChildItem -Folder Inbox
```
```
FolderClass DisplayName TotalCount UnreadCount
----------- ----------- ---------- -----------
IPF.Note    Subfolder   1          1
```

Lists the subfolders of the Inbox folder.

## PARAMETERS

### -Folder
Supplies the parent folder to find the child folders.  Should be of type **Microsoft.Exchange.WebServices.Data.Folder** or **Microsoft.Exchange.WebServices.Data.WellKnownFolderName**.

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

### -FolderView
Provides the maximum number of folders to return.

```yaml
Type: FolderView
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: 1000
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