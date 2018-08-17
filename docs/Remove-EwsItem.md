---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Remove-EwsItem

## SYNOPSIS
Removes the specified item(s).

## SYNTAX

```
Remove-EwsItem [-Item] <Item[]> [-DeleteMode <DeleteMode>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes the specified item(s) from the mailbox, moving the item to the Deleted Items folder, by default.

## EXAMPLES

### Example 1
```
PS C:\> Remove-EwsItem -Item $Item
```

Moves the specified item to the Deleted Items folder.

### Example 2
```
PS C:\> Remove-EwsItem -Item $Item -DeleteMode SoftDelete
```
```
Are you sure that you want to SoftDelete the item?
Microsoft.Exchange.WebServices.Data.EmailMessage
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):
```

Deletes the specified item, but it is retained the the Recoverable Items folder.

### Example 3
```
PS C:\> Remove-EwsItem -Item $Item -DeleteMode HardDelete
```
```
Are you sure that you want to HardDelete the item?
Microsoft.Exchange.WebServices.Data.EmailMessage
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):
```

Deletes the specified item in an unrecoverable method, without a hold mechanism in place.

## PARAMETERS

### -Item
Specifies the item(s) which to delete.

```yaml
Type: Item[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeleteMode
Specifies the delete mode:

MoveToDeletedItems: moves the item to the Deleted Items folder
SoftDelete: deletes the item and places it in the Recoverable Items folder
HardDelete: permanently deletes the item, unless a hold mechanism is in place.

```yaml
Type: DeleteMode
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
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

### System.Void

## NOTES

## RELATED LINKS
