---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Move-EwsItem

## SYNOPSIS
Moves or copies an item to the specified target folder.

## SYNTAX

```
Move-EwsItem [-Item] <Item[]> [-TargetFolder] <Folder> [[-MoveType] <String>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Moves or copies an item to the specified target folder, with a default action of move.

## EXAMPLES

### Example 1
```
PS C:\> Move-EwsItem -Item $Item -TargetFolder $Folder
```
```
Are you sure that you want to move the item?
Microsoft.Exchange.WebServices.Data.EmailMessage
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):
```

Moves item $Item to target folder $Folder

### Example 2
```
PS C:\> Move-EwsItem -Item $Item -TargetFolder $Folder -MoveType Copy
```
```
Are you sure that you want to copy the item?
Microsoft.Exchange.WebServices.Data.EmailMessage
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):
```

Copies item $Item to target folder $Folder

## PARAMETERS

### -Item
Specifies the item(s) which to move/copy.

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

### -MoveType
Specifies whether to move or copy the item(s).

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Move, Copy

Required: False
Position: 2
Default value: Move
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetFolder
Specifies the target folder of type **Folder**.

```yaml
Type: Folder
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
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
