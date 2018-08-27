---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Get-EwsFolderItem

## SYNOPSIS
Gets items from specified Folder.

## SYNTAX

### ResultSize (Default)
```
Get-EwsFolderItem [-Folder] <Folder> [-StartDate <DateTime>] [-EndDate <DateTime>] [-ResultSize <Int32>]
 [-Skip <Int32>] [-NoDetails] [-Service <ExchangeService>] [<CommonParameters>]
```

### Last
```
Get-EwsFolderItem [-Folder <Folder>] [-StartDate <DateTime>] [-EndDate <DateTime>] [-Last <Int32>]
 [-Skip <Int32>] [-NoDetails] [-Service <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Getss items from the specified folder and allows for various filtering criteria, including start and end dates, first or last number of items, skip items for paging, and exclude item details (like the item body) to improve the speed.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-EwsFolderItem -Folder $Folder -First 2
```
```
From                                    Sent                 Subject
----                                    ------------         -------
Service <SMTP:service@tailspintoys.com> 8/16/2018 8:49:09 PM Unimportant message from the Service
BigBoss <SMTP:boss@contoso.com>         8/16/2018 8:47:19 PM Urgent message from the boss
```

Gets the first two messages from the specified folder.  The **Folder** can be specific with the **Get-EwsWellKnownFolder** cmdlet.

## PARAMETERS

### -EndDate
Used to limit criteria to items before specified date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Folder
Specifies the **Microsoft.Exchange.WebServices.Data.Folder** to search.  Can be set with the **Get-EwsWellKnownFolder** cmdlet.

```yaml
Type: Microsoft.Exchange.WebServices.Data.Folder
Parameter Sets: (All)
Aliases: FolderName

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Last
Specifies to get the last N items from the folder.

```yaml
Type: Int32
Parameter Sets: Last
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoDetails
Specifies to not get the details of the item, such as the **Body**, but at a significant speed improvement.  The Id properties can then be passed to the **Get-EwsItem** cmdlet to get the details.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResultSize
Limits results to the first N items.  The default value is 10, but up to 1000 items can be retrieved at once.

```yaml
Type: Int32
Parameter Sets: ResultSize
Aliases: First

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Optionally, specifies an ExchangeService object.

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

### -Skip
Specifies N items to skip, which can be used to page through additional items on subsequent executions.

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

### -StartDate
Used to limit criteria to items after specified date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

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

### Microsoft.Exchange.WebServices.Data.Item


## NOTES

## RELATED LINKS

[Microsoft.Exchange.WebServices.Data.Item](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.item?view=exchange-ews-api)