---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Set-EwsImpersonationMailbox

## SYNOPSIS
Set which mailbox to access via Application Impersonation rights (requires rights already established).

## SYNTAX

```
Set-EwsImpersonationMailbox [-EmailAddress] <MailAddress> [-Service <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Set which mailbox to access via Application Impersonation rights.  Use the <MailAddress> with the **Connect-Ews** cmdlet when the identity associated with the authentication credentials has no assigned mailbox prior to running **Set-EwsImpersonationMailbox**.  Can be used to set the default service or any saved service object.

## EXAMPLES

### Example 1
```powershell
PS C:\> Set-EwsImpersonationMailbox -EmailAddress $MailAddress
```
```
KeepAlive      Url                                                              ImpersonatedUserId
---------      ---                                                              ------------------
True           https://outlook.office365.com/EWS/Exchange.asmx                  user@contoso.com
```

In the default ExchangeService, set the ImperonatedUserId in the EWS Managed API for the service.

### Example 2
```powershell
PS C:\> Set-EwsImpersonationMailbox -EmailAddress $MailAddress -Service $Service
```
```
KeepAlive      Url                                                              ImpersonatedUserId
---------      ---                                                              ------------------
True           https://outlook.office365.com/EWS/Exchange.asmx                  user@contoso.com
```

In the **$Service** ExchangeService object, set the ImperonatedUserId in the EWS Managed API for the service.

## PARAMETERS

### -EmailAddress
Provide a valid recipient address for an Exchange mailbox to which Application Impersonation rights have been granted.

```yaml
Type: MailAddress
Parameter Sets: (All)
Aliases: None

Required: True
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

### Microsoft.Exchange.WebServices.Data.ExchangeService

## NOTES

## RELATED LINKS

[microsoft.exchange.webservices.data.impersonateduserid](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.impersonateduserid?view=exchange-ews-api)