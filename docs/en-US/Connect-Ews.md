---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version: https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.exchangeservice?view=exchange-ews-api
schema: 2.0.0
---

# Connect-Ews

## SYNOPSIS
Creates an Exchange Web Services connection using the EWS API 2.2

## SYNTAX

### UserPrincipalName (Default)
```
Connect-Ews [-UserPrincipalName] <String> -Password <SecureString> [-Version <ExchangeVersion>]
 [<CommonParameters>]
```

### UserPrincipalName+Endpoint
```
Connect-Ews [-UserPrincipalName] <String> -Password <SecureString> [-Version <ExchangeVersion>] -Endpoint <Uri>
 [<CommonParameters>]
```

### UserPrincipalName+Autodiscover
```
Connect-Ews [-UserPrincipalName] <String> -Password <SecureString> [-Version <ExchangeVersion>]
 -EmailAddress <MailAddress> [<CommonParameters>]
```

### Credential+Endpoint
```
Connect-Ews -Credential <PSCredential> [-Version <ExchangeVersion>] -Endpoint <Uri> [<CommonParameters>]
```

### Credential+Autodiscover
```
Connect-Ews -Credential <PSCredential> [-Version <ExchangeVersion>] -EmailAddress <MailAddress>
 [<CommonParameters>]
```

### Credential
```
Connect-Ews -Credential <PSCredential> [-Version <ExchangeVersion>] [<CommonParameters>]
```

## DESCRIPTION
Creates an Exchange Web Services connction using the EWS API 2.2.
The UserPrincipalName/UserName is assumed to be the email address unless the EmailAddress parameter is specified for the AutoDiscover lookup.  Alternatively, an EWS endpoint can be supplied to entirely bypass the Autodiscover process (which is significantly faster).

## EXAMPLES

### Example 1
```
PS C:\> Connect-Ews -Credential $PSCredential

KeepAlive      Url                                                              ImpersonatedUserId
---------      ---                                                              ------------------
True           https://outlook.office365.com/EWS/Exchange.asmx
```

Creates an Exchange Web Services connection using the supplied **PSCredential**.
The PSCredential can be created like $PSCredential = Get-Credential and entering the credentials in the prompt.

### Example 2
```
PS C:\> Connect-Ews -Credential $PSCredential -Endpoint https://mail.contoso.com/EWS/Exchange.asmx

KeepAlive      Url                                                              ImpersonatedUserId
---------      ---                                                              ------------------
True           https://mail.contoso.com/EWS/Exchange.asmx
```

Creates and Exchange Web Services connection using the supplied **PSCredential** and **Uri** for the endpoint.

## PARAMETERS

### -Credential
Specifies the credential to use to connect to EWS supplied as a **PSCredential**.

```yaml
Type: PSCredential
Parameter Sets: Credential+Endpoint, Credential+Autodiscover, Credential
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailAddress
Specifies the email address to use for AutoDiscover lookups supplied as a **MailAddress**.

```yaml
Type: MailAddress
Parameter Sets: UserPrincipalName+Autodiscover, Credential+Autodiscover
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Specifies a password supplied as a **SecureString**.

```yaml
Type: SecureString
Parameter Sets: UserPrincipalName, UserPrincipalName+Endpoint, UserPrincipalName+Autodiscover
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UserPrincipalName
Specifies the username, commonly in the format of UserPrincipalName, but can also be in the Netbios format <Domain\>\\<SamAccountName\> (for Exchange Server on-premises, only).

```yaml
Type: String
Parameter Sets: UserPrincipalName, UserPrincipalName+Endpoint, UserPrincipalName+Autodiscover
Aliases: Username, User

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Version
The version of Exchange server to which to connect.
Default is Exchange2013_SP1 and works with Exchange Online and newer versions of Exchange server.

```yaml
Type: ExchangeVersion
Parameter Sets: (All)
Aliases: V
Accepted values: Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013, Exchange2013_SP1

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Endpoint
Specifies the EWS Endpoint to use in order to bypass the Autodiscover process as an **Uri**.

```yaml
Type: Uri
Parameter Sets: UserPrincipalName+Endpoint, Credential+Endpoint
Aliases:

Required: True
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

### Microsoft.Exchange.WebServices.Data.ExchangeService

## NOTES

## RELATED LINKS

[Microsoft.Exchange.WebServices.Data.ExchangeService](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.exchangeservice?view=exchange-ews-api)

