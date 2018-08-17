---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Connect-Ews

## SYNOPSIS
Creates an Exchange Web Services connection using the EWS API 2.2

## SYNTAX

### Credential
```
Connect-Ews [-Version <ExchangeVersion>] -Credential <PSCredential> [-EmailAddress <MailAddress>]
 [<CommonParameters>]
```

### UserPrincipalName
```
Connect-Ews [-Version <ExchangeVersion>] [-UserPrincipalName] <String> -Password <SecureString>
 [-EmailAddress <MailAddress>] [<CommonParameters>]
```

### DEFAULT
```
Connect-Ews -Version <ExchangeVersion> [-EmailAddress <MailAddress>] [<CommonParameters>]
```

## DESCRIPTION
Creates an Exchange Web Services connction using the EWS API 2.2 via AutoDiscover.  The UserPrincipalName/UserName is assumed to be the email address unless the EmailAddress parameter is specified for the AutoDiscover lookup.

## EXAMPLES

### Example 1
```powershell
PS C:\> Connect-Ews -Credential $PSCredential
```
```
KeepAlive      Url                                                              ImpersonatedUserId
---------      ---                                                              ------------------
True           https://outlook.office365.com/EWS/Exchange.asmx
```

Creates an Exchange Web Services connection using the supplied PSCredential.  The PSCredential can be created like **$PSCredential = Get-Credential** and entering the credentials in the prompt.

## PARAMETERS

### -Credential
Specifies the credential to use to connect to EWS.  To obtain a **PSCredential** object, use the **Get-Credential** cmdlet.

```yaml
Type: PSCredential
Parameter Sets: Credential
Aliases: None

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailAddress
Specifies the email address to use for AutoDiscover lookups.

```yaml
Type: MailAddress
Parameter Sets: (All)
Aliases: None

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Specifies a password supplied as a **SecureString**.

```yaml
Type: SecureString
Parameter Sets: UserPrincipalName
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UserPrincipalName
Specifies the username, commonly in the format of UserPrincipalName, but can also be in the Netbios format <Domain>\<SamAccountName> (for Exhange Server on-premises, only).

```yaml
Type: String
Parameter Sets: UserPrincipalName
Aliases: Username, User

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Version
The version of Exchange server to which to connect.  Default is Exchange2013_SP1 and works with Exchange Online and newer versions of Exchange server.

```yaml
Type: ExchangeVersion
Parameter Sets: Credential, UserPrincipalName
Aliases: V
Accepted values: Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013, Exchange2013_SP1

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

### Microsoft.Exchange.WebServices.Data.ExchangeService


## NOTES

## RELATED LINKS

[Microsoft.Exchange.WebServices.Data.ExchangeService](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.exchangeservice?view=exchange-ews-api)