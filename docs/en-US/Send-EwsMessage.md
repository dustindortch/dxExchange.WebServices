---
external help file: dxExchange.WebServices-help.xml
Module Name: dxExchange.WebServices
online version:
schema: 2.0.0
---

# Send-EwsMessage

## SYNOPSIS
Sends a message as the signed in account.

## SYNTAX

```
Send-EwsMessage [-To] <MailAddress[]> [-Subject] <String> [-Body] <String> [-BodyType <BodyType>]
 [[-ItemAttachment] <Object>] [-DoNotKeep] [-Service <ExchangeService>] [<CommonParameters>]
```

## DESCRIPTION
Sends a message as the signed in account.

## EXAMPLES

### Example 1
```
PS C:\> Send-EwsMessage -To user@contoso.com -Subject 'EWS Message from Me' -Body 'Message here'
```

Sends a message to user@contoso.com as the authenticated user.

### Example 2
```
PS C:\> Send-EwsMessage -To user@contoso.com -Subject 'EWS Message from Me' -Body 'Message here' -Service $ImpersonatedService
```

Sends a message to user@contoso.com as the Impersonated Account.

### Example 3
```
PS C:\> Send-EwsMessage -To user@contoso.com -Subject 'EWS Message from Me' -Body 'Message here' -DoNotKeep
```

Sends a message to user@contoso.com as the authenticated user and does not keep it in the Sent Items folder.

## PARAMETERS

### -Body
The body of the email message.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ItemAttachment
The file attachment to send.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
The subject of the message.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
The recipient(s) of the message.

```yaml
Type: MailAddress[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyType
Specifies the body type, either Text (default) or HTML.

```yaml
Type: BodyType
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Text
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoNotKeep
Do not place the item in the Sent Items folder.

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

### -Service
Optionally, specifies an ExchangeService object to set the impersonation mailbox.

```yaml
Type: ExchangeService
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

### System.Void

## NOTES

## RELATED LINKS
