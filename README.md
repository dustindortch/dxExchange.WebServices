TOPIC
    about_dxexchange.webservices

SHORT DESCRIPTION
    Depends on the EWS Managed API 2.2 and connects to EWS via AutoDiscover.

LONG DESCRIPTION
    Module supports connecting to EWS for Exchange Server 2007, or newer, and
    Exchange Online (without Modern Authentication).  Requires functional
    AutoDiscover.

EXAMPLES
    PS C:\> Import-Module dxExchange.WebServices
    PS C:\> $PSCredential = Get-Credential
    PS C:\> $ExchangeService = Connect-Ews -Credential $PSCredential

    KeepAlive      Url                                                              ImpersonatedUserId
    ---------      ---                                                              ------------------
    True           https://outlook.office365.com/EWS/Exchange.asmx

    PS C:\> $MailAddress = user@contoso.com
    PS C:\> Set-EwsImpersonationMailbox -EmailAddress $MailAddress

    KeepAlive      Url                                                              ImpersonatedUserId
    ---------      ---                                                              ------------------
    True           https://outlook.office365.com/EWS/Exchange.asmx                  user@contoso.com

    PS C:\> $Folder = Get-EwsWellKnownFolder -FolderName Inbox -Service $ExchangeService
    PS C:\> $Folder

    FolderClass    DisplayName                      TotalCount   UnreadCount
    -----------    -----------                      ----------   -----------
    IPF.Note       Inbox                            76496        60099

    PS C:\> $Items = Get-EwsFolderItem -Folder $Folder -Service $ExchangeService -First 2
    PS C:\> $Items

    From                                    Sent                 Subject
    ----                                    ------------         -------
    Service <SMTP:service@tailspintoys.com> 8/16/2018 8:49:09 PM Unimportant message from the Service
    BigBoss <SMTP:boss@contoso.com>         8/16/2018 8:47:19 PM Urgent message from the boss

NOTE
    The UserPrincipalName is used as the AutoDiscover lookup value if
    -EmailAddress is not specified.  If the credentials are in the format of
    <NTDomain><SamAccountName>, specifiy the EmailAddress.

TROUBLESHOOTING NOTE
    Ensure that the proper version is specified.  The default is
    Exchange2013_SP1, which works with Exchange Server 2013 SP1, or newer, and
    Exchange Online.

SEE ALSO
    Exchange WebServices

KEYWORDS
    - EWS