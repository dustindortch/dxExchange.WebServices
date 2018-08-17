Function Set-EwsImpersonationMailbox {
    [CmdletBinding()]
    [OutputType([Microsoft.Exchange.WebServices.Data.ExchangeService])]
    Param(
        [Parameter(Mandatory)]
        [MailAddress]$EmailAddress,

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$EmailAddress)
    $Service
}