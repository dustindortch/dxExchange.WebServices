Function Send-EwsMessage {
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param(
        [Parameter(Mandatory)]
        [MailAddress[]]$To,

        [Parameter(Mandatory)]
        [String]$Subject,

        [Parameter(Mandatory)]
        [String]$Body,

        [Microsoft.Exchange.WebServices.Data.BodyType]$BodyType = 'Text',

        [ValidateNotNullOrEmpty()]
        $ItemAttachment,

        [Switch]$DoNotKeep,

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $Service
    $Message.Subject = $Subject

    ForEach($Recipient in $To) {
        $Null = $Message.ToRecipients.Add($Recipient)
    }

    $Message.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody
    $Message.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
    $Message.Body.Text = $Body

    If($ItemAttachment) {
        $Message.Attachments.AddItemAttachment($ItemAttachment)
    }

    If(-Not $DoNotKeep) {
        $Message.SendAndSaveCopy((Get-EwsWellKnownFolder -FolderName SentItems -Service $Service).Id)
    }
}