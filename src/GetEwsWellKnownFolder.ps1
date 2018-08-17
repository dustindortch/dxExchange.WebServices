Function Get-EwsWellKnownFolder {
    [CmdletBinding()]
    [OutputType([Microsoft.Exchange.WebServices.Data.Folder])]
    Param(
        [Parameter(Position=0)]
        [Alias('Folder')]
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$FolderName = 'Inbox',

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderName)
}