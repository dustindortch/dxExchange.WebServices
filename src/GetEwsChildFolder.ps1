Function Get-EwsChildFolder {
    [CmdletBinding()]
    [OutputType([Microsoft.Exchange.WebServices.Data.Folder])]
    Param(
        [ValidateNotNullOrEmpty()]
        $Folder,

        [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView = 1000,

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    $Service.FindFolders($Folder,$FolderView)
}