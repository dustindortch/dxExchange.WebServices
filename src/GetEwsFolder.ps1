Function Get-EwsFolder {
    [CmdletBinding()]
    [OutputType([Microsoft.Exchange.WebServices.Data.Folder])]
    Param(
        [ValidateNotNullOrEmpty()]
        $FolderPath = '\',

        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$WellKnownFolder = 'MsgFolderRoot',

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    $Parent = $WellKnownFolder
    $View = 1000

    ForEach($Node in $FolderPath.Split('\\')) {
        If($Node -ne '') {
            Write-Verbose "PARENT: $Parent"
            Write-Verbose "NODE: $Node"
            $SubFolders = Get-EwsChildFolder -Folder $Parent -FolderView $View
            $Result = $SubFolders | Where-Object {$_.DisplayName -eq $Node}
            If(($Result.ChildFolderCount | Measure-Object).Count -gt 0) {
                Write-Verbose "ID: $($Result.Id.UniqueId)"
                Write-Verbose "SUBFOLDER COUNT: $($Result.ChildFolderCount)"
                If($Result.ChildFolderCount -ne 0) {
                    $View = New-Object Microsoft.Exchange.WebServices.Data.FolderView($Result.ChildFolderCount)
                }
                $Parent = New-Object Microsoft.Exchange.WebServices.Data.FolderId($Result.Id.UniqueId)
            } Else {
                Throw "Exception: ExFolderDoesNotExistException '$Node' does not exist"
            }
        }
    }
    [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$Parent)
}