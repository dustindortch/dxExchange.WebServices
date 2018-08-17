Function Remove-EwsItem {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [Microsoft.Exchange.WebServices.Data.Item[]]$Item,

        [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteMode = 'MoveToDeletedItems'
    )

    BEGIN {
        $YesToAll = $False
        $NoToAll = $False
    }

    PROCESS {
        ForEach($Instance in $Item) {
            If($PSCmdlet.ShouldProcess($Instance)) {
                If(
                    ($DeleteMode -eq 'MoveToDeletedItems' -and $ConfirmPreference -eq 'High') -or
                    $Force -or
                    $PSCmdlet.ShouldContinue(
                        $Instance,
                        'Are you sure that you want to delete the item?',
                        [Ref]$YesToAll,
                        [Ref]$NoToAll
                    )
                ) {
                    $Instance.Delete($DeleteMode)
                }
            }
        }
    }
}