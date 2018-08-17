Function Move-EwsItem {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [Microsoft.Exchange.WebServices.Data.Item[]]$Item,

        [Parameter(Mandatory)]
        [Microsoft.Exchange.WebServices.Data.Folder]$TargetFolder,

        [ValidateSet('Move','Copy')]
        [String]$MoveType = 'Move'
    )

    BEGIN {
        $YesToAll = $False
        $NoToAll = $False
    }

    PROCESS {
        ForEach($Instance in $Item) {
            If($PSCmdlet.ShouldProcess($Instance)) {
                If(
                    $Force -or
                    $PSCmdlet.ShouldContinue(
                        $Instance,
                        "Are you sure that you want to $($MoveType.ToLower()) the item?",
                        [Ref]$YesToAll,
                        [Ref]$NoToAll
                    )
                ) {
                    Switch($MoveType) {
                        'Move' {$Null = $Instance.Move($TargetFolder.Id)}
                        'Copy' {$Null = $Instance.Copy($TargetFolder.Id)}
                    }
                }
            }
        }
    }
}