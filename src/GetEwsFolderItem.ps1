Function Get-EwsFolderItem {
    [CmdletBinding(DefaultParameterSetName='ResultSize')]
    [OutputType([Microsoft.Exchange.WebServices.Data.Item])]
    Param(
        [Parameter(Position=0,Mandatory)]
        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [Alias('FolderName')]
        [ValidateScript({$_ -as [Microsoft.Exchange.WebServices.Data.Folder]})]
        [PSObject]$Folder,

        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [System.DateTime]$StartDate,

        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [System.DateTime]$EndDate,

        [Parameter(ParameterSetName='ResultSize')]
        [Alias('First')]
        [ValidateRange(1,1000)]
        [int]$ResultSize = 10,

        [Parameter(ParameterSetName='Last')]
        [ValidateRange(1,1000)]
        [int]$Last,

        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [ValidateRange(0,999)]
        [int]$Skip = 0,

        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [switch]$NoDetails,

        [Parameter(ParameterSetName='ResultSize')]
        [Parameter(ParameterSetName='Last')]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:Service
    )

    If ($Last) {
        $Offset = [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::End
    } Else {
        $Offset = [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning
    }

    If($Folder.FolderClass -eq 'IPF.Appointment') {
        If($StartDate) {
            $CalendarStartDate = $StartDate
        } Else {
            $CalendarStartDate = Get-Date
        }
        If($EndDate) {
            $CalendarEndDate = $EndDate
        } Else {
            $CalendarEndDate = (Get-Date).AddMonths(6)
        }
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($CalendarStartDate,$CalendarEndDate,$ResultSize)
    } Else {
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($ResultSize,$Skip,$Offset)
    }

    If(($StartDate -or $EndDate) -and $Folder.FolderClass -ne 'IPF.Appointment') {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        If($StartDate) {
            $AfterFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$StartDate)
            $SearchFilter.Add($AfterFilter)
        }

        If($EndDate) {
            $BeforeFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$EndDate)
            $SearchFilter.Add($BeforeFilter)
        }
        $Items = $Service.FindItems($Folder.Id,$SearchFilter,$ItemView)
    } Else {
        Write-Verbose 'Calendar or no dates'
        $Items = $Service.FindItems($Folder.Id,$ItemView)
    }

    If($NoDetails) {
        Write-Verbose 'ITEM CLASS: NoDetails specified'
        $Items
    } Else {
        ForEach ($Item in $Items) {
            $ItemId = $Item.Id.UniqueId
            Write-Verbose "ITEM CLASS: $($Item.ItemClass)"
            [Microsoft.Exchange.WebServices.Data.Item]::Bind($Service,$ItemId)
        }
    }
}