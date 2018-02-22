[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null

Function Connect-Ews
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,ParameterSetName="DEFAULT")]
        [Parameter(ParameterSetName="UserPrincipalName")]
        [Parameter(ParameterSetName="Credential")]
        [Alias("V")]
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]$Version = "Exchange2013_SP1",

        [Parameter(Position=0,Mandatory,ParameterSetName="UserPrincipalName")]
        [Alias("User")]
        [String]$UserPrincipalName,
        
        [Parameter(Mandatory, ParameterSetName="UserPrincipalName")]
        [SecureString]$Password,

        [Parameter(Mandatory, ParameterSetName="Credential")]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory=$False)]
        [Parameter(ParameterSetName="UserPrincipalName")]
        [Parameter(ParameterSetName="Credential")]
        [String]$EmailAddress
    )

    If($Credential) {
        $UserPrincipalName = $Credential.UserName
        $PlaintextPwd = $Credential.GetNetworkCredential().Password
    }

    If($UserPrincipalName -and $Password) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        [String]$PlaintextPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }

    If(!($EmailAddress)) {
        $EmailAddress = $UserPrincipalName
    }

    $TestUrlCallback = {
        param([String] $Url)
        Write-Verbose "AUTODISCOVER REDIRECT: $Url"
        If($Url -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml") {$True} Else {$False}
    }

    $Credentials = New-Object System.Net.NetworkCredential($UserPrincipalName,$PlaintextPwd)
    $Script:Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Version)
    $Script:Service.Credentials = $Credentials
    $Script:Service.AutodiscoverUrl($EmailAddress, $TestUrlCallback)
}

Function Get-EwsWellKnownFolder
{
    [CmdletBinding()]
    Param(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$FolderName = "Inbox"
    )

    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:Service,$FolderName)
    $Folder
}

Function Get-EwsFolderItems
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Folder,
        [System.DateTime]$StartDate,
        [System.DateTime]$EndDate,
        [int]$ResultSize = 10,
        [int]$Skip
    )

    If(!$Skip){
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($ResultSize)
    } Else {
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($ResultSize,0,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::$Skip)
    }

    If($StartDate) {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$StartDate)
    }

    If($EndDate) {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$EndDate)
    }

    If($SearchFilter) {
        $Items = $Script:Service.FindItems($Folder.Id,$SearchFilter,$ItemView)
    } Else {
        $Items = $Script:Service.FindItems($Folder.Id,$ItemView)
    }

    $Items
}

Function Get-EwsServiceConnection
{
    $Script:Service
}

Function Send-EwsMessage
{
    Param(
        [Parameter(Mandatory)]
        [String]$To,
        [Parameter(Mandatory)]
        [String]$Subject,
        [Parameter(Mandatory)]
        [String]$Body,
        [Parameter(Mandatory=$False)]
        $ItemAttachment
    )

    $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $Script:Service
    $Message.Subject = $Subject
    $Message.ToRecipients.Add($To)
    $Message.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody
    $Message.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
    $Message.Body.Text = $Body
    If($ItemAttachment) {
        $Message.Attachments.AddItemAttachment($ItemAttachment)
    }
    $Message.SendAndSaveCopy((Get-EwsWellKnownFolder -FolderName SentItems).Id)
}