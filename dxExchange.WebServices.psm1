[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null

Function Connect-Ews {
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

Function Set-EwsImpersonationMailbox {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$EmailAddress
    )

    $Script:Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$EmailAddress)
}

Function Get-EwsWellKnownFolder {
    [CmdletBinding()]
    Param(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$FolderName = "Inbox"
    )

    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:Service,$FolderName)
    $Folder
}

Function Get-EwsFolder {
    [CmdletBinding()]
    Param(
        $Path
    )

    #Incomplete function
    $MailboxRoot = Get-EwsWellKnownFolder -FolderName MsgFolderRoot
    $PathArray = $Path.Split("\")

    For($i = 1; $i -lt $PathArray.Length; $i++) {
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $Search = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$Path[$i]
        )
    }
}

Function Get-EwsFolderItems {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Folder,
        [Parameter(Mandatory=$False)]
        $Id,
        [System.DateTime]$StartDate,
        [System.DateTime]$EndDate,
        [int]$ResultSize = 10,
        [int]$Skip
    )

    If(!$Skip){
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($ResultSize)
    } Else {
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($ResultSize,$Skip,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
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

    Switch ($Folder.FolderClass) {
        "IPF.Appointment" {
            ForEach($Item in $Items) {
                $ItemId = $Item.Id.UniqueId
                $ItemId
                [Microsoft.Exchange.WebServices.Data.Appointment]::Bind((Get-EwsServiceConnection),$ItemId)
            }
        }
        Default {$Items}
    }
}

Function Get-EwsServiceConnection {
    $Script:Service
}

Function Send-EwsMessage {
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

Function Remove-EwsItem {
    Param(
        $Item,
        [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteType
    )
    
    $Item.Delete($DeleteType)
}

Function Move-EwsItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Item,
        [Parameter(Mandatory)]
        $TargetFolder,
        [ValidateSet("Move","Copy")]
        [String]$MoveType = "Move"
    )

    Switch($Move) {
        "Move" {$Item.Move($TargetFolder)}
        "Copy" {$Item.Copy($TargetFolder)}
    }
}
