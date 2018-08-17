$EWS = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

If(Test-Path $EWS) {
    $Null = [Reflection.Assembly]::LoadFile($EWS)
} Else {
    Throw "Required EWS Managed API 2.2 not found at $EWS.  Available at https://www.microsoft.com/en-us/download/details.aspx?id=42951"
}

Function Connect-Ews {
    [CmdletBinding(DefaultParameterSetName='UserPrincipalName')]
    [OutputType([Microsoft.Exchange.WebServices.Data.ExchangeService])]
    Param
    (
        [Parameter(Position=0,Mandatory,ParameterSetName='UserPrincipalName')]
        [Alias('Username','User')]
        [String]$UserPrincipalName,

        [Parameter(Mandatory, ParameterSetName='UserPrincipalName')]
        [SecureString]$Password,

        [Parameter(Mandatory, ParameterSetName='Credential')]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(ParameterSetName='UserPrincipalName')]
        [Parameter(ParameterSetName='Credential')]
        [Alias('V')]
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]$Version = 'Exchange2013_SP1',

        [Parameter(ParameterSetName='UserPrincipalName')]
        [Parameter(ParameterSetName='Credential')]
        [MailAddress]$EmailAddress
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
        If($Url -eq 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml') {$True} Else {$False}
    }

    $Credentials = New-Object System.Net.NetworkCredential($UserPrincipalName,$PlaintextPwd)
    $Script:Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Version)
    $Script:Service.Credentials = $Credentials
    Try {
        $Service.AutodiscoverUrl($EmailAddress, $TestUrlCallback)
    } Catch {
        Write-Error "Error: Autodiscover failed for recipient: $EmailAddress"
    }

    $Script:Service
}