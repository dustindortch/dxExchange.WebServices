$EWS = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

If(Test-Path $EWS) {
    $Null = [Reflection.Assembly]::LoadFile($EWS)
} Else {
    Throw "Required EWS Managed API 2.2 not found at: $EWS.  Available at https://www.microsoft.com/en-us/download/details.aspx?id=42951"
}

Function Connect-Ews {
    [CmdletBinding(DefaultParameterSetName='UserPrincipalName')]
    [OutputType([Microsoft.Exchange.WebServices.Data.ExchangeService])]
    Param(
        [Parameter(Position=0,Mandatory,ParameterSetName='UserPrincipalName')]
        [Parameter(Position=0,Mandatory,ParameterSetName='UserPrincipalName+Autodiscover')]
        [Parameter(Position=0,Mandatory,ParameterSetName='UserPrincipalName+Endpoint')]
        [Alias('Username','User')]
        [String]$UserPrincipalName,

        [Parameter(Mandatory,ParameterSetName='UserPrincipalName')]
        [Parameter(Mandatory,ParameterSetName='UserPrincipalName+Autodiscover')]
        [Parameter(Mandatory,ParameterSetName='UserPrincipalName+Endpoint')]
        [SecureString]$Password,

        [Parameter(Mandatory,ParameterSetName='Credential')]
        [Parameter(Mandatory,ParameterSetName='Credential+Autodiscover')]
        [Parameter(Mandatory,ParameterSetName='Credential+Endpoint')]
        [System.Management.Automation.PSCredential]$Credential,

        [Alias('V')]
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]$Version = 'Exchange2013_SP1',

        [Parameter(Mandatory,ParameterSetName='UserPrincipalName+Autodiscover')]
        [Parameter(Mandatory,ParameterSetName='Credential+Autodiscover')]
        [MailAddress]$EmailAddress,

        [Parameter(Mandatory,ParameterSetName='UserPrincipalName+Endpoint')]
        [Parameter(Mandatory,ParameterSetName='Credential+Endpoint')]
        [Uri]$Endpoint
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

    $Credentials = New-Object System.Net.NetworkCredential($UserPrincipalName,$PlaintextPwd)
    $Script:Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Version)
    $Script:Service.Credentials = $Credentials

    Try {
        If($Endpoint) {
            Write-Verbose "EWS ENDPOINT: $Endpoint"
            $Service.Url = $Endpoint
        } Else {
            $TestUrlCallback = {
                Param([Uri]$Url)
                Write-Verbose "AUTODISCOVER REDIRECT: $Url"
                If($Url -eq 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml') {$True} Else {$False}
            }

            $Service.AutodiscoverUrl($EmailAddress, $TestUrlCallback)
        }
    } Catch {
        Write-Error "Error: Autodiscover failed for recipient: $EmailAddress"
    }

    $Script:Service
}