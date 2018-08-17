. $PSScriptRoot\..\src\ConnectEws.ps1
$PSCredential = Import-Clixml "$PSScriptRoot\..\.private\Credentials.clixml"
$BadPSCredential = Import-Clixml "$PSScriptRoot\..\.private\BadCredentials.clixml"
$UserPrincipalName = $PSCredential.UserName
$SecureString = $PSCredential.Password
$EmailAddress = Get-Content "$PSScriptRoot\..\.private\EmailAddress.txt"

Describe 'Connect-Ews' {
    Context 'when given a PSCredential' {
        It 'without Email Address' {
            Connect-Ews -Credential $PSCredential
        }

        It 'with Email Address' {
            Connect-Ews -Credential $PSCredential -EmailAddress $EmailAddress
        }

        It 'should fail with a Bad PSCredential' {
            {Connect-Ews -Credential $BadPSCredential -ErrorAction Stop} | Should Throw
        }
    }

    Context 'when given a UserPrincipalName' {
        It 'without Email Address' {
            Connect-Ews -UserPrincipalName $UserPrincipalName -Password $SecureString
        }

        It 'with Email Address' {
            Connect-Ews -UserPrincipalName $UserPrincipalName -Password $SecureString -EmailAddress $EmailAddress
        }

        It 'should fail with a Bad PSCredential' {
            {Connect-Ews -UserPrincipalName $BadPSCredential.UserName -Password $BadPSCredential.Password -EmailAddress $EmailAddress -ErrorAction Stop} | Should Throw
        }
    }
}