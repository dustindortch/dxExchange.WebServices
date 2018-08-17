. $PSScriptRoot\..\src\ConnectEws.ps1
. $PSScriptRoot\..\src\SetEwsImpersonationMailbox.ps1
$PSCredential = Import-Clixml "$PSScriptRoot\..\.private\Credentials.clixml"
$EmailAddress = Get-Content "$PSScriptRoot\..\.private\EmailAddress.txt"

Describe 'Set-EwsImpersonationMailbox' {
    Context 'when given an Email Address' {
        BeforeAll {
            $Service = Connect-Ews -Credential $PSCredential -EmailAddress $EmailAddress
        }

        It 'should be not be set' {
            $Service.ImpersonatedUserId | Should BeNullOrEmpty
        }

        It 'with an Email Address' {
            Set-EwsImpersonationMailbox -EmailAddress $EmailAddress -Service $Service
            $Service.ImpersonatedUserId.Id | Should Be $EmailAddress
        }
    }
}