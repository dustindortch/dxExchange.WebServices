. $PSScriptRoot\..\src\ConnectEws.ps1
. $PSScriptRoot\..\src\GetEwsWellKnownFolder.ps1
$PSCredential = Import-Clixml "$PSScriptRoot\..\.private\Credentials.clixml"

Describe 'Get-EwsWellKnownFolder' {
    Context 'for current mailbox' {
        BeforeAll {
            $Service = Connect-Ews -Credential $PSCredential
        }

        It 'for Inbox folder' {
            $Folder = Get-EwsWellKnownFolder -FolderName Inbox -Service $Service
            $Folder.DisplayName | Should Be 'Inbox'
            $Folder.FolderClass | Should Be 'IPF.Note'
        }

        It 'for Calendar folder' {
            $Folder = Get-EwsWellKnownFolder -FolderName Calendar -Service $Service
            $Folder.DisplayName | Should Be 'Calendar'
            $Folder.FolderClass | Should Be 'IPF.Appointment'
        }

        It 'for Sent Items folder' {
            $Folder = Get-EwsWellKnownFolder -FolderName SentItems -Service $Service
            $Folder.DisplayName | Should Be 'Sent Items'
            $Folder.FolderClass | Should Be 'IPF.Note'
        }
    }
}