. $PSScriptRoot\..\src\ConnectEws.ps1
. $PSScriptRoot\..\src\GetEwsWellKnownFolder.ps1
. $PSScriptRoot\..\src\GetEwsFolderItem.ps1
$PSCredential = Import-Clixml "$PSScriptRoot\..\.private\Credentials.clixml"

Describe 'Get-EwsFolderItem' {
    Context 'for current mailbox' {
        BeforeAll {
            $Service = Connect-Ews -Credential $PSCredential
            $Folder = Get-EwsWellKnownFolder -FolderName Inbox -Service $Service
        }

        It 'for Inbox folder' {
            $Items = Get-EwsFolderItem -FolderName $Folder -ResultSize 1 -Service $Service
            ($Items | Measure-Object).Count | Should Be 1
        }
    }
}