@{
    RootModule = 'xExchange.WebServices.psm1'
    ModuleVersion = '0.1'
    Author = 'Dustin Dortch'
    Description = 'Exchange Web Services PowerShell Module'
    PowerShellVersion = '5.0'
    PowerShellHostVersion = '1.0'
    DotNetFrameworkVersion = '4.5.0.0'
    FunctionsToExport = "*"
    CmdletsToExport = "*"
    VariablesToExport = "*"
    AliasesToExport = "*"
    ModuleList = @("xExchange.WebServices")
    DefaultCommandPrefix = ''
    FileList = @("xExchange.WebServices.psd1","xExchange.WebServices.psm1")
    PrivateData = @{
        PSData = @{
            Tags = @('EWS')
        }
    }
}