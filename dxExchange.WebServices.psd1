@{
    RootModule = 'dxExchange.WebServices.psm1'
    ModuleVersion = '0.2'
    Author = 'Dustin Dortch'
    Description = 'Exchange Web Services PowerShell Module'
    PowerShellVersion = '5.0'
    PowerShellHostVersion = '1.0'
    DotNetFrameworkVersion = '4.5.0.0'
    FunctionsToExport = "*"
    CmdletsToExport = "*"
    VariablesToExport = "*"
    AliasesToExport = "*"
    ModuleList = @("dxExchange.WebServices")
    DefaultCommandPrefix = ''
    FileList = @("dxExchange.WebServices.psd1","dxExchange.WebServices.psm1")
    PrivateData = @{
        PSData = @{
            Tags = @('EWS')
        }
    }
}
