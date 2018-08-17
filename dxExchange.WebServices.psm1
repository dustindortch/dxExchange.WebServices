# Implement your module commands in this script.

$Scripts = Get-ChildItem $PSScriptRoot\src\*.ps1
ForEach($Script in $Scripts) {
    . $Script.VersionInfo.FileName
}


# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*
