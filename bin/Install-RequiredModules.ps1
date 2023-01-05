Push-Location $PSScriptRoot\..\
try {
    Install-Script Install-RequiredModule
    Install-RequiredModule -RequiredModulesFile $PSScriptRoot\RequiredModules.psd1
} finally {
    Pop-Location
}
