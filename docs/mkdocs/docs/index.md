
<p style="vertical-alignment:middle">
    <img src="./resources/images/powershellcore-preview.ico" height="8%" width="8%" align="left">
    <img src="./resources/images/excel.ico" align="right" style="float:right">
    <h1 align="center">PSXLDevTools</h1>
</p>
<br>

<p align="center">
    <b>PowerShell Excel Developer Tools Module</b><br>
    <em>PowerShell Core Module for Advanced Office/Excel-based Developers.</em><br>
    <br><b>Links:</b><br>
    <a href="https://github.com/jimbrig/PSXLDevTools">Source Code</a> |
    <a href="https://docs.jimbrig.com/PSXLDevTools/">Documentation</a> |
    <a href="https://github.com/jimbrig/PSXLDevTools/releases/tag/v0.0.1">Latest Release: [v0.0.1]</a> |
    <a href="https://www.powershellgallery.com/packages/PSXLDevTools/0.0.1">Published Module</a>
    <br><br>
    <em>View the repo's <a href="./About/CHANGELOG">CHANGELOG</a> for details on the progression of the codebase over time.</em>
    <br><br>
</p>

<span align="center">
<center>
<!-- Badges:Begin -->

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSClearHost?color=0092ff&label=PowerShell%20Gallery&logoColor=0092ff)](https://www.powershellgallery.com/packages/PSClearHost/1.0.0)


[![Test Module](https://github.com/jimbrig/PSXLDevTools/actions/workflows/test.yml/badge.svg)](https://github.com/jimbrig/PSXLDevTools/actions/workflows/test.yml)
[![Build Module](https://github.com/jimbrig/PSXLDevTools/actions/workflows/build.yml/badge.svg)](https://github.com/jimbrig/PSXLDevTools/actions/workflows/build.yml)
[![Publish Module](https://github.com/jimbrig/PSXLDevTools/actions/workflows/publish.yml/badge.svg)](https://github.com/jimbrig/PSXLDevTools/actions/workflows/publish.yml)

[![Publish Documentation](https://github.com/jimbrig/PSXLDevTools/actions/workflows/mkdocs.yml/badge.svg)](https://github.com/jimbrig/PSXLDevTools/actions/workflows/mkdocs.yml)
[![Automate Changelog](https://github.com/jimbrig/PSXLDevTools/actions/workflows/changelog.yml/badge.svg)](https://github.com/jimbrig/PSXLDevTools/actions/workflows/changelog.yml)

<!-- Badges:End -->
</center>
</span>

## Installation

> **Note** View my other PowerShell creations from my [PowerShell Gallery Packages Profile](https://www.powershellgallery.com/profiles/jimbrig)!
    
The module `PSXLDevTools` is published on the [PowerShell Gallery](https://powershellgallery.com/PSXLDevTools/) and can be installed via `PowerShellGet` using the command(s) below:

```powershell
# Install from the PowerShell Gallery
Install-Module -Name PSXLDevTools -Scope CurrentUser -Force

# Import the module
Import-Module -Name PSXLDevTools
```

## Overview

`#TODO`

## Repository

<details>
<summary>Click to Expand Repository File Structure Diagram</summary>

```powershell
> tree /F
<root>
│
├───bin
│       Install-RequiredModules.ps1
│       Invoke-PesterStub.ps1
│       Update-ReadMeIndex.ps1
│
├───docs
│   └───en-US
│           about_PSXLDevTools.help.md
│
├───PSXLDevTools
│   │   PSXLDevTools.psd1
│   │   PSXLDevTools.psm1
│   │
│   ├───Dev
│   │   │   Invoke-XLBuild.ps1
│   │   │   New-VBAProject.ps1
│   │   │   New-VBAProjectConfig.ps1
│   │   │
│   │   ├───Exports
│   │   │       Export-CustomCellStyles.ps1
│   │   │       Export-DataMashup.ps1
│   │   │       Export-ListObject.ps1
│   │   │       Export-ListObjects.ps1
│   │   │       Export-PowerQuery.ps1
│   │   │       Export-PowerQueryConnection.ps1
│   │   │       Export-TableStyles.ps1
│   │   │       Export-VBAComponent.ps1
│   │   │       Export-VBAProjectProps.ps1
│   │   │       Export-VBAReferences.ps1
│   │   │       Export-WorksheetMetadata.ps1
│   │   │       Export-XLConditionalFormatting.ps1
│   │   │       Export-XLCustomLists.ps1
│   │   │       Export-XLCustomRibbonX.ps1
│   │   │       Export-XLDataModel.ps1
│   │   │       Export-XLDataValidation.ps1
│   │   │       Export-XLDocumentProps.ps1
│   │   │       Export-XLPivotCache.ps1
│   │   │       Export-XLPivotTable.ps1
│   │   │       Export-XLTheme.ps1
│   │   │       Export-XLThemeColors.ps1
│   │   │       Export-XLThemeFonts.ps1
│   │   │
│   │   └───Imports
│   │           Import-DataMashup.ps1
│   │
│   ├───Private
│   │       GetHelloWorld.ps1
│   │
│   └───Public
│           Export-PowerQueries.ps1
│           Get-HelloWorld.ps1
│
├───resources
│   │   dirtree.js
│   │   md.config.js
│   │
│   └───images
│           excel.ico
│           office365.ico
│           powershell.ico
│           powershellcore-preview.ico
│           powershellcore.png
│           regedit.ico
│           win10.ico
│           windowspowershell.ico
│
├───tests
│   │   Export-PowerQueries.tests.ps1
│   │   Help.tests.ps1
│   │   Manifest.tests.ps1
│   │   Meta.tests.ps1
│   │   MetaFixers.psm1
│   │   ScriptAnalyzerSettings.psd1
│   │
│   └───TestWorkbooks
│
│   .editorconfig
│   .gitattributes
│   .gitignore
│   build.ps1
│   CHANGELOG.md
│   cliff.toml
│   LICENSE
│   mkdocs.yml
│   psakeFile.ps1
│   README.md
│   RequiredModules.psd1
│   requirements.psd1
│
├───.devcontainer
│       devcontainer.json
│       Dockerfile
│
├───.github
│   │   CONTRIBUTING.md
│   │   ISSUE_TEMPLATE.md
│   │   PULL_REQUEST_TEMPLATE.md
│   │
│   └───workflows
│           build.yml
│           changelog.yml
│           lint.yml
│           mkdocs.yml
│           publish.yml
│           readme.yml
│           test.yml
│
├───.vscode
│       extensions.json
│       launch.json
│       settings.json
│       tasks.json
```
</details>

***

## Appendices

### Contributing

`#TODO`

### License

[LICENSE](https://github/com/jimbrig/PSXLDevTools/blob/main/LICENSE)

### Credits

`#TODO`

### Changelog

[Changelog](About/CHANGELOG)

