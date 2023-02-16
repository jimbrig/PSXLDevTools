# About

## PowerShell Gallery

The module `PSClearHost` is published on the [PowerShell Gallery](https://www.powershellgallery.com): <https://powershellgallery.com/PSXLDevTools/>

View my other PowerShell creations from [My PowerShell Gallery Profile](https://www.powershellgallery.com/profiles/jimbrig)!

## Changelog

- See the [Changelog](CHANGELOG.md) for detailed development over time!

## Installation

```powershell
# Install from the PowerShell Gallery
Install-Module -Name PSXLDevTools -Scope CurrentUser -Force

# Import the module
Import-Module -Name PSXLDevTools
```

## Roadmap

- [VBA Related](#vba-related)
    - [Source Code Management](#source-code-management)
    - [VBA Project Development and Engineering](#vba-project-development-and-engineering)
    - [VBA Project Builds](#vba-project-builds)
- [Excel Workbook Development and Engineering](#excel-workbook-development-and-engineering)
    - [Excel Workbook Source Code Management](#excel-workbook-source-code-management)
    - [Excel Workbook Builds](#excel-workbook-builds)
- [Excel Workbook Development and Engineering](#excel-workbook-development-and-engineering)
    - [Office Fluent Ribbon XML](#office-fluent-ribbon-xml)
    - [Workbook Metadata, Styling, Themes, and Custom Properties](#workbook-metadata-styling-themes-and-custom-properties)
    - [Workbook Styling, Themes, and Formats](#workbook-styling-themes-and-formats)
- [Data Engineering and Modeling with PowerQuery and PowerPivot Data Models](#data-engineering-and-modeling-with-powerquery-and-powerpivot-data-models)
    - [PowerQuery Source Code Management (M-Code Formulae and Metadata)](#powerquery-source-code-management-m-code-formulae-and-metadata)
    - [Excel Data Model Integration](#excel-data-model-integration)
- [Developer Tool Integration and Automation](#developer-tool-integration-and-automation)

### VBA Related

#### Source Code Management

- Import/Export VBA Project Code Modules:
    - Standard Modules: `Source/VBA/Modules/mod*.bas`
    - Class Modules: `Source/VBA/Classes/cls*.cls`
    - Class Interface Modules: `Source/VBA/Interfaces/Icls*.cls`
    - Excel Object Class Modules: `Source/VBA/ExcelObjects/<WorksheetCodeName>.docls|ThisWorkbook.docls`
    - User Forms: `Source/VBA/Forms/frm*.frm` (*Optionally can remove any unnecessary `.frx` exported binaries.*)
    - User Form Controls and Properties: `Source/VBA/Forms/FormControlsProperties.txt`

- Import/Export VBA Project Properties: `VBAProjectProps.txt`
- Import/Export VBA Project References / Dependencies: `References.txt`

- Bundle `_DEV.xlsm` development workbooks and separate `_DEV.xlam` addins for preliminary setup and development tasks.

### VBA Project Development and Engineering

- Debug, Compile, and Run Unit Tests on VBA Source Code without Opening Excel.
- Automate template creation for VBA Projects and VBA Project Groups (Standard/Common Library Support).
- Automate Documentation of VBA Project Code Modules.
- Automate Documentation of Dependencies and References for VBA Projects and VBA Project Groups.
- Lint and Format VBA Source Code.
- Report test-coverage on VBA Source Code.
- Report on VBA Source Code Complexity (Cyclomatic, etc.).

### VBA Project Builds

- Create VBA Project Group (Library) Manifest: `<ProjectGroupName>.vbg`
- Create VBA Project Manifest: `<ProjectName>.vbp`

- Bundle, Build, and Compile Excel VBA Projects and in turn, de-compile or extract items from `.rels`, `VBAProject.bin`,
    and other potential helpful internals.

- Curate separate layers for `Debug`, `Dev`, `Test`, and `Prod` environments using conditional compilation directives or
    arguments. This allows for the creation of a single VBA Project that can be used in multiple environments. It also
    allows for the developer to toggle on extra components with features such as debug and trace event logging, developer tools,
    tests, documentation generation, code-review, etc.

- Strip final production builds of all unnecessary code, comments, and other items that are not needed for the final
    production build.

### Excel Workbook Development and Engineering

#### Office Fluent Ribbon XML

- Manage and Import/Export Custom RibbonX Components:
    - Excel Workbook Custom Ribbon XML: `Source/Excel/RibbonX/customUI14.xml|customUI.xml`
    - Ribbon Callback Procedures Skeleton: `Source/Excel/RibbonX/Callbacks.txt`
    - Any Icons or Images used in the Ribbon: `Source/Excel/RibbonX/Images/*[.ico|.png|.jpg]`

#### Workbook Metadata, Styling, Themes, and Custom Properties

- Manage and Import/Export Excel Workbook Metadata:
    - Workbook Custom Document Properties: `Source/Excel/Metadata/DocumentProperties.txt`
    - WorkSheet Metadata (Code Names, Display Names, Tab Colors, etc.): `Source/Excel/Metadata/Worksheets.txt`
    - NamedRanges: `Source/Excel/Metadata/NamedRanges.txt`
    - ListObjects (Tables): `Source/Excel/Metadata/ListObjects.txt`
    - PivotTables: `Source/Excel/Metadata/PivotTables.txt`
    - PivotCaches: `Source/Excel/Metadata/PivotCaches.xml`
    - Custom XML: `Source/Excel/Metadata/CustomXMLMap.xlm`

#### Workbook Styling, Themes, and Formats

- Manage and Import/Export Excel Workbook Design Elements:
    - Custom Themes, Theme Colors, Theme Fonts, and Theme Effects:
        - `Source/Excel/Themes/<ProjectName>.thmx`
        - `Source/Excel/Themes/Theme Colors/<ProjectName>_Colors.xml`
        - `Source/Excel/Themes/Theme Fonts/<ProjectName>_Fonts.xml`
        - `Source/Excel/Themes/Theme Effects/<ProjectName>_Effects.xml`
    - Cell Styles: `Source/Excel/Metadata/CellStyles.xml`
    - Number Formats: `Source/Excel/Metadata/NumberFormats.txt`
    - Conditional Formatting Rules: `Source/Excel/Metadata/ConditionalFormattingRules.txt`
    - TableStyles: `Source/Excel/Metadata/TableStyles.xml`
    - Custom Lists: `Source/Excel/Metadata/CustomLists.xml`
    - Charts: `Source/Excel/Metadata/Charts.txt`

### Data Engineering and Modeling with PowerQuery and PowerPivot Data Models

#### PowerQuery Source Code Management (M-Code Formulae and Metadata)

- Manage and Import/Export PowerQuery M-Code:
    - PowerQuery Queries: `Source/PowerQuery/Queries/*.pq`
    - PowerQuery Query Metadata: `Source/PowerQuery/Queries/*.meta.pq`
    - PowerQuery Query Data Source Dependencies: `Source/PowerQuery/DataSources/*.pqd`
    - Data Mashups: `Source/PowerQuery/Mashups/*.pqm`
    - Custom User Defined Functions and Parameters: `Source/PowerQuery/UDFs/*.pq`, `Source/PowerQuery/Parameters/*.pq`
    - Associated Workbook Connections: `Source/PowerQuery/Connections/*.odc`
    - Associated ListObjects/Query Tables: `Source/PowerQuery/Tables/*.xml`

- Aggregate a library of PowerQuery Formulae and associated metadata into a single file for easy import/export.

- Manage Data Flow by Removing Connections and Queries from the Workbook after used, or as necessary.

#### Excel Data Model Integration

- Manage and Import/Export Excel Workbook Data Models:
    - Data Model Office Data Connection: `Source/PowerQuery/Connections/DataModel.odc`
    - Data Model Metadata: `Source/Excel/Model/DataModel.xml`
    - Data Model Relationships: `Source/Excel/Model/DataModelRelationships.xml`
    - Data Model Tables: `Source/Excel/Model/DataModelTables.xml`

- Import/Export Connection Strings and Parameterized Queries from the Data Model.

- Manage `ADO` and `DAO` based connections.

### Developer Tool Integration and Automation

- Version Control over Excel Workbooks, Data Models, PowerQuery, VBA Projects.
- Office RibbonX Editor
- VSCode + Extensions
- DAXStudio + PowerBI Desktop
- Excel + VBE Customizations + PowerQuery Editor

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
