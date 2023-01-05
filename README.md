[SOURCE_CODE]: ./PSXLDevTools/
[TESTS]: ./tests/
[DOCS]: ./docs/
[DOCS_SITE]: https://docs.jimbrig.com/PSXLDevTools/
[BIN]: ./bin/
[CHANGELOG]: ./CHANGELOG.md

<span>
    <h1 align="left">
        <img src="./resources/images/powershellcore.png" height="10%" width="10%" align=left />
PSXLDevTools: PowerShell Excel Developer Tools Module
    </h1>
</span>
<img src="./resources/images/excel.ico" align=right />
<br>

*PowerShell Core Module Containing Various Utility and Helpers for Advanced Office/Excel-based Developers.*

---

<!-- Badges:Begin -->

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSClearHost?color=0092ff&label=PowerShell%20Gallery&logoColor=0092ff)](https://www.powershellgallery.com/packages/PSClearHost/1.0.0)


[![Test Module](https://github.com/jimbrig/PSClearHost/actions/workflows/test.yml/badge.svg)](https://github.com/jimbrig/PSClearHost/actions/workflows/test.yml)
[![Build Module](https://github.com/jimbrig/PSClearHost/actions/workflows/build.yml/badge.svg?branch=develop)](https://github.com/jimbrig/PSClearHost/actions/workflows/build.yml)
[![Publish Module](https://github.com/jimbrig/PSClearHost/actions/workflows/publish.yml/badge.svg)](https://github.com/jimbrig/PSClearHost/actions/workflows/publish.yml)
[![Publish Documentation](https://github.com/jimbrig/PSClearHost/actions/workflows/mkdocs.yml/badge.svg)](https://github.com/jimbrig/PSClearHost/actions/workflows/mkdocs.yml)
[![Automate Changelog](https://github.com/jimbrig/PSClearHost/actions/workflows/changelog.yml/badge.svg)](https://github.com/jimbrig/PSClearHost/actions/workflows/changelog.yml)

<!-- Badges:End -->

---

*View the repo's [CHANGELOG] for details on the progression of the codebase over time.*

## Links

- [Source Code][SOURCE_CODE]
- [Published Documentation](https://docs.jimbrig.com/PSXLDevTools/)
- [Latest Release: *Unreleased*](https://github.com/jimbrig/PSXLDevTools/releases/tag/v0.0.0.9999)
- [Published Module (PowerShell Gallery)](https://www.powershellgallery.com/packages/PSClearHost/1.0.0)


## Overview

`#TODO`

## Installation

```powershell
# Install from the PowerShell Gallery
Install-Module -Name PSXLDevTools -Scope CurrentUser -Force

# Import the module
Import-Module -Name PSXLDevTools
```

## Roadmap

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

---


`#TODO`


---

## Appendices

### Contributing

`#TODO`

### License

[MIT](./LICENSE)

### Credits

`#TODO`

### Changelog

- [CHANGELOG]

