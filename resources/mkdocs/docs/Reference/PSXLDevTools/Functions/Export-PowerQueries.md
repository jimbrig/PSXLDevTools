---
external help file: PSXLDevTools-help.xml
Module Name: PSXLDevTools
online version:
schema: 2.0.0
---

# Export-PowerQueries

## SYNOPSIS
Exports Power Queries' M-Code Formulae from an Excel PowerQuery Enabled Workbook to a specified folder.

## SYNTAX

```
Export-PowerQueries [-Path] <String> [[-ExportPath] <String>] [[-Extension] <String>] [-Force]
 [<CommonParameters>]
```

## DESCRIPTION
This function exports Power Queries' M-Code Formulae from an Excel PowerQuery Enabled Workbook to a specified
destination source code folder.
This allows for the M-Code to be version controlled and maintained in a
source code repository alongside the rest of the workbook's source code (VBA, XML, SQL, DAX, etc.).

The function is designed to be used in conjunction with the Import-PowerQueries function, which imports all of
the Power Queries' M-Code Formulae from the specified source code folder into the Excel PowerQuery Enabled Workbook.

## EXAMPLES

### EXAMPLE 1
```
Export-PowerQueries -Path ".\MyWorkbook.xlsx" -ExportPath ".\Source\PowerQuery"
Successfully exported MyQuery to file C:\MyProject\Source\PowerQuery\MyQuery.pq
Successfully exported MyOtherQuery to file C:\MyProject\Source\PowerQuery\MyOtherQuery.pq
```

### EXAMPLE 2
```
Export-PowerQueries -Path .\Test.xlsm -ExportPath .\Source\PQ -Extension .pqm -Force
Successfully exported MyQuery to file C:\MyProject\Source\PQ\MyQuery.pqm
Successfully exported MyOtherQuery to file C:\MyProject\Source\PQ\MyOtherQuery.pqm
```

## PARAMETERS

### -Path
The path to the Excel PowerQuery Enabled Workbook.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExportPath
(Optional) The path to the folder where the Power Queries' M-Code Formulae will be exported to.
If not specified,
\`\<ProjectRoot\>/Source/PowerQuery/*\` is used as the default source code export path for the queries.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: .\Source\PowerQuery
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
(Optional) The file extension to use for the exported Power Queries' M-Code Formulae.
If not specified, \`.pq\` is used
as the default file extension.
Typically, \`.pq\` is used for Power Query M-Code files, but other extensions are also
common such as \`.m\`, \`.pqm\`, \`.txt\`, etc.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: .pq
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
(Optional) If specified, the function will overwrite any existing files in the specified source code export path.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.Collections.ArrayList
## NOTES
During Development of Excel based applications, an essential component of developing and maintaining the
project's source code is continuous export/import and synchronization of source files with the
host application for portability and most of all, version control.

One area typically overlooked in this regard is the M-Code behind the Power Query components in the workbook's
data model.
Whether it be a Dynamic Query, User Defined Function, Query Parameter, Lookup Table, or any other
Power Query component type (i.e.
template, data source, properties, metadata, etc.), the M-Code behind
the scenes is the foundation that all queries are built from and what drives the core behaviour of the query's
component.

## RELATED LINKS
