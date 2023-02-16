---
external help file: PSXLDevTools-help.xml
Module Name: PSXLDevTools
online version:
schema: 2.0.0
---

# Export-XLShape

## SYNOPSIS
Export Excel shapes.

## SYNTAX

```
Export-XLShape [[-Path] <Object>] [[-Excel] <Object>] [<CommonParameters>]
```

## DESCRIPTION
This script exports properties of every shapes in a Excel file.
It's intended to use them to analyze a diagram made by Excel shapes programatically.

This script only works on Windows with Excel installed.

Almost all properties in exported objects are simply copied from underlying API (Excel COM objects).
So you can find thier meanings or functionalities by searching them on the internet.

## EXAMPLES

### EXAMPLE 1
```
Export-ExcelShape -Path .\test.xlsx | Export-Csv -Path .\out.csv -Encoding UTF8 -NotypeInformation
```

Export shapes in .\test.xlsx as CSV.

### EXAMPLE 2
```
Get-ChildItem -Filter *.xlsx | Export-ExcelShape | Tee-Object -Variable out | Out-GridView
```

Extract shapes from *.xlsx in current directory, Set into $out, and display in gridview.

## PARAMETERS

### -Path
Path to a Excel file

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Excel
Use this parameter to specify your own instance of Excel Application to deal with the Excel file.
If not specified, the script uses its own Excel Application.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
