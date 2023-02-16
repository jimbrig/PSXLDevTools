Function Export-XLShape {
    <#
    .SYNOPSIS
        Export Excel shapes.

    .DESCRIPTION
        This script exports properties of every shapes in a Excel file.
        It's intended to use them to analyze a diagram made by Excel shapes programatically.

        This script only works on Windows with Excel installed.

        Almost all properties in exported objects are simply copied from underlying API (Excel COM objects).
        So you can find thier meanings or functionalities by searching them on the internet.

    .PARAMETER Path
        Path to a Excel file

    .PARAMETER Excel
        Use this parameter to specify your own instance of Excel Application to deal with the Excel file.
        If not specified, the script uses its own Excel Application.

    .EXAMPLE
        PS> Export-ExcelShape -Path .\test.xlsx | Export-Csv -Path .\out.csv -Encoding UTF8 -NotypeInformation

        Export shapes in .\test.xlsx as CSV.
    .EXAMPLE
        PS> Get-ChildItem -Filter *.xlsx | Export-ExcelShape | Tee-Object -Variable out | Out-GridView

        Extract shapes from *.xlsx in current directory, Set into $out, and display in gridview.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $true)]
        $Path,
        $Excel
    )

    Begin {
        $_excel = $null

        if ($null -eq $Excel) {
            Write-Verbose 'Starting Excel app ...'
            $_excel = New-Object -ComObject Excel.Application
        } else {
            $_excel = $Excel
        }

        # @see https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
        $MsoShapeType = @{
            -2 = 'msoShapeTypeMixed' # Mixed shape type
            1  = 'msoAutoShape' # AutoShape
            2  = 'msoCallout' # Callout
            3  = 'msoChart' # Chart
            4  = 'msoComment' # Comment
            5  = 'msoFreeform' # Freeform
            6  = 'msoGroup' # Group
            7  = 'msoEmbeddedOLEObject' # Embedded OLE object
            8  = 'msoFormControl' # Form control
            9  = 'msoLine' # Line
            10 = 'msoLinkedOLEObject' # Linked OLE object
            11 = 'msoLinkedPicture' # Linked picture
            12 = 'msoOLEControlObject' # OLE control object
            13 = 'msoPicture' # Picture
            14 = 'msoPlaceholder' # Placeholder
            15 = 'msoTextEffect' # Text effect
            16 = 'msoMedia' # Media
            17 = 'msoTextBox' # Text box
            18 = 'msoScriptAnchor' # Script anchor
            19 = 'msoTable' # Table
            20 = 'msoCanvas' # Canvas
            21 = 'msoDiagram' # Diagram
            22 = 'msoInk' # Ink
            23 = 'msoInkComment' # Ink comment
            24 = 'msoIgxGraphic' # SmartArt graphic
            26 = 'msoWebVideo' # Web video
            27 = 'msoContentApp' # Content Office Add-in
            28 = 'msoGraphic' # Graphic
            29 = 'msoLinkedGraphic' # Linked graphic
            30 = 'mso3DModel' # 3D model
            31 = 'msoLinked3DModel' # Linked 3D model
        }

        # @see https://docs.microsoft.com/en-us/office/vba/api/office.msoautoshapetype
        $MsoAutoShapeType = @{
            -2  = 'msoShapeMixed' # Return value only; indicates a combination of the other states.
            1   = 'msoShapeRectangle' # Rectangle
            2   = 'msoShapeParallelogram' # Parallelogram
            3   = 'msoShapeTrapezoid' # Trapezoid
            4   = 'msoShapeDiamond' # Diamond
            5   = 'msoShapeRoundedRectangle' # Rounded rectangle
            6   = 'msoShapeOctagon' # Octagon
            7   = 'msoShapeIsoscelesTriangle' # Isosceles triangle
            8   = 'msoShapeRightTriangle' # Right triangle
            9   = 'msoShapeOval' # Oval
            10  = 'msoShapeHexagon' # Hexagon
            11  = 'msoShapeCross' # Cross
            12  = 'msoShapeRegularPentagon' # Pentagon
            13  = 'msoShapeCan' # Can
            14  = 'msoShapeCube' # Cube
            15  = 'msoShapeBevel' # Bevel
            16  = 'msoShapeFoldedCorner' # Folded corner
            17  = 'msoShapeSmileyFace' # Smiley face
            18  = 'msoShapeDonut' # Donut
            19  = 'msoShapeNoSymbol' # "No" symbol
            20  = 'msoShapeBlockArc' # Block arc
            21  = 'msoShapeHeart' # Heart
            22  = 'msoShapeLightningBolt' # Lightning bolt
            23  = 'msoShapeSun' # Sun
            24  = 'msoShapeMoon' # Moon
            25  = 'msoShapeArc' # Arc
            26  = 'msoShapeDoubleBracket' # Double bracket
            27  = 'msoShapeDoubleBrace' # Double brace
            28  = 'msoShapePlaque' # Plaque
            29  = 'msoShapeLeftBracket' # Left bracket
            30  = 'msoShapeRightBracket' # Right bracket
            31  = 'msoShapeLeftBrace' # Left brace
            32  = 'msoShapeRightBrace' # Right brace
            33  = 'msoShapeRightArrow' # Block arrow that points right
            34  = 'msoShapeLeftArrow' # Block arrow that points left
            35  = 'msoShapeUpArrow' # Block arrow that points up
            36  = 'msoShapeDownArrow' # Block arrow that points down
            37  = 'msoShapeLeftRightArrow' # Block arrow with arrowheads that point both left and right
            38  = 'msoShapeUpDownArrow' # Block arrow that points up and down
            39  = 'msoShapeQuadArrow' # Block arrows that point up, down, left, and right
            40  = 'msoShapeLeftRightUpArrow' # Block arrow with arrowheads that point left, right, and up
            41  = 'msoShapeBentArrow' # Block arrow that follows a curved 90-degree angle.
            42  = 'msoShapeUTurnArrow' # Block arrow forming a U shape
            43  = 'msoShapeLeftUpArrow' # Block arrow with arrowheads that point left and up
            44  = 'msoShapeBentUpArrow' # Block arrow that follows a sharp 90-degree angle. Points up by default.
            45  = 'msoShapeCurvedRightArrow' # Block arrow that curves right
            46  = 'msoShapeCurvedLeftArrow' # Block arrow that curves left
            47  = 'msoShapeCurvedUpArrow' # Block arrow that curves up
            48  = 'msoShapeCurvedDownArrow' # Block arrow that curves down
            49  = 'msoShapeStripedRightArrow' # Block arrow that points right with stripes at the tail
            50  = 'msoShapeNotchedRightArrow' # Notched block arrow that points right
            51  = 'msoShapePentagon' # Pentagon
            52  = 'msoShapeChevron' # Chevron
            53  = 'msoShapeRightArrowCallout' # Callout with arrow that points right
            54  = 'msoShapeLeftArrowCallout' # Callout with arrow that points left
            55  = 'msoShapeUpArrowCallout' # Callout with arrow that points up
            56  = 'msoShapeDownArrowCallout' # Callout with arrow that points down
            57  = 'msoShapeLeftRightArrowCallout' # Callout with arrowheads that point both left and right
            58  = 'msoShapeUpDownArrowCallout' # Callout with arrows that point up and down
            59  = 'msoShapeQuadArrowCallout' # Callout with arrows that point up, down, left, and right
            60  = 'msoShapeCircularArrow' # Block arrow that follows a curved 180-degree angle
            61  = 'msoShapeFlowchartProcess' # Process flowchart symbol
            62  = 'msoShapeFlowchartAlternateProcess' # Alternate process flowchart symbol
            63  = 'msoShapeFlowchartDecision' # Decision flowchart symbol
            64  = 'msoShapeFlowchartData' # Data flowchart symbol
            65  = 'msoShapeFlowchartPredefinedProcess' # Predefined process flowchart symbol
            66  = 'msoShapeFlowchartInternalStorage' # Internal storage flowchart symbol
            67  = 'msoShapeFlowchartDocument' # Document flowchart symbol
            68  = 'msoShapeFlowchartMultidocument' # Multi-document flowchart symbol
            69  = 'msoShapeFlowchartTerminator' # Terminator flowchart symbol
            70  = 'msoShapeFlowchartPreparation' # Preparation flowchart symbol
            71  = 'msoShapeFlowchartManualInput' # Manual input flowchart symbol
            72  = 'msoShapeFlowchartManualOperation' # Manual operation flowchart symbol
            73  = 'msoShapeFlowchartConnector' # Connector flowchart symbol
            74  = 'msoShapeFlowchartOffpageConnector' # Off-page connector flowchart symbol
            75  = 'msoShapeFlowchartCard' # Card flowchart symbol
            76  = 'msoShapeFlowchartPunchedTape' # Punched tape flowchart symbol
            77  = 'msoShapeFlowchartSummingJunction' # Summing junction flowchart symbol
            78  = 'msoShapeFlowchartOr' # "Or" flowchart symbol
            79  = 'msoShapeFlowchartCollate' # Collate flowchart symbol
            80  = 'msoShapeFlowchartSort' # Sort flowchart symbol
            81  = 'msoShapeFlowchartExtract' # Extract flowchart symbol
            82  = 'msoShapeFlowchartMerge' # Merge flowchart symbol
            83  = 'msoShapeFlowchartStoredData' # Stored data flowchart symbol
            84  = 'msoShapeFlowchartDelay' # Delay flowchart symbol
            85  = 'msoShapeFlowchartSequentialAccessStorage' # Sequential access storage flowchart symbol
            86  = 'msoShapeFlowchartMagneticDisk' # Magnetic disk flowchart symbol
            87  = 'msoShapeFlowchartDirectAccessStorage' # Direct access storage flowchart symbol
            88  = 'msoShapeFlowchartDisplay' # Display flowchart symbol
            89  = 'msoShapeExplosion1' # Explosion
            90  = 'msoShapeExplosion2' # Explosion
            91  = 'msoShape4pointStar' # 4-point star
            92  = 'msoShape5pointStar' # 5-point star
            93  = 'msoShape8pointStar' # 8-point star
            94  = 'msoShape16pointStar' # 16-point star
            95  = 'msoShape24pointStar' # 24-point star
            96  = 'msoShape32pointStar' # 32-point star
            97  = 'msoShapeUpRibbon' # Ribbon banner with center area above ribbon ends
            98  = 'msoShapeDownRibbon' # Ribbon banner with center area below ribbon ends
            99  = 'msoShapeCurvedUpRibbon' # Ribbon banner that curves up
            100 = 'msoShapeCurvedDownRibbon' # Ribbon banner that curves down
            101 = 'msoShapeVerticalScroll' # Vertical scroll
            102 = 'msoShapeHorizontalScroll' # Horizontal scroll
            103 = 'msoShapeWave' # Wave
            104 = 'msoShapeDoubleWave' # Double wave
            105 = 'msoShapeRectangularCallout' # Rectangular callout
            106 = 'msoShapeRoundedRectangularCallout' # Rounded rectangle-shaped callout
            107 = 'msoShapeOvalCallout' # Oval-shaped callout
            108 = 'msoShapeCloudCallout' # Cloud callout
            109 = 'msoShapeLineCallout1' # Callout with border and horizontal callout line
            110 = 'msoShapeLineCallout2' # Callout with diagonal straight line
            111 = 'msoShapeLineCallout3' # Callout with angled line
            112 = 'msoShapeLineCallout4' # Callout with callout line segments forming a U-shape
            113 = 'msoShapeLineCallout1AccentBar' # Callout with horizontal accent bar
            114 = 'msoShapeLineCallout2AccentBar' # Callout with diagonal callout line and accent bar
            115 = 'msoShapeLineCallout3AccentBar' # Callout with angled callout line and accent bar
            116 = 'msoShapeLineCallout4AccentBar' # Callout with accent bar and callout line segments forming a U-shape
            117 = 'msoShapeLineCallout1NoBorder' # Callout with horizontal line
            118 = 'msoShapeLineCallout2NoBorder' # Callout with no border and diagonal callout line
            119 = 'msoShapeLineCallout3NoBorder' # Callout with no border and angled callout line
            120 = 'msoShapeLineCallout4NoBorder' # Callout with no border and callout line segments forming a U-shape
            121 = 'msoShapeLineCallout1BorderandAccentBar' # Callout with border and horizontal accent bar
            122 = 'msoShapeLineCallout2BorderandAccentBar' # Callout with border, diagonal straight line, and accent bar
            123 = 'msoShapeLineCallout3BorderandAccentBar' # Callout with border, angled callout line, and accent bar
            124 = 'msoShapeLineCallout4BorderandAccentBar' # Callout with border, accent bar, and callout line segments forming a U-shape
            125 = 'msoShapeActionButtonCustom' # Button with no default picture or text. Supports mouse-click and mouse-over actions.
            126 = 'msoShapeActionButtonHome' # Home button. Supports mouse-click and mouse-over actions.
            127 = 'msoShapeActionButtonHelp' # Help button. Supports mouse-click and mouse-over actions.
            128 = 'msoShapeActionButtonInformation' # Information button. Supports mouse-click and mouse-over actions.
            129 = 'msoShapeActionButtonBackorPrevious' # Back or Previous button. Supports mouse-click and mouse-over actions.
            130 = 'msoShapeActionButtonForwardorNext' # Forward or Next button. Supports mouse-click and mouse-over actions.
            131 = 'msoShapeActionButtonBeginning' # Beginning button. Supports mouse-click and mouse-over actions.
            132 = 'msoShapeActionButtonEnd' # End button. Supports mouse-click and mouse-over actions.
            133 = 'msoShapeActionButtonReturn' # Return button. Supports mouse-click and mouse-over actions.
            134 = 'msoShapeActionButtonDocument' # Document button. Supports mouse-click and mouse-over actions.
            135 = 'msoShapeActionButtonSound' # Sound button. Supports mouse-click and mouse-over actions.
            136 = 'msoShapeActionButtonMovie' # Movie button. Supports mouse-click and mouse-over actions.
            137 = 'msoShapeBalloon' # Balloon
            138 = 'msoShapeNotPrimitive' # Not supported
            139 = 'msoShapeFlowchartOfflineStorage' # Offline storage flowchart symbol
            140 = 'msoShapeLeftRightRibbon' # Ribbon with an arrow at both ends
            141 = 'msoShapeDiagonalStripe' # Rectangle with two triangles-shapes removed; a diagonal stripe
            142 = 'msoShapePie' # Circle ('pie') with a portion missing
            143 = 'msoShapeNonIsoscelesTrapezoid' # Trapezoid with asymmetrical non-parallel sides
            144 = 'msoShapeDecagon' # Decagon
            145 = 'msoShapeHeptagon' # Heptagon
            146 = 'msoShapeDodecagon' # Dodecagon
            147 = 'msoShape6pointStar' # 6-point star
            148 = 'msoShape7pointStar' # 7-point star
            149 = 'msoShape10pointStar' # 10-point star
            150 = 'msoShape12pointStar' # 12-point star
            151 = 'msoShapeRound1Rectangle' # Rectangle with one rounded corner
            152 = 'msoShapeRound2SameRectangle' # Rectangle with two-rounded corners that share a side
            154 = 'msoShapeSnipRoundRectangle' # Rectangle with one snipped corner and one rounded corner
            155 = 'msoShapeSnip1Rectangle' # Rectangle with one snipped corner
            156 = 'msoShapeSnip2SameRectangle' # Rectangle with two snipped corners that share a side
            157 = 'msoShapeRound2DiagRectangle' # Rectangle with two rounded corners, diagonally-opposed
            # 157 = 'msoShapeSnip2DiagRectangle' # Rectangle with two snipped corners, diagonally-opposed
            158 = 'msoShapeFrame' # Rectangular picture frame
            159 = 'msoShapeHalfFrame' # Half of a rectangular picture frame
            160 = 'msoShapeTear' # Water droplet
            161 = 'msoShapeChord' # Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord
            162 = 'msoShapeCorner' # Rectangle with rectangular-shaped hole.
            163 = 'msoShapeMathPlus' # Addition symbol +
            164 = 'msoShapeMathMinus' # Subtraction symbol -
            165 = 'msoShapeMathMultiply' # Multiplication symbol x
            166 = 'msoShapeMathDivide' # Division symbol
            167 = 'msoShapeMathEqual' # Equivalence symbol =
            168 = 'msoShapeMathNotEqual' # Non-equivalence symbol
            169 = 'msoShapeCornerTabs' # Four right triangles aligning along a rectangular path; four 'snipped' corners.
            170 = 'msoShapeSquareTabs' # Four small squares that define a rectangular shape
            171 = 'msoShapePlaqueTabs' # Four quarter-circles defining a rectangular shape
            172 = 'msoShapeGear6' # Gear with six teeth
            173 = 'msoShapeGear9' # Gear with nine teeth
            174 = 'msoShapeFunnel' # Funnel
            175 = 'msoShapePieWedge' # Quarter of a circular shape
            176 = 'msoShapeLeftCircularArrow' # Circular arrow pointing counter-clockwise
            177 = 'msoShapeLeftRightCircularArrow' # Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends
            178 = 'msoShapeSwooshArrow' # Curved arrow
            179 = 'msoShapeCloud' # Cloud shape
            180 = 'msoShapeChartX' # Square divided into four parts along diagonal lines
            181 = 'msoShapeChartStar' # Square divided into six parts along vertical and diagonal lines
            182 = 'msoShapeChartPlus' # Square divided vertically and horizontally into four quarters
            183 = 'msoShapeLineInverse' # Line inverse
        }

        # @see https://docs.microsoft.com/en-us/office/vba/api/Office.MsoArrowheadStyle
        $MsoArrowheadStyle = @{
            -2 = 'msoArrowheadStyleMixed' # Return value only; indicates a combination of the other states.
            1  = 'msoArrowheadNone' # No arrowhead
            2  = 'msoArrowheadTriangle' # Triangular
            3  = 'msoArrowheadOpen' # Open
            4  = 'msoArrowheadStealth' # Stealth-shaped
            5  = 'msoArrowheadDiamond' # Diamond-shaped
            6  = 'msoArrowheadOval' # Oval-shaped
        }

        # @see https://docs.microsoft.com/en-us/office/vba/api/office.msolinestyle
        $MsoLineStyle = @{
            -2 = 'msoLineStyleMixed' # Not supported.
            1  = 'msoLineSingle' # Single line.
            2  = 'msoLineThinThin' # Two thin lines.
            3  = 'msoLineThinThick' # Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
            4  = 'msoLineThickThin' # Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
            5  = 'msoLineThickBetweenThin' # Thick line with a thin line on each side.
        }

        # @see https://docs.microsoft.com/en-us/office/vba/api/office.msofilltype
        $MsoFillType = @{
            -2 = 'msoFillMixed' # Mixed fill
            1  = 'msoFillSolid' # Solid fill
            2  = 'msoFillPatterned' # Patterned fill
            3  = 'msoFillGradient' # Gradient fill
            4  = 'msoFillTextured' # Textured fill
            5  = 'msoFillBackground' # Fill is the same as the background.
            6  = 'msoFillPicture' # Picture fill
        }
    }

    Process {
        if ($null -eq $_excel) {
            Write-Error 'Excel app not exists.'
            return
        }

        if ([string]::IsNullOrEmpty($Path) -or $Path -notmatch '\.xls[xbm]?$') {
            Write-Warning ('Skipped because it seems like not excel file: {0}' -f $Path)
            return
        }

        Write-Verbose ('Opening Workbook: {0}' -f $Path)

        $book = $_excel.Workbooks.Open($Path, $False, $True)

        if ($null -eq $book) {
            return
        }

        $book.Worksheets |
            ForEach-Object -Process {
                $sheet = $_

                Write-Verbose ('Processing Worksheet: {0} (Shapes.Count={1})' -f $sheet.Name, $sheet.Shapes.Count)

                for ($i = 1; $i -le $sheet.Shapes.Count; $i++) {
                    Export-Shape -WorkbookName $book.Name -WorksheetName $sheet.Name -Shape $sheet.Shapes.Item($i)
                }

                $sheet = $null
            }

        Write-Verbose ('Closing Workbook')

        $book.Saved = $True
        $book.Close()
        $book = $null
    }

    End {
        if ($null -eq $Excel) {
            if ($null -ne $_excel) {
                Write-Verbose 'Closing Excel app ...'
                $_excel.Quit()
                $_excel = $null

                [GC]::Collect()
            }
        }
    }
}


Function Export-Shape {
    Param(
        $WorkbookName,
        $WorksheetName,
        $Shape
    )

    $prop = [Ordered]@{
        Path                                  = $Path
        Workbook                              = $WorkbookName
        Worksheet                             = $WorksheetName
        Id                                    = $Shape.Id
        Name                                  = $Shape.Name
        Type                                  = $Shape.Type
        TypeName                              = $MsoShapeType[$Shape.Type]
        AutoShapeType                         = $Shape.AutoShapeType
        AutoShapeTypeName                     = $MsoAutoShapeType[$Shape.AutoShapeType]
        ParentGroup                           = $( if ($null -ne $Shape.ParentGroup) { $Shape.ParentGroup.Id } )
        Text                                  = $( if ($Shape.TextFrame2.HasText) { $Shape.TextFrame2.TextRange.Text } )
        Visible                               = $Shape.Visible
        TopLeftCell                           = $Shape.TopLeftCell.Address($False, $False)
        BottomRightCell                       = $Shape.BottomRightCell.Address($False, $False)
        Top                                   = $Shape.Top
        Left                                  = $Shape.Left
        Width                                 = $Shape.Width
        Height                                = $Shape.Height
        HorizontalFlip                        = $Shape.HorizontalFlip
        VerticalFlip                          = $Shape.VerticalFlip
        Rotation                              = $Shape.Rotation
        ZOrderPosition                        = $Shape.ZOrderPosition
        'Line.BackColor'                      = $Shape.Line.BackColor.RGB
        'Line.ForeColor'                      = $Shape.Line.ForeColor.RGB
        'Line.Style'                          = $Shape.Line.Style
        'Line.StyleName'                      = $MsoLineStyle[$Shape.Line.Style]
        'Line.Transparency'                   = $Shape.Line.Transparency
        'Line.Weight'                         = $Shape.Line.Weight
        'Line.BeginArrowheadStyle'            = $Shape.Line.BeginArrowheadStyle
        'Line.BeginArrowheadStyleName'        = $MsoArrowheadStyle[$Shape.Line.BeginArrowheadStyle]
        'Line.EndArrowheadStyle'              = $Shape.Line.EndArrowheadStyle
        'Line.EndArrowheadStyleName'          = $MsoArrowheadStyle[$Shape.Line.EndArrowheadStyle]
        'Fill.BackColor'                      = $Shape.Fill.BackColor.RGB
        'Fill.ForeColor'                      = $Shape.Fill.ForeColor.RGB
        'Fill.Transparency'                   = $Shape.Fill.Transparency
        'Fill.Type'                           = $Shape.Fill.Type
        'Fill.TypeName'                       = $MsoFillType[$Shape.Fill.Type]
        'Nodes.Count'                         = $Shape.Nodes.Count
        'Nodes.Points'                        = $(
            if (0 -lt $Shape.Nodes.Count) {
                # @see https://stackoverflow.com/questions/48168130/how-do-i-move-around-nodes-in-a-shape#answer-48169908
                $Shape.Nodes.Insert($Shape.Nodes.Count, 0, 1, 0, 0)
                $Shape.Nodes.Delete($Shape.Nodes.Count)

                $points = '['
                for ($i = 1; $i -le $Shape.Nodes.Count; $i++) {
                    $points += '[{0},{1}],' -f $Shape.Nodes.Item($i).Points[1, 1], $Shape.Nodes.Item($i).Points[1, 2]
                }
                $points = $points.Substring(0, $points.Length - 1) + ']'

                Write-Output $points
            } else {
                Write-Output '[]'
            }
        )
        'ConnectorFormat.BeginConnectedShape' = $( if ($Shape.ConnectorFormat.BeginConnected) { $Shape.ConnectorFormat.BeginConnectedShape.Id } )
        'ConnectorFormat.BeginConnectionSite' = $( if ($Shape.ConnectorFormat.BeginConnected) { $Shape.ConnectorFormat.BeginConnectionSite } )
        'ConnectorFormat.EndConnectedShape'   = $( if ($Shape.ConnectorFormat.EndConnected) { $Shape.ConnectorFormat.EndConnectedShape.Id } )
        'ConnectorFormat.EndConnectionSite'   = $( if ($Shape.ConnectorFormat.EndConnected) { $Shape.ConnectorFormat.EndConnectionSite } )
    }

    New-Object -TypeName PSObject -Property $prop

    if ($null -ne $Shape.GroupItems) {
        for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
            Export-Shape -WorkbookName $WorkbookName -WorksheetName $WorksheetName -Shape $Shape.GroupItems.Item($i)
        }
    }

    $Shape = $null
}

