'           Bitpusher
'            \`._,'/
'            (_- -_)
'              \o/
'          The Digital
'              Fox
'          @VinceVulpes
'    https://theTechRelay.com
' https://github.com/bitpusher2k
'
' LibreOfficeMacros.bas - By Bitpusher/The Digital Fox
' v1.0.0 last updated 2026-04-18
' LibreOffice Calc port of ExcelMacros.vba (v1.6.3)
' Simple set of useful LibreOffice Calc macros.
'
' Converted from Excel VBA to LibreOffice Basic (UNO API).
' Original VBA version: https://github.com/bitpusher2k/ExcelMacros
'
' Key conversion notes (VBA -> LibreOffice Basic):
'   - ActiveSheet       -> ThisComponent.getCurrentController().getActiveSheet()
'   - ActiveWorkbook     -> ThisComponent
'   - Cells(row, col)    -> oSheet.getCellByPosition(col-1, row-1) [0-indexed]
'   - Range("A1")        -> oSheet.getCellRangeByName("A1")
'   - Application.ScreenUpdating = False -> ThisComponent.lockControllers()
'   - Application.ScreenUpdating = True  -> ThisComponent.unlockControllers()
'   - Selection          -> ThisComponent.getCurrentSelection()
'   - ActiveCell         -> getCurrentSelection() (single cell)
'   - Columns(i).Hidden  -> oSheet.getColumns().getByIndex(i-1).IsVisible = False
'   - Rows(j).Hidden     -> oSheet.getRows().getByIndex(j-1).IsVisible = False
'   - .Interior.Color    -> .CellBackColor (same Long color values - but BGR vs RGB)
'   - ActiveWindow.FreezePanes -> getCurrentController().freezeAtPosition(col, row)
'   - AutoFilter         -> .uno:DataFilterAutoFilter dispatch
'   - VBScript RegExp    -> com.sun.star.util.TextSearch or manual Like pattern
'   - MsgBox             -> MsgBox (same syntax in LibreOffice Basic)
'   - WorksheetFunction.CountIf -> use sheet function via functionAccess
'   - Selection.AutoFilter -> dispatch .uno:DataFilterAutoFilter
'   - ExportAsFixedFormat -> storeToURL with "calc_pdf_Export" filter
'   - SaveAs XLSX        -> storeToURL with "Calc MS Excel 2007 XML" filter
'   - xlNone (color)     -> -1 (com.sun.star.util.Color NOT_SET) or resetPropertyToDefault
'
' Color note: Excel VBA uses BGR ordering for Long colors (e.g. RGB(255,255,0) = 65535).
'   LibreOffice uses standard RGB hex Long (e.g. RGB(255,255,0) = &H00FFFF00 = 16776960).
'   All color values have been converted accordingly.
'
' References:
'   LibreOffice Basic Programming Guide:
'     https://wiki.documentfoundation.org/Macros/Basic/Calc
'   LibreOffice API Reference:
'     https://api.libreoffice.org/docs/idl/ref/
'   VBA Compatibility in LibreOffice:
'     https://help.libreoffice.org/latest/en-US/text/sbasic/shared/vbasupport.html
'   Calc Macros Guide (Ch.12):
'     https://wiki.documentfoundation.org/images/a/a1/CG7212-CalcMacros.pdf
'   Calc as Simple Database (Ch.13, filters):
'     https://wiki.documentfoundation.org/images/9/95/CG6413-CalcAsASimpleDatabase.pdf
'   freezeAtPosition usage:
'     https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=72745
'
' Usage in LibreOffice Calc:
'
' To install macros into LibreOffice:
'   1. Open LibreOffice Calc
'   2. Go to Tools > Macros > Organize Macros > LibreOffice Basic
'   3. Select "My Macros & Dialogs" > "Standard" > "Module1"
'   4. Click "Edit" to open the Basic IDE
'   5. Paste desired macros from this file into the module and save
'
' To add macros to a custom toolbar:
'   1. Go to Tools > Customize > Toolbars tab
'   2. Click "Add" to create a new toolbar or select existing
'   3. Under Category, expand "LibreOffice Basic Macros" > "My Macros" > "Standard" > "Module1"
'   4. Select desired macro and click "Add"
'   5. Use "Modify" to rename or change icon
'
' Alternatively, assign macros to keyboard shortcuts:
'   1. Go to Tools > Customize > Keyboard tab
'   2. Select desired shortcut key combination
'   3. Under Category, navigate to macro location as above
'   4. Select macro and click "Modify"
'
' #libreoffice #calc #basic #macro #useful #toolbar #autofit #row #column
' #filter #guid #highlight #selected #blankiferror #formula #value #duplicate #xlsx #pdf


' ============================================================================
' Helper function: Convert Excel BGR Long color to LibreOffice RGB Long color
' Excel VBA stores colors as BGR (Blue * 65536 + Green * 256 + Red)
' LibreOffice stores colors as RGB (Red * 65536 + Green * 256 + Blue)
' ============================================================================
Function BGRtoRGB(ByVal bgrColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = bgrColor Mod 256
    g = (bgrColor \ 256) Mod 256
    b = (bgrColor \ 65536) Mod 256
    BGRtoRGB = RGB(r, g, b)
End Function

' Pre-computed color constants (already converted from Excel BGR to LO RGB):
' Excel: rgbYellow = 65535 (BGR) -> RGB(255, 255, 0) = 16776960
' Excel: rgbPaleGoldenrod = 7071982 (BGR) -> RGB(238, 232, 107) ~= adjusted
' Excel: rgbPink = 13353215 (BGR) -> RGB(255, 192, 203) = 16761035
' Excel: rgbPaleVioletRed = 9662683 (BGR) -> RGB(219, 112, 147) = 14381203
' Excel: rgbOrangeRed = 17919 (BGR) -> RGB(255, 69, 0) = 16729344
' Excel: rgbOrange = 42495 (BGR) -> RGB(255, 165, 0) = 16753920
' Excel: rgbPaleGreen = 10025880 (BGR) -> RGB(152, 251, 152) = 10025880 (same in this case)
' Excel: rgbPaleTurquoise = 15658671 (BGR) -> RGB(175, 238, 238) = 11529966
'
' Note: The Excel BGR values are converted at runtime by BGRtoRGB() in the macros below.
' The actual RGB values used by LibreOffice are computed dynamically.


Sub InitializeCsv()
' InitializeCsv Macro - Applies AutoFitAllColumns50, AutoFitAllRows50, AddFilter,
' HideEmptyColumns, and HideGuidColumns macros, then freezes the top row.
' Ref: freezeAtPosition - https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=72745
    Call AutoFitAllColumns50
    Call AutoFitAllRows50
    Call AddFilter
    Call HideEmptyColumns
    Call HideGuidColumns

    ' Navigate to A1 and freeze the first row
    Dim oDoc As Object
    Dim oCtrl As Object
    Dim oSheet As Object
    Dim oCell As Object

    oDoc = ThisComponent
    oCtrl = oDoc.getCurrentController()
    oSheet = oCtrl.getActiveSheet()
    oCell = oSheet.getCellByPosition(0, 0) ' A1

    ' Select A1 first to ensure freeze position is correct
    oCtrl.select(oCell)

    ' Freeze at row 1 (below row 0 = header row)
    ' freezeAtPosition(nColumns, nRows) - columns and rows to freeze
    ' Ref: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XViewFreezable.html
    oCtrl.freezeAtPosition(0, 1)
End Sub


Sub AutoFitAllColumns50()
' AutoFitAllColumns50 Macro - Auto-fits all column width with maximum width of 50 (characters)
' In LibreOffice, column width is in 1/100mm. 50 Excel column-width units ~ 17640 (1/100mm).
' Ref: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableColumn.html
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oColumns As Object
    Dim oRange As Object
    Dim nLastCol As Long
    Dim i As Long

    ' Excel column width unit ~= 352.8 hundredths of mm (based on default font)
    ' 50 Excel units ~= 17640 hundredths of mm
    Const MAX_WIDTH_HUNDREDTHS_MM = 17640

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    ' Determine last used column
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastCol = oCursor.getRangeAddress().EndColumn

    oColumns = oSheet.getColumns()

    For i = 0 To nLastCol
        ' OptimalWidth auto-fits the column
        oColumns.getByIndex(i).OptimalWidth = True

        ' Cap at maximum width
        If oColumns.getByIndex(i).Width > MAX_WIDTH_HUNDREDTHS_MM Then
            oColumns.getByIndex(i).Width = MAX_WIDTH_HUNDREDTHS_MM
        End If
    Next i

    ' Select A1
    oDoc.getCurrentController().select(oSheet.getCellByPosition(0, 0))

    oDoc.unlockControllers()
End Sub


Sub AutoFitAllRows50()
' AutoFitAllRows50 Macro - Auto-fits all row height with maximum height of 50 (points)
' In LibreOffice, row height is in 1/100mm. 50 Excel row-height points ~= 1764 (1/100mm).
' Ref: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableRow.html
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oRows As Object
    Dim nLastRow As Long
    Dim i As Long

    ' 50 points ~= 1764 hundredths of mm (1 point = 35.28 hundredths of mm)
    Const MAX_HEIGHT_HUNDREDTHS_MM = 1764

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    ' Determine last used row
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastRow = oCursor.getRangeAddress().EndRow

    oRows = oSheet.getRows()

    For i = 0 To nLastRow
        ' OptimalHeight auto-fits the row
        oRows.getByIndex(i).OptimalHeight = True

        ' Cap at maximum height
        If oRows.getByIndex(i).Height > MAX_HEIGHT_HUNDREDTHS_MM Then
            oRows.getByIndex(i).Height = MAX_HEIGHT_HUNDREDTHS_MM
        End If
    Next i

    ' Select A1
    oDoc.getCurrentController().select(oSheet.getCellByPosition(0, 0))

    oDoc.unlockControllers()
End Sub


Sub AddFilter()
' AddFilter Macro - Adds AutoFilter to header row across all used columns
' Uses UNO dispatch command equivalent to Data > AutoFilter menu.
' Must select the full header row of the used range so LO recognizes the data region;
' selecting only A1 causes AutoFilter to apply to just the first column.
' Ref: https://ask.libreoffice.org/t/find-autofilter-state-from-macro/58024
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCtrl As Object

    oDoc = ThisComponent
    oCtrl = oDoc.getCurrentController()
    oSheet = oCtrl.getActiveSheet()

    ' Determine the full width of the used area
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    Dim nLastCol As Long
    nLastCol = oCursor.getRangeAddress().EndColumn

    ' Select the entire header row (row 0, from column 0 to last used column)
    Dim oHeaderRange As Object
    oHeaderRange = oSheet.getCellRangeByPosition(0, 0, nLastCol, 0)
    oCtrl.select(oHeaderRange)

    ' Use dispatch to toggle AutoFilter
    Dim oDispatcher As Object
    oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oDispatcher.executeDispatch(oCtrl.getFrame(), ".uno:DataFilterAutoFilter", "", 0, Array())
End Sub


Sub HideEmptyColumns()
' HideEmptyColumns Macro - Hides all columns with data only in the first row
' (which is assumed to be the header row)
' Ref: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableColumn.html
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oColumns As Object
    Dim oRows As Object
    Dim nLastRow As Long
    Dim nLastCol As Long
    Dim i As Long, j As Long
    Dim bHideIt As Boolean

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    ' Determine used area
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    Dim oAddr As Object
    oAddr = oCursor.getRangeAddress()
    nLastRow = oAddr.EndRow
    nLastCol = oAddr.EndColumn

    oColumns = oSheet.getColumns()
    oRows = oSheet.getRows()

    For i = 0 To nLastCol
        bHideIt = True

        For j = 1 To nLastRow  ' Skip row 0 (header)
            ' Only check visible rows
            If oRows.getByIndex(j).IsVisible Then
                Dim oCell As Object
                oCell = oSheet.getCellByPosition(i, j)
                If Trim(oCell.getString()) <> "" Then
                    bHideIt = False
                    Exit For
                End If
            End If
        Next j

        oColumns.getByIndex(i).IsVisible = Not bHideIt
    Next i

    oDoc.unlockControllers()
End Sub


Sub HideGuidColumns()
' HideGuidColumns Macro - Hide all columns with a GUID in the second row
' (the first row is assumed to be the header)
' Uses LibreOffice Basic Like operator for GUID pattern matching
' since VBScript RegExp is not available in LibreOffice.
' GUID format: optional { or ( + 8 hex + - + 4 hex + - + 4 hex + - + 4 hex + - + 12 hex + optional } or )
' Ref: https://help.libreoffice.org/latest/en-US/text/sbasic/shared/03120314.html (Like operator)
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oColumns As Object
    Dim nLastCol As Long
    Dim i As Long
    Dim sCellVal As String

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    ' Determine last used column
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastCol = oCursor.getRangeAddress().EndColumn

    oColumns = oSheet.getColumns()

    For i = 0 To nLastCol
        sCellVal = Trim(oSheet.getCellByPosition(i, 1).getString()) ' Row index 1 = second row
        If IsGUID(sCellVal) Then
            oColumns.getByIndex(i).IsVisible = False
        End If
    Next i

    oDoc.unlockControllers()
End Sub


' Helper function to check if a string matches GUID format
' Supports optional surrounding braces {} or parentheses ()
Function IsGUID(ByVal s As String) As Boolean
    Dim sClean As String
    IsGUID = False
    If Len(s) = 0 Then Exit Function

    sClean = s

    ' Strip optional surrounding braces or parentheses
    If Left(sClean, 1) = "{" And Right(sClean, 1) = "}" Then
        sClean = Mid(sClean, 2, Len(sClean) - 2)
    ElseIf Left(sClean, 1) = "(" And Right(sClean, 1) = ")" Then
        sClean = Mid(sClean, 2, Len(sClean) - 2)
    End If

    ' GUID without braces should be 36 chars: 8-4-4-4-12
    If Len(sClean) <> 36 Then Exit Function

    ' Check dashes at positions 9, 14, 19, 24
    If Mid(sClean, 9, 1) <> "-" Then Exit Function
    If Mid(sClean, 14, 1) <> "-" Then Exit Function
    If Mid(sClean, 19, 1) <> "-" Then Exit Function
    If Mid(sClean, 24, 1) <> "-" Then Exit Function

    ' Check that all other characters are hex digits
    Dim sHexOnly As String
    sHexOnly = Replace(sClean, "-", "")
    If Len(sHexOnly) <> 32 Then Exit Function

    Dim ch As String
    Dim k As Long
    For k = 1 To 32
        ch = UCase(Mid(sHexOnly, k, 1))
        If InStr("0123456789ABCDEF", ch) = 0 Then Exit Function
    Next k

    IsGUID = True
End Function


Sub HighlightCellsWithSelectedValue()
' HighlightCellsWithSelectedValue Macro - Highlights all cells which contain the selected value
' Uses createSearchDescriptor() for cell searching
' Ref: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchable.html
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oSel As Object
    Dim sValue As String
    Dim oSD As Object
    Dim oFound As Object
    Dim i As Long

    ' Yellow in LibreOffice RGB
    Const COLOR_YELLOW = 16776960 ' RGB(255, 255, 0)

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    oSel = oDoc.getCurrentSelection()

    sValue = oSel.getString()
    If sValue = "" Then Exit Sub

    oDoc.lockControllers()

    ' Search entire sheet for matching cells
    oSD = oSheet.createSearchDescriptor()
    oSD.SearchString = sValue
    oSD.SearchWords = False   ' Partial match (like VBA Find)
    oSD.SearchRegularExpression = False

    oFound = oSheet.findAll(oSD)

    If Not IsNull(oFound) Then
        For i = 0 To oFound.getCount() - 1
            oFound.getByIndex(i).CellBackColor = COLOR_YELLOW
        Next i
    End If

    oDoc.unlockControllers()
End Sub


Sub HighlightRowsWithSelectedValue()
' HighlightRowsWithSelectedValue Macro - Highlights all rows that have a cell
' containing the selected value. Row = Pale Goldenrod, matching cell = Yellow.
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sValue As String

    ' Converted from Excel BGR to LibreOffice RGB
    Const COLOR_YELLOW = 16776960       ' RGB(255, 255, 0)
    Const COLOR_PALE_GOLDENROD = 15657130 ' RGB(238, 232, 170) - standard CSS PaleGoldenrod

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    sValue = oDoc.getCurrentSelection().getString()
    If sValue = "" Then Exit Sub

    ' Remove any active filters first so all rows are visible
    Call RemoveSheetFilterIfActive(oSheet)

    Call HighlightRowsByValue(oSheet, sValue, COLOR_PALE_GOLDENROD, COLOR_YELLOW)
End Sub


Sub HighlightRowsWithSelectedValueRed()
' Highlights rows with selected value - Red/Pink color scheme
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sValue As String

    Const COLOR_PINK = 16761035          ' RGB(255, 192, 203)
    Const COLOR_PALE_VIOLET_RED = 14381203 ' RGB(219, 112, 147)

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    sValue = oDoc.getCurrentSelection().getString()
    If sValue = "" Then Exit Sub

    Call RemoveSheetFilterIfActive(oSheet)
    Call HighlightRowsByValue(oSheet, sValue, COLOR_PINK, COLOR_PALE_VIOLET_RED)
End Sub


Sub HighlightRowsWithSelectedValueOrange()
' Highlights rows with selected value - Orange color scheme
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sValue As String

    Const COLOR_ORANGE_RED = 16729344    ' RGB(255, 69, 0)
    Const COLOR_ORANGE = 16753920        ' RGB(255, 165, 0)

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    sValue = oDoc.getCurrentSelection().getString()
    If sValue = "" Then Exit Sub

    Call RemoveSheetFilterIfActive(oSheet)
    Call HighlightRowsByValue(oSheet, sValue, COLOR_ORANGE_RED, COLOR_ORANGE)
End Sub


Sub HighlightRowsWithSelectedValueGreen()
' Highlights rows with selected value - Green color scheme
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sValue As String

    Const COLOR_PALE_GREEN = 10025880    ' RGB(152, 251, 152)
    Const COLOR_PALE_TURQUOISE = 11529966 ' RGB(175, 238, 238)

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    sValue = oDoc.getCurrentSelection().getString()
    If sValue = "" Then Exit Sub

    Call RemoveSheetFilterIfActive(oSheet)
    Call HighlightRowsByValue(oSheet, sValue, COLOR_PALE_GREEN, COLOR_PALE_TURQUOISE)
End Sub


' Helper: Highlight all rows containing a value with specified colors
' Ref: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchable.html
Sub HighlightRowsByValue(oSheet As Object, sValue As String, lRowColor As Long, lCellColor As Long)
    Dim oDoc As Object
    Dim oSD As Object
    Dim oFound As Object
    Dim oCell As Object
    Dim oRows As Object
    Dim i As Long
    Dim nRow As Long

    oDoc = ThisComponent
    oDoc.lockControllers()

    oRows = oSheet.getRows()

    ' Search for all matching cells
    oSD = oSheet.createSearchDescriptor()
    oSD.SearchString = sValue
    oSD.SearchWords = False
    oSD.SearchRegularExpression = False

    oFound = oSheet.findAll(oSD)

    If Not IsNull(oFound) Then
        ' Determine used area for full-row highlighting
        Dim oCursor As Object
        oCursor = oSheet.createCursor()
        oCursor.gotoStartOfUsedArea(False)
        oCursor.gotoEndOfUsedArea(True)
        Dim nLastCol As Long
        nLastCol = oCursor.getRangeAddress().EndColumn

        For i = 0 To oFound.getCount() - 1
            Dim oFoundRange As Object
            oFoundRange = oFound.getByIndex(i)

            ' findAll() returns SheetCellRanges - each element is a CellRange,
            ' not a single Cell, so use getRangeAddress() instead of getCellAddress().
            ' Ref: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetCellRanges.html
            Dim oRangeAddr As Object
            oRangeAddr = oFoundRange.getRangeAddress()

            ' Iterate all rows in this found range (usually just one row per match)
            Dim nFoundRow As Long
            For nFoundRow = oRangeAddr.StartRow To oRangeAddr.EndRow
                ' Highlight entire row (used area portion)
                Dim oRowRange As Object
                oRowRange = oSheet.getCellRangeByPosition(0, nFoundRow, nLastCol, nFoundRow)
                oRowRange.CellBackColor = lRowColor
            Next nFoundRow

            ' Highlight the matching range itself with accent color
            oFoundRange.CellBackColor = lCellColor
        Next i
    End If

    oDoc.unlockControllers()
End Sub


' Helper: Remove sheet filter if one is active (equivalent to ShowAllData)
' Ref: https://pvanb.wordpress.com/2011/05/03/macro-to-remove-filters-in-libreopenoffice/
Sub RemoveSheetFilterIfActive(oSheet As Object)
    Dim oFilterDesc As Object
    oFilterDesc = oSheet.createFilterDescriptor(True) ' Empty = remove filter
    oSheet.filter(oFilterDesc)
End Sub


Sub HideRowsWithSelectedValue()
' HideRowsWithSelectedValue Macro - Hides all rows that have a cell containing the selected value
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sValue As String
    Dim oSD As Object
    Dim oFound As Object
    Dim oCell As Object
    Dim i As Long

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    sValue = oDoc.getCurrentSelection().getString()
    If sValue = "" Then Exit Sub

    Call RemoveSheetFilterIfActive(oSheet)

    oDoc.lockControllers()

    oSD = oSheet.createSearchDescriptor()
    oSD.SearchString = sValue
    oSD.SearchWords = False
    oSD.SearchRegularExpression = False

    oFound = oSheet.findAll(oSD)

    If Not IsNull(oFound) Then
        Dim oRows As Object
        oRows = oSheet.getRows()
        For i = 0 To oFound.getCount() - 1
            ' findAll() returns CellRange objects, use getRangeAddress()
            Dim oFoundRange As Object
            oFoundRange = oFound.getByIndex(i)
            Dim oRangeAddr As Object
            oRangeAddr = oFoundRange.getRangeAddress()
            Dim nFoundRow As Long
            For nFoundRow = oRangeAddr.StartRow To oRangeAddr.EndRow
                oRows.getByIndex(nFoundRow).IsVisible = False
            Next nFoundRow
        Next i
    End If

    oDoc.unlockControllers()
End Sub


Public Sub BlankIfError()
' BlankIfError Macro - Surround formulas in all selected cells with =IFERROR(,"")
' In LibreOffice Calc, IFERROR works the same way as in Excel.
' Ref: https://help.libreoffice.org/latest/en-US/text/scalc/01/04060104.html
    Dim oDoc As Object
    Dim oSel As Object
    Dim oCell As Object
    Dim sFormula As String
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long

    oDoc = ThisComponent
    oSel = oDoc.getCurrentSelection()

    ' Check if selection supports cell enumeration
    If Not HasUnoInterfaces(oSel, "com.sun.star.sheet.XCellRangesQuery") Then
        MsgBox "Please select a range of cells.", 48, "Error"
        Exit Sub
    End If

    oDoc.lockControllers()

    Dim oAddr As Object
    oAddr = oSel.getRangeAddress()
    Dim oSheet As Object
    oSheet = oDoc.getCurrentController().getActiveSheet()

    For r = oAddr.StartRow To oAddr.EndRow
        For c = oAddr.StartColumn To oAddr.EndColumn
            oCell = oSheet.getCellByPosition(c, r)
            sFormula = oCell.getFormula()

            ' Check if cell has a formula (starts with =)
            If Left(sFormula, 1) = "=" Then
                ' Don't wrap if already wrapped in IFERROR
                If LCase(Left(sFormula, 9)) <> "=iferror(" Then
                    ' Semicolon separator required for setFormula() in LibreOffice
                    oCell.setFormula("=IFERROR(" & Mid(sFormula, 2) & ";"""")")
                End If
            End If
        Next c
    Next r

    oDoc.unlockControllers()
End Sub


Sub ConvertSelectedToValues()
' ConvertSelectedToValues Macro - Converts formulas in selected cells to values
    Dim oDoc As Object
    Dim oSel As Object
    Dim oSheet As Object
    Dim oCell As Object
    Dim r As Long, c As Long

    Dim nAnswer As Integer
    nAnswer = MsgBox("Caution: Action cannot be undone. " & _
        "Save Workbook First?", 3 + 48, "Alert")
    ' 3 = vbYesNoCancel, 48 = vbExclamation

    Select Case nAnswer
        Case 6 ' vbYes
            oDoc = ThisComponent
            oDoc.store()
        Case 2 ' vbCancel
            Exit Sub
    End Select

    oDoc = ThisComponent
    oSel = oDoc.getCurrentSelection()
    oSheet = oDoc.getCurrentController().getActiveSheet()

    oDoc.lockControllers()

    Dim oAddr As Object
    oAddr = oSel.getRangeAddress()

    For r = oAddr.StartRow To oAddr.EndRow
        For c = oAddr.StartColumn To oAddr.EndColumn
            oCell = oSheet.getCellByPosition(c, r)
            ' com.sun.star.table.CellContentType.FORMULA = 2
            If oCell.getType() = 2 Then
                ' Get the computed value and replace formula with it
                Dim dValue As Double
                Dim sStr As String
                ' Check if result is numeric or string
                If oCell.getError() = 0 Then
                    Select Case oCell.getType()
                        Case 2 ' FORMULA
                            ' Determine result type by checking FormulaResultType
                            ' 1 = DOUBLE, 2 = STRING
                            If oCell.FormulaResultType2 = com.sun.star.sheet.FormulaResult.STRING Then
                                sStr = oCell.getString()
                                oCell.setFormula("")
                                oCell.setString(sStr)
                            Else
                                dValue = oCell.getValue()
                                oCell.setFormula("")
                                oCell.setValue(dValue)
                            End If
                    End Select
                End If
            End If
        Next c
    Next r

    oDoc.unlockControllers()
End Sub


Sub HighlightDuplicateValuesSelected()
' HighlightDuplicateValuesSelected Macro - Highlights duplicate values in selected range
' Uses LibreOffice's built-in function access for COUNTIF equivalent
' Ref: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XFunctionAccess.html
    Dim oDoc As Object
    Dim oSel As Object
    Dim oSheet As Object
    Dim oCell As Object
    Dim r As Long, c As Long
    Dim sVal As String

    ' ColorIndex 36 in Excel ≈ Light Yellow = RGB(255, 255, 153)
    Const COLOR_LIGHT_YELLOW = 16777113

    oDoc = ThisComponent
    oSel = oDoc.getCurrentSelection()
    oSheet = oDoc.getCurrentController().getActiveSheet()

    oDoc.lockControllers()

    Dim oAddr As Object
    oAddr = oSel.getRangeAddress()

    ' Build a dictionary of value counts
    Dim aValues() As String
    Dim aCounts() As Long
    Dim nEntries As Long
    nEntries = 0

    ' First pass: count all values
    Dim nTotalCells As Long
    nTotalCells = (oAddr.EndRow - oAddr.StartRow + 1) * (oAddr.EndColumn - oAddr.StartColumn + 1)
    ReDim aValues(nTotalCells - 1)
    ReDim aCounts(nTotalCells - 1)

    For r = oAddr.StartRow To oAddr.EndRow
        For c = oAddr.StartColumn To oAddr.EndColumn
            sVal = oSheet.getCellByPosition(c, r).getString()
            Dim bFoundEntry As Boolean
            bFoundEntry = False
            Dim k As Long
            For k = 0 To nEntries - 1
                If aValues(k) = sVal Then
                    aCounts(k) = aCounts(k) + 1
                    bFoundEntry = True
                    Exit For
                End If
            Next k
            If Not bFoundEntry Then
                aValues(nEntries) = sVal
                aCounts(nEntries) = 1
                nEntries = nEntries + 1
            End If
        Next c
    Next r

    ' Second pass: highlight cells whose value appears more than once
    For r = oAddr.StartRow To oAddr.EndRow
        For c = oAddr.StartColumn To oAddr.EndColumn
            oCell = oSheet.getCellByPosition(c, r)
            sVal = oCell.getString()
            For k = 0 To nEntries - 1
                If aValues(k) = sVal Then
                    If aCounts(k) > 1 Then
                        oCell.CellBackColor = COLOR_LIGHT_YELLOW
                    End If
                    Exit For
                End If
            Next k
        Next c
    Next r

    oDoc.unlockControllers()
End Sub


Sub AddFrequencyColumn()
' AddFrequencyColumn Macro - Adds column to right of selected column and populates
' with COUNTIF frequency of values in selected column
' Ref: https://wiki.documentfoundation.org/Macros/Basic/Calc (column/row insertion)
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oSel As Object
    Dim nSelCol As Long
    Dim nLastRow As Long
    Dim i As Long

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    oSel = oDoc.getCurrentSelection()

    ' Get the column of the selection
    nSelCol = oSel.getRangeAddress().StartColumn

    ' Find last row with data in the selected column
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastRow = oCursor.getRangeAddress().EndRow

    If nLastRow < 1 Then
        MsgBox "No data found in the selected column.", 48
        Exit Sub
    End If

    oDoc.lockControllers()

    ' Insert a new column to the right of selected
    oSheet.getColumns().insertByIndex(nSelCol + 1, 1)

    ' Add header
    Dim sOrigHeader As String
    sOrigHeader = oSheet.getCellByPosition(nSelCol, 0).getString()
    If sOrigHeader = "" Then
        oSheet.getCellByPosition(nSelCol + 1, 0).setString("Frequency")
    Else
        oSheet.getCellByPosition(nSelCol + 1, 0).setString(sOrigHeader & "Frequency")
    End If

    ' Get column letter for formula references
    ' In LibreOffice, we can use getCellByPosition and getFormula with $column references
    ' Build COUNTIF formula for each data row
    For i = 1 To nLastRow
        Dim oDataCell As Object
        oDataCell = oSheet.getCellByPosition(nSelCol, i)
        Dim oFreqCell As Object
        oFreqCell = oSheet.getCellByPosition(nSelCol + 1, i)

        ' Build cell address strings for the COUNTIF formula
        ' Get the address of the data range and the current cell
        Dim sDataRangeAddr As String
        Dim sDataColLetter As String
        Dim sCellRef As String

        ' Use the cell address to construct formula
        ' getCellByPosition returns CellAddress with .Column and .Row
        ' We need column letter - convert column index to letter(s)
        sDataColLetter = ColumnIndexToLetter(nSelCol)
        sCellRef = sDataColLetter & CStr(i + 1) ' +1 because LO formulas are 1-based in display
        sDataRangeAddr = "$" & sDataColLetter & "$2:$" & sDataColLetter & "$" & CStr(nLastRow + 1)

        ' setFormula() in LibreOffice requires semicolons as argument separators,
        ' not commas — commas cause Err:508 (parentheses/parsing error).
        ' Ref: https://ask.libreoffice.org/t/result-of-setformula-not-recognised-solved/17383
        oFreqCell.setFormula("=COUNTIF(" & sDataRangeAddr & ";" & sCellRef & ")")
    Next i

    ' Auto-fit the new column
    oSheet.getColumns().getByIndex(nSelCol + 1).OptimalWidth = True

    oDoc.unlockControllers()

    MsgBox "Frequency column added successfully!", 64
End Sub


' Helper: Convert 0-based column index to letter(s) (e.g., 0="A", 25="Z", 26="AA")
Function ColumnIndexToLetter(ByVal nCol As Long) As String
    Dim sResult As String
    sResult = ""
    Do
        sResult = Chr(65 + (nCol Mod 26)) & sResult
        nCol = (nCol \ 26) - 1
    Loop While nCol >= 0
    ColumnIndexToLetter = sResult
End Function


' Helper: Find last occurrence of a substring (replacement for VBA's InStrRev
' which does not exist in LibreOffice Basic)
' Returns 0 if not found, otherwise the 1-based position of the last occurrence.
Function LastInStr(ByVal sText As String, ByVal sSearch As String) As Long
    Dim nPos As Long
    Dim nLast As Long
    nLast = 0
    nPos = InStr(1, sText, sSearch)
    Do While nPos > 0
        nLast = nPos
        nPos = InStr(nPos + 1, sText, sSearch)
    Loop
    LastInStr = nLast
End Function


Sub SaveWorksheetAsPDF()
' SaveWorksheetAsPDF Macro - Saves current worksheet as PDF
' Uses storeToURL with calc_pdf_Export filter
' Ref: https://wiki.documentfoundation.org/Macros/Basic/Calc#Export_as_PDF
    Dim oDoc As Object
    Dim oSheet As Object
    Dim sPath As String
    Dim sName As String
    Dim sTime As String
    Dim sFullPath As String

    On Error GoTo errHandler

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()

    sTime = Format(Now(), "YYYYMMDD\_HHmm")

    ' Get document path
    sPath = ConvertFromURL(oDoc.getURL())
    If sPath = "" Then
        ' Document not saved yet - use home directory
        sPath = Environ("HOME")
        If sPath = "" Then sPath = Environ("USERPROFILE") ' Windows fallback
    Else
        ' Extract directory from full file path
        Dim nPos As Long
        ' Handle both / and \ separators
        nPos = LastInStr(sPath, "/")
        Dim nPos2 As Long
        nPos2 = LastInStr(sPath, "\")
        If nPos2 > nPos Then nPos = nPos2
        If nPos > 0 Then sPath = Left(sPath, nPos)
    End If

    ' Clean sheet name for filename
    sName = oSheet.getName()
    sName = Join(Split(sName, " "), "")
    sName = Join(Split(sName, "."), "_")

    ' Build default filename
    Dim sDefaultFile As String
    sDefaultFile = sPath & sName & "_" & sTime & ".pdf"

    ' Show file picker dialog
    Dim oFilePicker As Object
    oFilePicker = com.sun.star.ui.dialogs.FilePicker.createWithMode( _
        com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_AUTOEXTENSION)
    oFilePicker.setDefaultName(sName & "_" & sTime & ".pdf")
    oFilePicker.appendFilter("PDF Files (*.pdf)", "*.pdf")
    oFilePicker.setCurrentFilter("PDF Files (*.pdf)")

    ' Set initial directory
    If InStr(sPath, "/") > 0 Or InStr(sPath, "\") > 0 Then
        oFilePicker.setDisplayDirectory(ConvertToURL(sPath))
    End If

    Dim nResult As Integer
    nResult = oFilePicker.execute()

    If nResult = com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
        Dim aFiles() As String
        aFiles = oFilePicker.getFiles()
        sFullPath = aFiles(0)

        ' Ensure .pdf extension
        If LCase(Right(sFullPath, 4)) <> ".pdf" Then
            sFullPath = sFullPath & ".pdf"
        End If

        ' Set up PDF export properties
        Dim aPDFProps(0) As New com.sun.star.beans.PropertyValue
        aPDFProps(0).Name = "FilterName"
        aPDFProps(0).Value = "calc_pdf_Export"

        ' Export
        oDoc.storeToURL(sFullPath, aPDFProps())

        MsgBox "PDF file has been created: " & Chr(10) & ConvertFromURL(sFullPath), 64
    End If

    Exit Sub

errHandler:
    MsgBox "Could not create PDF file", 16
End Sub


Sub SaveWorksheetAsXLSX()
' SaveWorksheetAsXLSX Macro - Saves current document as XLSX
' Uses storeToURL with "Calc MS Excel 2007 XML" filter
' Ref: https://wiki.documentfoundation.org/Macros/Basic/Calc
    Dim oDoc As Object
    Dim sURL As String
    Dim sNewURL As String

    oDoc = ThisComponent
    sURL = oDoc.getURL()

    If sURL = "" Then
        MsgBox "Please save the document first.", 48
        Exit Sub
    End If

    ' Replace extension with .xlsx
    Dim nDotPos As Long
    nDotPos = LastInStr(sURL, ".")
    If nDotPos > 0 Then
        sNewURL = Left(sURL, nDotPos - 1) & ".xlsx"
    Else
        sNewURL = sURL & ".xlsx"
    End If

    ' Set up XLSX export properties
    Dim aProps(0) As New com.sun.star.beans.PropertyValue
    aProps(0).Name = "FilterName"
    aProps(0).Value = "Calc MS Excel 2007 XML"

    oDoc.storeToURL(sNewURL, aProps())
End Sub


Sub ClearAllHighlighting()
' ClearAllHighlighting Macro - Clears all cell background colors in the sheet
' Sets CellBackColor to -1 (transparent/no color)
' Ref: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1CellProperties.html
    Dim oDoc As Object
    Dim oSheet As Object

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    ' Get the used range and clear background colors
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)

    ' Set background to transparent (-1 means no color / COL_AUTO)
    oCursor.CellBackColor = -1

    oDoc.unlockControllers()
End Sub


Sub UnhideAllRowsColumns()
' UnhideAllRowsColumns Macro - Un-hides all rows and columns
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oColumns As Object
    Dim oRows As Object
    Dim i As Long

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()
    oColumns = oSheet.getColumns()
    oRows = oSheet.getRows()

    For i = 0 To oColumns.getCount() - 1
        oColumns.getByIndex(i).IsVisible = True
    Next i

    For i = 0 To oRows.getCount() - 1
        oRows.getByIndex(i).IsVisible = True
    Next i

    oDoc.unlockControllers()
End Sub


Sub CustomSort()
' CustomSort Macro - Opens the Sort dialog
' Uses UNO dispatch equivalent to Data > Sort
' Ref: https://wiki.documentfoundation.org/Development/DispatchCommands
    Dim oDispatcher As Object
    oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oDispatcher.executeDispatch( _
        ThisComponent.getCurrentController().getFrame(), _
        ".uno:DataSort", "", 0, Array())
End Sub


Sub DeleteHiddenColumns()
' DeleteHiddenColumns Macro - Deletes all hidden columns
' Iterates in reverse to avoid index shifting issues
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oColumns As Object
    Dim nLastCol As Long
    Dim i As Long

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastCol = oCursor.getRangeAddress().EndColumn

    oColumns = oSheet.getColumns()

    For i = nLastCol To 0 Step -1
        If Not oColumns.getByIndex(i).IsVisible Then
            oColumns.removeByIndex(i, 1)
        End If
    Next i

    oDoc.unlockControllers()
End Sub


Sub DeleteHiddenRows()
' DeleteHiddenRows Macro - Deletes all hidden rows
' Iterates in reverse to avoid index shifting issues
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oRows As Object
    Dim nLastRow As Long
    Dim i As Long

    oDoc = ThisComponent
    oDoc.lockControllers()

    oSheet = oDoc.getCurrentController().getActiveSheet()

    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastRow = oCursor.getRangeAddress().EndRow

    oRows = oSheet.getRows()

    For i = nLastRow To 0 Step -1
        If Not oRows.getByIndex(i).IsVisible Then
            oRows.removeByIndex(i, 1)
        End If
    Next i

    oDoc.unlockControllers()
End Sub


Sub SplitDateAndTimeToNewColumns()
' SplitDateAndTimeToNewColumns Macro - Splits selected column with date and time
' to new "DateOnly" and "TimeOnly" columns created to the right
' Ref: https://help.libreoffice.org/latest/en-US/text/sbasic/shared/03030000.html (Date/Time functions)
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oSel As Object
    Dim nSelCol As Long
    Dim nLastRow As Long
    Dim i As Long

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    oSel = oDoc.getCurrentSelection()

    If IsNull(oSel) Then
        MsgBox "Please select a cell in a column containing date and time data.", 48, "Error"
        Exit Sub
    End If

    nSelCol = oSel.getRangeAddress().StartColumn

    ' Find last row
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    nLastRow = oCursor.getRangeAddress().EndRow

    If nLastRow < 1 Then
        MsgBox "No data found.", 48
        Exit Sub
    End If

    oDoc.lockControllers()

    ' Insert two new columns to the right
    oSheet.getColumns().insertByIndex(nSelCol + 1, 2)

    ' Set headers
    oSheet.getCellByPosition(nSelCol + 1, 0).setString("DateOnly")
    oSheet.getCellByPosition(nSelCol + 2, 0).setString("TimeOnly")

    ' Process each row (skip header row 0)
    For i = 1 To nLastRow
        Dim oSrcCell As Object
        oSrcCell = oSheet.getCellByPosition(nSelCol, i)

        ' In LibreOffice, dates are stored as numeric values
        ' The integer part is the date serial, fractional part is the time
        Dim dDateTime As Double
        dDateTime = oSrcCell.getValue()

        If dDateTime <> 0 Then
            Dim dDatePart As Double
            Dim dTimePart As Double
            dDatePart = Int(dDateTime)
            dTimePart = dDateTime - dDatePart

            ' Set date value and format
            Dim oDateCell As Object
            oDateCell = oSheet.getCellByPosition(nSelCol + 1, i)
            oDateCell.setValue(dDatePart)

            ' Apply YYYY-MM-DD date format
            Dim oFormats As Object
            oFormats = oDoc.getNumberFormats()
            Dim oLocale As New com.sun.star.lang.Locale
            Dim nDateFmt As Long
            nDateFmt = oFormats.getStandardFormat( _
                com.sun.star.util.NumberFormat.DATE, oLocale)
            ' Try to get/add custom format
            Dim nCustomDateFmt As Long
            nCustomDateFmt = oFormats.queryKey("YYYY-MM-DD", oLocale, False)
            If nCustomDateFmt = -1 Then
                nCustomDateFmt = oFormats.addNew("YYYY-MM-DD", oLocale)
            End If
            oDateCell.NumberFormat = nCustomDateFmt

            ' Set time value and format
            Dim oTimeCell As Object
            oTimeCell = oSheet.getCellByPosition(nSelCol + 2, i)
            oTimeCell.setValue(dTimePart)

            Dim nTimeFmt As Long
            nTimeFmt = oFormats.queryKey("HH:MM:SS", oLocale, False)
            If nTimeFmt = -1 Then
                nTimeFmt = oFormats.addNew("HH:MM:SS", oLocale)
            End If
            oTimeCell.NumberFormat = nTimeFmt
        End If
    Next i

    ' Auto-fit new columns
    oSheet.getColumns().getByIndex(nSelCol + 1).OptimalWidth = True
    oSheet.getColumns().getByIndex(nSelCol + 2).OptimalWidth = True

    oDoc.unlockControllers()
End Sub


Sub CheckValueMatch()
' CheckValueMatch Macro - Compares values in two selected columns.
' Select exactly two separate column ranges (Ctrl+Click).
' If a value from the first column appears anywhere in the second column,
' marks "TRUE" in a new column to the right of the second column.
'
' Note: LibreOffice multi-selection works differently than Excel.
' The user must Ctrl+Click to select two separate column ranges.
' Ref: https://www.debugpoint.com/calc-cell-selection-processing-using-macro/
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oSel As Object
    Dim bIsMultiRange As Boolean

    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    oSel = oDoc.getCurrentSelection()

    ' Check if selection is null
    If IsNull(oSel) Or IsEmpty(oSel) Then
        MsgBox "Please select exactly two columns (hold Ctrl to select both).", 48
        Exit Sub
    End If

    ' Check if we have a multi-range selection (two areas).
    ' Must guard supportsService with error handling because calling it
    ' on non-cell objects (charts, drawing shapes) can throw a RuntimeException.
    bIsMultiRange = False
    On Local Error Resume Next
    bIsMultiRange = oSel.supportsService("com.sun.star.sheet.SheetCellRanges")
    On Local Error GoTo 0

    If Not bIsMultiRange Then
        MsgBox "Please select exactly two columns (hold Ctrl to select both)." & Chr(10) & _
               "Use Ctrl+Click to select the second column.", 48
        Exit Sub
    End If

    Dim nAreaCount As Long
    nAreaCount = 0
    On Local Error Resume Next
    nAreaCount = oSel.getCount()
    On Local Error GoTo 0

    If nAreaCount <> 2 Then
        MsgBox "Please select exactly two columns (hold Ctrl to select both).", 48
        Exit Sub
    End If

    Dim oFirst As Object, oSecond As Object
    oFirst = oSel.getByIndex(0)
    oSecond = oSel.getByIndex(1)

    Dim oFirstAddr As Object, oSecondAddr As Object
    oFirstAddr = oFirst.getRangeAddress()
    oSecondAddr = oSecond.getRangeAddress()

    ' Validate single column selections
    If oFirstAddr.StartColumn <> oFirstAddr.EndColumn Or _
       oSecondAddr.StartColumn <> oSecondAddr.EndColumn Then
        MsgBox "Please select single columns only.", 48
        Exit Sub
    End If

    Dim nFirstCol As Long, nSecondCol As Long
    nFirstCol = oFirstAddr.StartColumn
    nSecondCol = oSecondAddr.StartColumn

    ' Determine actual data bounds using the sheet's used area,
    ' because selecting entire columns gives EndRow = 1048575 which
    ' would cause an extremely long loop.
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(False)
    oCursor.gotoEndOfUsedArea(True)
    Dim nSheetLastRow As Long
    nSheetLastRow = oCursor.getRangeAddress().EndRow

    ' Clamp selection row ranges to actual data
    Dim nFirstStartRow As Long, nFirstEndRow As Long
    nFirstStartRow = oFirstAddr.StartRow
    nFirstEndRow = oFirstAddr.EndRow
    If nFirstEndRow > nSheetLastRow Then nFirstEndRow = nSheetLastRow

    Dim nSecStartRow As Long, nSecEndRow As Long
    nSecStartRow = oSecondAddr.StartRow
    nSecEndRow = oSecondAddr.EndRow
    If nSecEndRow > nSheetLastRow Then nSecEndRow = nSheetLastRow

    ' Safety check - must have at least some data
    If nFirstEndRow < nFirstStartRow Or nSecEndRow < nSecStartRow Then
        MsgBox "No data found in selected columns.", 48
        Exit Sub
    End If

    ' === Begin locked section ===
    ' All validation is done above. Only data manipulation below.
    ' Use Resume Next so that if anything throws, we still unlock.
    On Local Error Resume Next

    oDoc.lockControllers()

    ' Insert a new column to the right of the second column
    oSheet.getColumns().insertByIndex(nSecondCol + 1, 1)
    ' Note: if nSecondCol < nFirstCol, inserting shifts nFirstCol
    If nSecondCol < nFirstCol Then nFirstCol = nFirstCol + 1

    Dim nResultCol As Long
    nResultCol = nSecondCol + 1

    ' Collect second column values into an array for fast lookup
    Dim aSecondVals() As String
    ReDim aSecondVals(nSecEndRow - nSecStartRow)
    Dim idx As Long
    For idx = nSecStartRow To nSecEndRow
        aSecondVals(idx - nSecStartRow) = oSheet.getCellByPosition(nSecondCol, idx).getString()
    Next idx

    ' Loop through each cell in first column and check for matches
    For idx = nFirstStartRow To nFirstEndRow
        Dim sVal As String
        sVal = oSheet.getCellByPosition(nFirstCol, idx).getString()
        If sVal <> "" Then
            Dim bMatch As Boolean
            bMatch = False
            Dim m As Long
            For m = 0 To UBound(aSecondVals)
                If aSecondVals(m) = sVal Then
                    bMatch = True
                    Exit For
                End If
            Next m

            If bMatch Then
                oSheet.getCellByPosition(nResultCol, idx).setString("TRUE")
            End If
        End If
    Next idx

    ' Add header
    Dim sFirstLetter As String, sSecondLetter As String
    sFirstLetter = ColumnIndexToLetter(nFirstCol)
    sSecondLetter = ColumnIndexToLetter(nSecondCol)
    oSheet.getCellByPosition(nResultCol, 0).setString(sFirstLetter & " in " & sSecondLetter)

    ' Auto-fit result column
    oSheet.getColumns().getByIndex(nResultCol).OptimalWidth = True

    ' === Always unlock, even if errors occurred above ===
    oDoc.unlockControllers()

    On Local Error GoTo 0

    MsgBox "Complete! Matches marked as TRUE.", 64
End Sub
