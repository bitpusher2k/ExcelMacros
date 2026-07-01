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
' ExcelMacros.vba - By Bitpusher/The Digital Fox
' v2.5.0 last updated 2026-07-04
' Simple set of useful Excel macros.
'
' Usage:
'
' Activate "Developer" tab in Excel to enable macro manipulation:
' https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45
' Right-click on the ribbon and select "Customize the Ribbon".
' In list "Main Tabs" on the right check the "Developer" box and click OK.
'
' Save desired macros to your "Personal Macro Workbook" so they are available to all workbooks:
' Go to the "Developer" tab in a workbook.
' Click "Record Macro".
' Under "Store macro in" select "Personal Macro Workbook".
' Click "Stop Recording".
' Click the "Visual Basic" button.
' Select VBAProject "PERSONAL.XLSB".
' To use RegEx and the macro "HideGuidColumns()" go to "Tools" > "References...", check "Microsoft VBScript Regular Expression 5.5" and click "OK".
' Expand "Modules" and double-click "Module1"
' Paste desired macros from here and elsewhere into the project and save.
' Workbook "PERSONAL.XLSB" will be created in %appdata%\Microsoft\Excel\XLSTART
'
' Add desired macros as buttons to the ribbon:
' Right-clicking the ribbon > "Customize the Ribbon..."
' "New Tab", rename as desired.
' Create groups, rename as desired.
' Under "Choose commands from:" select "Macros".
' Select desired macros and arrange in group list.
' Rename & select desired icon for each macro-button.
'
' Profit!
'
' Can also place an already created copy of PERSONAL.XLSB into %appdata%\Microsoft\Excel\XLSTART.
' If PERSONAL.XLSB does not load or becomes corrupted delete it from %appdata%\Microsoft\Excel\XLSTART and recreate.
' If PERSONAL.XLSB cannot be loaded from default location a custom location can be defined in "Options" > "Advanced" > "General" > "At startup, open all files in:"
'
' #excel #vba #macro #useful #toolbar #ribbon #autofit #row #column #filter #guid #highlight #selected #blankiferror #formula #value #duplicate #xlsx #pdf

Option Explicit


Sub InitializeCsv()
' InitializeCsv Macro - Applies the AutoFitAllColumns50, AutoFitAllRows50, AddFilter, HideEmptyColumns, and HideGuidColumns macros, then freeze the top row.
    Call AutoFitAllColumns50
    Call AutoFitAllRows50
    Call AddFilter
    Call HideEmptyColumns
    Call HideGuidColumns
    Range("a1").Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub


Sub AutoFitAllColumns50()
' AutoFitAllColumns50 Macro - Auto-fits all column width with maximum with of 50
    Dim mCell As Range
    Application.ScreenUpdating = False

    For Each mCell In ActiveSheet.UsedRange.Rows(1).Cells
        mCell.EntireColumn.AutoFit
        If mCell.EntireColumn.ColumnWidth > 50 Then _
        mCell.EntireColumn.ColumnWidth = 50
    Next mCell
    ActiveSheet.Range("A1").Select

    Application.ScreenUpdating = True
End Sub


Sub AutoFitAllRows50()
' AutoFitAllRows50 Macro - Auto-fits all row height with maximum height of 50
    Dim mCell As Range
    Application.ScreenUpdating = False

    For Each mCell In ActiveSheet.UsedRange.Columns(1).Cells
        mCell.EntireRow.AutoFit
        If mCell.EntireRow.RowHeight > 50 Then _
        mCell.EntireRow.RowHeight = 50
    Next mCell
    ActiveSheet.Range("A1").Select

    Application.ScreenUpdating = True
End Sub


Sub AddFilter()
' AddFilter Macro - Adds filter to first row (ctrl+shift+l)
    ActiveSheet.Range("A1").Select
    Selection.AutoFilter
End Sub


Sub HideEmptyColumns()
' HideEmptyColumns Macro - hides all columns with data only in the first row (assumed header).
' Emptiness is judged across VISIBLE rows only. Fast single-read Variant-array implementation.
    Dim rng As Range
    Dim vals As Variant
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long
    Dim rowVisible() As Boolean
    Dim HideIt As Boolean

    Set rng = ActiveSheet.UsedRange
    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    ' Header only (or empty sheet): nothing to evaluate, leave columns visible
    If nRows < 2 Then Exit Sub

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    vals = rng.Value   ' one COM round-trip instead of nRows*nCols cell reads

    ' Precompute row visibility once (cheap relative to per-cell value reads)
    ReDim rowVisible(1 To nRows)
    For r = 1 To nRows
        rowVisible(r) = Not rng.Rows(r).EntireRow.Hidden
    Next r

    For c = 1 To nCols
        HideIt = True
        For r = 2 To nRows            ' skip header (first row of used range)
            If rowVisible(r) Then
                If IsError(vals(r, c)) Then
                    HideIt = False     ' an error value still counts as data
                    Exit For
                ElseIf Not IsEmpty(vals(r, c)) Then
                    If Len(CStr(vals(r, c))) > 0 Then
                        HideIt = False
                        Exit For
                    End If
                End If
            End If
        Next r
        rng.Columns(c).EntireColumn.Hidden = HideIt
    Next c

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "HideEmptyColumns error: " & Err.Description, vbCritical
End Sub


Sub HideGuidColumns()
' HideGuidColumns Macro - Hides columns whose data looks like a GUID.
' Decides on the first POPULATED data cell in each column (not just row 2),
' so a blank cell in row 2 no longer causes a GUID column to be missed.
' Be sure to enable "Microsoft VBScript Regular Expression 5.5" under "Tools" > "References..." for this to work.
    Const SAMPLE_ROWS As Long = 20
    Dim regex As RegExp
    Dim rng As Range
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long, maxSample As Long
    Dim v As String
    Dim isGuidCol As Boolean

    Set rng = ActiveSheet.UsedRange
    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    If nRows < 2 Then Exit Sub
    maxSample = SAMPLE_ROWS + 1
    If nRows < maxSample Then maxSample = nRows

    Set regex = New RegExp
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    For c = 1 To nCols
        isGuidCol = False
        For r = 2 To maxSample                  ' skip header (first row of used range)
            v = CStr(rng.Cells(r, c).Value)
            If Len(v) > 0 Then
                isGuidCol = regex.Test(v)       ' decide on the first populated cell
                Exit For
            End If
        Next r
        If isGuidCol Then rng.Columns(c).EntireColumn.Hidden = True
    Next c

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "HideGuidColumns error: " & Err.Description, vbCritical
End Sub


' =====================================================================
' Find-based selection macros. All share the ActOnSelectedValue engine.
' Default (unsuffixed) macros use PARTIAL/substring match.
' "...Whole" variants use whole-cell match. All are case-INSENSITIVE.
' Matching is now deterministic: LookIn/LookAt/MatchCase/SearchOrder are
' set explicitly on every Find. Previously these were inherited from the
' last Find dialog/API call, making results session-dependent.
' Ref: https://learn.microsoft.com/en-us/office/vba/api/excel.range.find
' =====================================================================

Private Sub ActOnSelectedValue(ByVal Action As String, _
                               Optional ByVal RowColor As Long = -1, _
                               Optional ByVal CellColor As Long = -1, _
                               Optional ByVal LookAtMode As XlLookAt = xlPart)
' Shared engine. Action: "cell" | "row" | "hide" | "clear".
' Finds every cell in UsedRange matching ActiveCell.Value and applies Action.
    Dim rCell As Range
    If ActiveCell Is Nothing Then Exit Sub
    If ActiveCell.Value = vbNullString Then Exit Sub

    ' Row-scoped actions need all data visible to behave correctly
    If Action <> "cell" Then
        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    End If

    Application.ScreenUpdating = False
    Set rCell = ActiveCell
    Do
        Set rCell = ActiveSheet.UsedRange.Find( _
            What:=ActiveCell.Value, After:=rCell, _
            LookIn:=xlValues, LookAt:=LookAtMode, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False)
        If rCell Is Nothing Then Exit Do

        Select Case Action
            Case "cell": rCell.Interior.Color = CellColor
            Case "row"
                rCell.EntireRow.Interior.Color = RowColor
                rCell.Interior.Color = CellColor
            Case "hide": rCell.EntireRow.Hidden = True
            Case "clear"
                rCell.EntireRow.Interior.ColorIndex = xlNone
                rCell.EntireRow.Font.ColorIndex = xlAutomatic
        End Select

        If rCell.Address = ActiveCell.Address Then Exit Do
    Loop
    Application.ScreenUpdating = True
End Sub


' --- Partial/substring match (default; existing ribbon names preserved) ---
Sub HighlightCellsWithSelectedValue()
' Highlights every cell containing the selected value (Yellow)
    ActOnSelectedValue "cell", , 65535, xlPart
End Sub
Sub HighlightRowsWithSelectedValue()
' Highlights rows containing the selected value (row PaleGoldenrod, cell Yellow)
    ActOnSelectedValue "row", 7071982, 65535, xlPart
End Sub
Sub HighlightRowsWithSelectedValueRed()
' Row Pink, cell PaleVioletRed
    ActOnSelectedValue "row", 13353215, 9662683, xlPart
End Sub
Sub HighlightRowsWithSelectedValueOrange()
' Row OrangeRed, cell Orange
    ActOnSelectedValue "row", 17919, 42495, xlPart
End Sub
Sub HighlightRowsWithSelectedValueGreen()
' Row PaleGreen, cell PaleTurquoise
    ActOnSelectedValue "row", 10025880, 15658671, xlPart
End Sub
Sub HighlightRowsWithSelectedValueCleared()
' Clears highlighting on rows containing the selected value
    ActOnSelectedValue "clear", , , xlPart
End Sub
Sub HideRowsWithSelectedValue()
' Hides rows containing the selected value
    ActOnSelectedValue "hide", , , xlPart
End Sub

' --- Whole-cell match variants ---
Sub HighlightCellsWithSelectedValueWhole()
    ActOnSelectedValue "cell", , 65535, xlWhole
End Sub
Sub HighlightRowsWithSelectedValueWhole()
    ActOnSelectedValue "row", 7071982, 65535, xlWhole
End Sub
Sub HighlightRowsWithSelectedValueRedWhole()
    ActOnSelectedValue "row", 13353215, 9662683, xlWhole
End Sub
Sub HighlightRowsWithSelectedValueOrangeWhole()
    ActOnSelectedValue "row", 17919, 42495, xlWhole
End Sub
Sub HighlightRowsWithSelectedValueGreenWhole()
    ActOnSelectedValue "row", 10025880, 15658671, xlWhole
End Sub
Sub HighlightRowsWithSelectedValueClearedWhole()
    ActOnSelectedValue "clear", , , xlWhole
End Sub
Sub HideRowsWithSelectedValueWhole()
    ActOnSelectedValue "hide", , , xlWhole
End Sub


Public Sub BlankIfError()
' BlankIfError Macro - Surround formulas in all selected cells with =IFERROR(,"")
    Dim row As Long
    Dim Col As Long
    Dim FormulaString As String
    Dim ReadArr As Variant

    If Selection.Cells.Count > 1 Then
        ReadArr = Selection.FormulaR1C1
        For row = LBound(ReadArr, 1) To UBound(ReadArr, 1)
            For Col = LBound(ReadArr, 2) To UBound(ReadArr, 2)
                If Left(ReadArr(row, Col), 1) = "=" Then
                If LCase(Left(ReadArr(row, Col), 8)) <> "=iferror" Then
                    ReadArr(row, Col) = "=iferror(" & Right(ReadArr(row, Col), Len(ReadArr(row, Col)) - 1) & ","""")"
                End If
                End If
            Next
        Next
        Selection.FormulaR1C1 = ReadArr
        Erase ReadArr
    Else
        FormulaString = Selection.FormulaR1C1
        If Left(FormulaString, 1) = "=" Then
            If LCase(Left(FormulaString, 8)) <> "=iferror" Then
                Selection.FormulaR1C1 = "=iferror(" & Right(FormulaString, Len(FormulaString) - 1) & ","""")"
            End If
        End If
    End If
End Sub


Sub ConvertSelectedToValues()
' ConvertSelectedToValues Macro - Converts formulas in selected cells to values
    Dim myRange As Range
    Dim myCell As Range
    Select Case _
        MsgBox("Caution: Action cannot be undone. " _
        & "Save Workbook First?", vbYesNoCancel, _
        "Alert")
        Case Is = vbYes
            ThisWorkbook.Save
        Case Is = vbCancel
            Exit Sub
    End Select
    Set myRange = Selection
    For Each myCell In myRange
        If myCell.HasFormula Then
            myCell.formula = myCell.Value
        End If
    Next myCell
End Sub


Sub HighlightDuplicateValuesSelected()
' HighlightDuplicateValuesSelected Macro - Highlights duplicate values in selected range.
' Two-pass Dictionary implementation (was O(n^2) CountIf-per-cell).
    Dim myCell As Range
    Dim counts As Object
    Dim k As String

    If Selection Is Nothing Then Exit Sub
    Set counts = CreateObject("Scripting.Dictionary")

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ' Pass 1: tally occurrences of each value
    For Each myCell In Selection.Cells
        k = CStr(myCell.Value)
        If counts.Exists(k) Then
            counts(k) = counts(k) + 1
        Else
            counts(k) = 1
        End If
    Next myCell

    ' Pass 2: colour cells whose value appears more than once
    For Each myCell In Selection.Cells
        If counts(CStr(myCell.Value)) > 1 Then
            myCell.Interior.ColorIndex = 36
        End If
    Next myCell

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "HighlightDuplicateValuesSelected error: " & Err.Description, vbCritical
End Sub


Sub AddFrequencyColumn()
' AddFrequencyColumn Macro - Adds column to right of selected column and populates with frequency of values in selected column
    Dim ws As Worksheet
    Dim selCol As Range
    Dim newCol As Range
    Dim dataRange As Range
    Dim LastRow As Long
    Dim colNum As Long
    Dim i As Long
    Dim formula As String

    ' Check if a range is selected
    If Selection Is Nothing Then
        MsgBox "Please select a column containing data.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Get the selected column
    Set ws = ActiveSheet
    Set selCol = Selection.EntireColumn
    colNum = selCol.Column

    ' Find the last row with data in the selected column
    LastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).row

    ' Check if there's data
    If LastRow < 2 Then
        MsgBox "No data found in the selected column.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Insert a new column to the right
    ws.Columns(colNum + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Set the new column range
    Set newCol = ws.Columns(colNum + 1)

    ' Add header (assuming row 1 is header)
    Dim originalHeader As String
    originalHeader = ws.Cells(1, colNum).Value
    If originalHeader = "" Then
        ws.Cells(1, colNum + 1).Value = "Frequency"
    Else
        ws.Cells(1, colNum + 1).Value = originalHeader & "Frequency"
    End If

    ' Define the data range (excluding header)
    Set dataRange = ws.Range(ws.Cells(2, colNum), ws.Cells(LastRow, colNum))

    ' Apply COUNTIF formula to each cell in the new column
    For i = 2 To LastRow
        ' Use absolute reference for the range, relative for the cell being counted
        formula = "=COUNTIF($" & Split(ws.Cells(2, colNum).Address, "$")(1) & _
                  "$2:$" & Split(ws.Cells(2, colNum).Address, "$")(1) & "$" & LastRow & _
                  "," & ws.Cells(i, colNum).Address(False, False) & ")"

        ws.Cells(i, colNum + 1).formula = formula
    Next i

    ' Auto-fit the new column
    newCol.AutoFit

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 Then
        MsgBox "AddFrequencyColumn error: " & Err.Description, vbCritical
    Else
        MsgBox "Frequency column added successfully!", vbInformation
    End If
End Sub


Sub SaveWorkshetAsPDF()
' SaveWorkshetAsPDF Macro - Saves current worksheet as PDF
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo errHandler

    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    strTime = Format(Now(), "yyyymmdd\_hhmm")

    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
        strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"

    'replace spaces and periods in sheet name
    strName = Replace(wsA.Name, " ", "")
    strName = Replace(strName, ".", "_")

    'create default name for savng file
    strFile = strName & "_" & strTime & ".pdf"
    strPathFile = strPath & strFile

    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")

    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        'confirmation message with file info
        MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFile
    End If

exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler
End Sub


Sub SaveWorksheetAsXLSX()
' SaveWorksheetAsXLSX Macro - Saves current worksheet as XLSX with same path & filename (will give error if already exists)
    Dim ActiveFileName, ActiveFilePath, ThisFileName, BaseFileName, NewFullPath As String
    Dim FileNameArray() As String
    Dim FileNameArrayLen As Integer
    ActiveFileName = ActiveWorkbook.Name
    FileNameArray = Split(ActiveFileName, ".")
    FileNameArrayLen = UBound(FileNameArray)
    ReDim Preserve FileNameArray(0 To FileNameArrayLen - 1) As String
    BaseFileName = Join(FileNameArray, ".")
    ActiveFilePath = ActiveWorkbook.Path
    NewFullPath = ActiveFilePath & "\" & BaseFileName & ".xlsx"
    'MsgBox NewFullPath
    ActiveWorkbook.SaveAs Filename:=NewFullPath, FileFormat:=51
End Sub


Sub ClearAllHighlighting()
' ClearAllHighlighting Macro - Clears all highlighting
    Rows.EntireRow.Interior.Color = xlNone
End Sub


Sub UnhideAllRowsColumns()
' UnhideAllRowsColumns Macro - Un-hides all rows and columns
    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False
End Sub


Sub CustomSort()
' CustomSort Macro - Starts the custom sort dialog (save a couple clicks)
    Application.CommandBars.ExecuteMso "SortCustomExcel"
End Sub


Sub DeleteHiddenColumns()
' DeleteHiddenColumns Macro - Deletes all hidden columns
    Dim Sheet As Worksheet
    Dim LastCol As Long
    Dim i As Long
    Set Sheet = ActiveSheet
    LastCol = Sheet.UsedRange.Columns(Sheet.UsedRange.Columns.Count).Column
    For i = LastCol To 1 Step -1
    If Columns(i).Hidden = True Then Columns(i).EntireColumn.Delete
    Next
End Sub


Sub DeleteHiddenRows()
' DeleteHiddenRows Macro - Deletes all hidden rows
    Dim Sheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Set Sheet = ActiveSheet
    LastRow = Sheet.UsedRange.Rows(Sheet.UsedRange.Rows.Count).row
    For i = LastRow To 1 Step -1
    If Rows(i).Hidden = True Then Rows(i).EntireRow.Delete
    Next
End Sub


Sub SplitDateAndTimeToNewColumns()
' SplitDateAndTimeToNewColumns Macro - Splits a selected column of date+time values into
' new "DateOnly" and "TimeOnly" columns to the right. Handles native Excel datetimes and
' common text formats including ISO 8601 (e.g. '2025-01-01T04:27:00Z', '2025-01-01 04:27am',
' '8/24/2023 1:01pm'). Rows that cannot be parsed are left blank and counted rather than
' aborting the whole run.
    Dim SelectedColumn As Range
    Dim LastRow As Long, i As Long, skipped As Long
    Dim dt As Date

    ' Check if a range is selected
    On Error Resume Next
    Set SelectedColumn = Selection.EntireColumn
    On Error GoTo 0
    If SelectedColumn Is Nothing Then
        MsgBox "Please select a cell in a column containing date and time data.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Find the last row in the selected column
    LastRow = Cells(Rows.Count, SelectedColumn.Column).End(xlUp).row
    If LastRow < 2 Then
        MsgBox "No data found in the selected column.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ' Insert new columns to the right
    SelectedColumn.Offset(, 1).Insert Shift:=xlToRight
    SelectedColumn.Offset(, 2).Insert Shift:=xlToRight

    SelectedColumn.Cells(1, 2).Value = "DateOnly"
    SelectedColumn.Cells(1, 3).Value = "TimeOnly"

    ' Loop through each row (skip header row)
    For i = 2 To LastRow
        If TryParseDateTime(SelectedColumn.Cells(i, 1).Value, dt) Then
            SelectedColumn.Cells(i, 2).Value = Int(dt)
            SelectedColumn.Cells(i, 3).Value = dt - Int(dt)
            SelectedColumn.Cells(i, 2).NumberFormat = "YYYY-MM-DD"
            SelectedColumn.Cells(i, 3).NumberFormat = "hh:mm:ss"
        Else
            skipped = skipped + 1
        End If
    Next i

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "SplitDateAndTimeToNewColumns error: " & Err.Description, vbCritical
    ElseIf skipped > 0 Then
        MsgBox "Done. " & skipped & " row(s) could not be parsed as date/time and were left blank.", vbInformation
    End If
End Sub


Private Function TryParseDateTime(ByVal raw As Variant, ByRef dt As Date) As Boolean
' Attempts to coerce a value to a Date. Handles native serials plus common ISO 8601
' text (T separator, trailing Z, +/-HH:MM offsets). Returns False if uninterpretable.
    Dim s As String
    Dim p As Long, tzPlus As Long, tzMinus As Long, firstColon As Long

    TryParseDateTime = False
    If IsEmpty(raw) Then Exit Function
    If IsError(raw) Then Exit Function

    ' Native date or numeric serial
    If IsDate(raw) Then
        dt = CDate(raw)
        TryParseDateTime = True
        Exit Function
    End If

    s = Trim(CStr(raw))
    If Len(s) = 0 Then Exit Function

    ' ISO 8601 normalisation: T -> space, drop trailing Z
    s = Replace(s, "T", " ")
    s = Replace(s, "t", " ")
    If Right(s, 1) = "Z" Or Right(s, 1) = "z" Then s = Left(s, Len(s) - 1)
    s = Trim(s)

    ' Strip a trailing timezone offset (+HH:MM / -HH:MM) that follows the time portion
    p = InStr(s, " ")                       ' start of time portion
    If p > 0 Then
        tzPlus = InStr(p, s, "+")
        If tzPlus > 0 Then s = Trim(Left(s, tzPlus - 1))
        firstColon = InStr(p, s, ":")
        tzMinus = InStrRev(s, "-")
        ' a '-' is only an offset if it sits after the first time colon
        If firstColon > 0 And tzMinus > firstColon Then s = Trim(Left(s, tzMinus - 1))
    End If

    If IsDate(s) Then
        dt = CDate(s)
        TryParseDateTime = True
    End If
End Function


Sub CheckValueMatch()
' CheckValueMatch Macro - Compares each row of one highlighted column with values in second highlighted column and if there is a match marks "true" in a new column to the right of second column - Used for manually combining results of queries into one CSV
    Dim firstCol As Range
    Dim secondCol As Range
    Dim cell As Range
    Dim matchResult As Variant
    Dim resultColumn As Long
    
    On Error GoTo ErrorHandler
    
    ' Check if two ranges are selected
    If Selection.Areas.Count <> 2 Then
        MsgBox "Please select exactly two columns (hold Ctrl to select both).", vbExclamation
        Exit Sub
    End If
    
    ' Assign the selections
    Set firstCol = Selection.Areas(1)
    Set secondCol = Selection.Areas(2)
    
    ' Validate single column selections
    If firstCol.Columns.Count > 1 Or secondCol.Columns.Count > 1 Then
        MsgBox "Please select single columns only.", vbExclamation
        Exit Sub
    End If
    
    ' Insert a new column to the right of second column
    secondCol.EntireColumn.Offset(0, 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Calculate result column (one right of second column)
    resultColumn = secondCol.Column + 1
    
    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each cell in first column
    For Each cell In firstCol
        If cell.Value <> "" Then
            ' Use COUNTIF to check if value exists in second column
            matchResult = Application.WorksheetFunction.CountIf(secondCol, cell.Value)
            
            ' Write TRUE if match found
            If matchResult > 0 Then
                Cells(cell.row, resultColumn).Value = "TRUE"
            Else
                Cells(cell.row, resultColumn).Value = ""
            End If
        End If
    Next cell
    
    ' Get column letters for header
    Dim firstColLetter As String, secondColLetter As String
    firstColLetter = Split(Cells(1, firstCol.Column).Address, "$")(1)
    secondColLetter = Split(Cells(1, secondCol.Column).Address, "$")(1)
    
    ' Add header to result column
    Cells(1, resultColumn).Value = firstColLetter & " in " & secondColLetter
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Auto-fit the result column width
    Columns(resultColumn).AutoFit
    
    MsgBox "Complete! Matches marked as TRUE.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub



