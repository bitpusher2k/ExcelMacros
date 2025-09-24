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
' v1.6.1 last updated 2025-09-23
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
' HideEmptyColumns Macro - hides all columns with data only in the first row (which is assumed to be the header row)
    Dim rng As Range
    Dim nLastRow As Long
    Dim nLastColumn As Integer
    Dim i As Integer
    Dim HideIt As Boolean
    Dim j As Long

    Set rng = ActiveSheet.UsedRange
    nLastRow = rng.Rows.Count + rng.row - 1
    nLastColumn = rng.Columns.Count + rng.Column - 1

    For i = 1 To nLastColumn
        HideIt = True

        For j = 2 To nLastRow
            If Not Rows(j).Hidden Then
                If Cells(j, i).Value <> "" Then
                    HideIt = False
                    Exit For
                End If
            End If
        Next

        Columns(i).EntireColumn.Hidden = HideIt
    Next
End Sub


Sub HideGuidColumns()
' HideGuidColumns Macro - Hide all columns with a GUID in the second row (the first is assumed to be the header)
' Be sure to enable "Microsoft VBScript Regular Expression 5.5" under "Tools" > "References..." for this to work.
    Dim cell As Range
    Dim regex As RegExp
    Set regex = New RegExp
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
    For Each cell In ActiveWorkbook.ActiveSheet.Rows("2").Cells
        If regex.Test(cell.Value) Then
            cell.EntireColumn.Hidden = True
        End If
    Next cell
End Sub


Sub HighlightCellsWithSelectedValue()
' HighlightCellsWithSelectedValue Macro - Highlights all cells which contains the selected value
    Dim rCell As Range
    If ActiveCell.Value = vbNullString Then Exit Sub
    Set rCell = ActiveCell
    Do
        Set rCell = ActiveSheet.UsedRange.Cells.Find(ActiveCell.Value, rCell)
        If rCell.Address <> ActiveCell.Address Then
            rCell.Interior.Color = 65535 ' rgbYellow/65535/Yellow
        Else
            Exit Do
        End If
    Loop
End Sub


Sub HighlightRowsWithSelectedValue()
' HighlightRowsWithSelectedValue Macro - Highlights all lines that have a cell which contains the selected value
    Dim rCell As Range
    If ActiveCell.Value = vbNullString Then Exit Sub
    Set rCell = ActiveCell
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Do
        Set rCell = ActiveSheet.UsedRange.Cells.Find(ActiveCell.Value, rCell)
        If rCell.Address <> ActiveCell.Address Then
            rCell.EntireRow.Interior.Color = 7071982 ' rgbPaleGoldenrod/7071982/Pale Goldenrod
            rCell.Interior.Color = 65535 ' rgbYellow/65535/Yellow
            ' Some handy colors:
            ' rgbOrange/42495/Orange
            ' rgbPink/13353215/Pink
            ' rgbYellow/65535/Yellow
            ' rgbOrangeRed/17919/Orange Red
            ' rgbPaleGreen/10025880/Pale Green
            ' rgbPaleGoldenrod/7071982/Pale Goldenrod
            ' rgbLightYellow/14745599/Light Yellow
            ' rgbPaleGreen/10025880/Pale Green
            ' rgbPaleTurquoise/15658671/Pale Turquoise
            ' rgbPaleVioletRed/9662683/Pale Violet Red
        Else
            rCell.EntireRow.Interior.Color = 7071982 ' rgbPaleGoldenrod/7071982/Pale Goldenrod
            rCell.Interior.Color = 65535 ' rgbYellow/65535/Yellow
            Exit Do
        End If
    Loop
End Sub


Sub HighlightRowsWithSelectedValueRed()
' HighlightRowsWithSelectedValue Macro - Highlights all lines that have a cell which contains the selected value
    Dim rCell As Range
    If ActiveCell.Value = vbNullString Then Exit Sub
    Set rCell = ActiveCell
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Do
        Set rCell = ActiveSheet.UsedRange.Cells.Find(ActiveCell.Value, rCell)
        If rCell.Address <> ActiveCell.Address Then
            rCell.EntireRow.Interior.Color = 13353215 ' rgbPink/13353215/Pink
            rCell.Interior.Color = 9662683 ' rgbPaleVioletRed/9662683/Pale Violet Red
        Else
            rCell.EntireRow.Interior.Color = 13353215
            rCell.Interior.Color = 9662683
            Exit Do
        End If
    Loop
End Sub


Sub HighlightRowsWithSelectedValueGreen()
' HighlightRowsWithSelectedValue Macro - Highlights all lines that have a cell which contains the selected value
    Dim rCell As Range
    If ActiveCell.Value = vbNullString Then Exit Sub
    Set rCell = ActiveCell
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Do
        Set rCell = ActiveSheet.UsedRange.Cells.Find(ActiveCell.Value, rCell)
        If rCell.Address <> ActiveCell.Address Then
            rCell.EntireRow.Interior.Color = 10025880 ' rgbPaleGreen/10025880/Pale Green
            rCell.Interior.Color = 15658671 ' rgbPaleTurquoise/15658671/Pale Turquoise
        Else
            rCell.EntireRow.Interior.Color = 10025880
            rCell.Interior.Color = 15658671
            Exit Do
        End If
    Loop
End Sub

Sub HideRowsWithSelectedValue()
' HideRowsWithSelectedValue Macro - Hide all lines that have a cell which contains the selected value
    Dim rCell As Range
    If ActiveCell.Value = vbNullString Then Exit Sub
    Set rCell = ActiveCell
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Do
        Set rCell = ActiveSheet.UsedRange.Cells.Find(ActiveCell.Value, rCell)
        If rCell.Address <> ActiveCell.Address Then
        Else
            rCell.EntireRow.Hidden = True
            Exit Do
        End If
    Loop
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
            myCell.Formula = myCell.Value
        End If
    Next myCell
End Sub


Sub HighlightDuplicateValuesSelected()
' HighlightDuplicateValuesSelected Macro - Highlights duplicate values in selected range
    Dim myRange As Range
    Dim myCell As Range
    Set myRange = Selection
        For Each myCell In myRange
        If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
            myCell.Interior.ColorIndex = 36
        End If
    Next myCell
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
' SaveWorksheetAsXLSX Macro - Saves current worksheet as XLSX with same path & filename
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
    Dim LastCol As Integer
    Set Sheet = ActiveSheet
    LastCol = Sheet.UsedRange.Columns(Sheet.UsedRange.Columns.Count).Column
    For i = LastCol To 1 Step -1
    If Columns(i).Hidden = True Then Columns(i).EntireColumn.Delete
    Next
End Sub


Sub DeleteHiddenRows()
' DeleteHiddenRows Macro - Deletes all hidden rows
    Dim Sheet As Worksheet
    Dim LastRow As Integer
    Set Sheet = ActiveSheet
    LastRow = Sheet.UsedRange.Rows(Sheet.UsedRange.Rows.Count).Row
    For i = LastRow To 1 Step -1
    If Rows(i).Hidden = True Then Rows(i).EntireRow.Delete
    Next
End Sub




Public Sub SplitDateAndTimeToNewColumns()
    Dim MyDateTime As Date
    Dim SelectedColumn As Range
    Dim LastRow As Long
    Dim i As Long

    ' Check if a column is selected
    On Error Resume Next
    Set SelectedColumn = Selection.EntireColumn
    On Error GoTo 0

    If SelectedColumn Is Nothing Then
        MsgBox "Please select a column containing date and time data.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Find the last row in the selected column
    LastRow = Cells(Rows.Count, SelectedColumn.Column).End(xlUp).row

    ' Insert new columns to the right
    SelectedColumn.Offset(, 1).Insert Shift:=xlToRight
    SelectedColumn.Offset(, 2).Insert Shift:=xlToRight

    ' Loop through each row (skip header row)
    For i = 2 To LastRow
        MyDateTime = SelectedColumn.Cells(i, 1).Value

        ' Extract date and time
        SelectedColumn.Cells(i, 2).Value = Int(MyDateTime) ' Date
        SelectedColumn.Cells(i, 3).Value = MyDateTime - Int(MyDateTime) ' Time

        ' Format the new columns
        SelectedColumn.Cells(i, 2).NumberFormat = "YYYY-MM-DD"
        SelectedColumn.Cells(i, 3).NumberFormat = "hh:mm:ss"
    Next i
End Sub
