Attribute VB_Name = "NS_4_Processing"
'---------------------------------------------------------------------
' Date Created : February 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 7, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashOrigSetupWorksheet
' Description  : This function copies the original data and adds two
'                new columns (DATE and UNID). It also finds the start
'                and end years. Returns by reference the daily last
'                row for the current timeseries.
' Parameters   : Workbook, Worksheet, String, String, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashOrigSetupWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpSheet As Worksheet, ByRef SeriesStartYear As String, _
ByRef SeriesEndYear As String, ByRef DlyLastRow As Long)

    Dim NewStartYearRow As Long
    Dim rngStart As Range, rng As Range
    Dim startRange As Long, LastRow As Long
    Dim newLastCol As Long, newLastRow As Long
    Dim colRange1 As Long, colRange2 As Long
    Dim monthStr As String
    Dim CountMissingValues As Long

    ' Copy Original Data
    wbMaster.Activate
    Worksheets(1).Copy After:=Sheets(Sheets.Count)
    Set tmpSheet = wbMaster.Worksheets(Sheets.Count)
    tmpSheet.Name = "NashData"
    tmpSheet.Activate

    ' Change Headers
    If Range("A1").Offset(0, 1).Value = "MO" Then Range("A1").Offset(0, 1).Value = "MONTH"
    If Range("A1").Offset(0, 2).Value = "DY" Then Range("A1").Offset(0, 2).Value = "DAY"

    ' Insert two new columns: Date and UNID
    ' Insert DATE column
    LastRow = Range("A1").End(xlDown).Row
    Range("A1").Offset(0, 2).Select
    Set rngStart = Selection
    startRange = rngStart.Column
    Range("A1").Offset(0, startRange).Select
    Set rng = Selection
    colRange1 = rng.Column
    Selection.EntireColumn.Insert
    ActiveCell.Value = "DATE"
    Range("A1").Offset(1, startRange).FormulaR1C1 = "=DATE(RC[-3],RC[-2],RC[-1])"
    Range("A1").Offset(1, startRange).AutoFill Destination:=Range(Cells(2, colRange1), Cells(LastRow, colRange1))

    ' Then Insert UNID column
    Range("A1").Offset(0, colRange1).Select
    Set rng = Selection
    colRange2 = rng.Column
    Selection.EntireColumn.Insert
    ActiveCell.Value = "UNID"
    For initRow = 2 To LastRow
        If Len(Cells(initRow, 2)) = 2 Then monthStr = Cells(initRow, 2)
        If Len(Cells(initRow, 2)) < 2 Then monthStr = "0" & Cells(initRow, 2)
        Range("A1").Offset(initRow - 1, colRange1).Value = Cells(initRow, 1) & monthStr
    Next initRow
    Range(Cells(2, colRange2), Cells(LastRow, colRange2)).NumberFormat = "General"

    ' Change Headers
    Range("A1").Offset(0, colRange2).Value = "OBS"
    Range("A1").Offset(0, colRange2 + 1).Value = "SIM"

    ' Look at Series Start and End Years
    ' Once next year row is found, delete the first year data
    Call FindYearRange(wbMaster, colRange1, LastRow, SeriesStartYear, SeriesEndYear)
    NewStartYearRow = FindNashDailyStartYearRow(wbMaster, colRange1, LastRow, SeriesStartYear)
    Range(Cells(2, 1), Cells(NewStartYearRow - 1, 1)).Select
    Selection.EntireRow.Delete
    Range("A1").Select

    ' Remove -99.9 Values and its associated rows
    Call FindLastRowColumn(newLastRow, newLastCol)
    CountMissingValues = 0
    For i = 2 To LastRow
        If Range("A1").Offset(i - 1, newLastCol - 2).Value = -99.9 Then
            CountMissingValues = CountMissingValues + 1
        End If
    Next
    Debug.Print "The NashData Row count: " & newLastRow
    Debug.Print "Missing Values Count: " & CountMissingValues

    ' Auto New Filter
    If CountMissingValues >= 1 Then
        ActiveSheet.Range(Cells(1, newLastCol - 1), _
            Cells(newLastRow, newLastCol - 1)).AutoFilter Field:=1, Criteria1:="=-99.9"
        Call FindLastRowColumn(newLastRow, newLastCol)
        Range(Cells(2, 1), Cells(newLastRow, newLastCol)).Select
        Selection.Delete Shift:=xlUp
        ActiveSheet.Range(Cells(1, newLastCol - 1), _
            Cells(newLastRow, newLastCol - 1)).AutoFilter Field:=1
        Selection.AutoFilter
    End If

    ' Re-calculate Auto New Filter
    Call FindLastRowColumn(newLastRow, newLastCol)
    Debug.Print "After removing all missing values, the NashData Row count: " & newLastRow
    DlyLastRow = newLastRow

End Function
'---------------------------------------------------------------------
' Date Created : February 25, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 25, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindNashDailyStartYearRow
' Description  : This function finds the row number for the next
'                starting year.
' Parameters   : Workbook, Long, Long, String
' Returns      : Long
'---------------------------------------------------------------------
Function FindNashDailyStartYearRow(ByRef wbMaster As Workbook, _
ByVal curCol As Long, ByVal LastRow As Long, _
ByVal SeriesStartYear As String) As Long

    Dim tmpSheet As Worksheet
    Dim StartRow As Long
    Dim DateMonth As Long, DateDay As Long
    Dim SeriesStart As Date, SeriesEnd As Date
    Dim BASEDATE As Date
    Dim newYear As Integer, newStartYear As String

    ' Initialize Variables
    FindNashDailyStartYearRow = 0
    StartRow = 2                        ' After header row
    newYear = CInt(SeriesStartYear) + 1 ' Add one to the current start year
    newStartYear = CStr(newYear)        ' Convert to string
    BASEDATE = DateValue("01/01/" & newStartYear)  ' Create the base date to compare to
    Debug.Print newYear, newStartYear, BASEDATE

    ' Check Original Data
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(Sheets.Count)
    tmpSheet.Activate

    For i = StartRow To LastRow
        SeriesStart = DateValue(Cells(i, curCol).Value)
        DateMonth = DateDiff("m", SeriesStart, BASEDATE)
        DateDay = DateDiff("d", SeriesStart, BASEDATE)
        Debug.Print SeriesStart, DateMonth, DateDay
        If DateMonth = 0 And DateDay = 0 Then
            FindNashDailyStartYearRow = i
            Exit For
        End If
    Next i

End Function
'---------------------------------------------------------------------
' Date Created : February 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 7 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashDataSetupWorksheet
' Description  : This function copies the original data and adds two
'                new columns (DATE and UNID). It also finds the start
'                and end years. Returns by reference the daily new
'                start year row for the current timeseries.
' Parameters   : Workbook, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function NashDataSetupWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpSheet As Worksheet)

    Dim copySheet As Worksheet

    ' Copy to a New Worksheet, Rename
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(Sheets.Count)
    tmpSheet.Copy After:=Sheets(Sheets.Count)
    Set copySheet = wbMaster.Worksheets(Sheets.Count)
    copySheet.Name = "DailyStats"
    copySheet.Activate

    ' Clean Worksheet!
    Columns("E").Delete   ' Delete the UNID
    Columns("D").Copy
    Range("D1").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Selection.ColumnWidth = 13

    ' Delete Unnecessary Columns
    If Range("A1").Offset(0, 0).Value = "YEAR" And _
        Range("A1").Offset(0, 1).Value = "MONTH" And _
        Range("A1").Offset(0, 2).Value = "DAY" Then
        Columns("A:C").Delete ' Delete the YEAR, MONTH, DAY
    End If
    Range("A1").Select

End Function
'---------------------------------------------------------------------
' Date Created : February 25, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 25, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindYearRange
' Description  : This function finds the start and end year for the
'                current ACRU file.
' Parameters   : Workbook, Long, Long, String, String
' Returns      : -
'---------------------------------------------------------------------
Function FindYearRange(ByRef wbMaster As Workbook, ByVal curCol As Long, _
ByVal LastRow As Long, ByRef SeriesStartYear As String, _
ByRef SeriesEndYear As String)

    Dim tmpSheet As Worksheet
    Dim StartRow As Long
    StartRow = 2 ' After header row

    ' Check Original Data
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(Sheets.Count)
    tmpSheet.Activate

    ' Get Start and End Year
    SeriesStartYear = Cells(2, 1).Value
    SeriesEndYear = Cells(LastRow, 1).Value

End Function
'---------------------------------------------------------------------
' Date Created : February 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 1, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashPivotSetupWorksheet
' Description  : This function processes all the ACRU_Out.XXXX files
'                and parses out the HRU number for reference. This
'                function is able to specify if the user would like
'                a different directory for the output file. If not,
'                it uses the input directory.
' Parameters   : Workbook, Worksheet, String, String, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashPivotSetupWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpSheet As Worksheet, ByRef SeriesStartYear As String, _
ByRef SeriesEndYear As String, ByRef MlyLastRow As Long)

    Dim LR As Long, LC As Long, LastRow As Long, newLastRow As Long
    Dim dataArray()
    Dim rngStart As Range, rng As Range
    Dim startRange As Long
    Dim pivotTableName As String
    Dim pivotSheet As Worksheet, MasterSheet As Worksheet
    Dim newSrcData As String, tblDest As String

    wbMaster.Activate
    Set MasterSheet = wbMaster.Worksheets(Sheets.Count - 1)
    MasterSheet.Activate

    ' Find Relevant Columns
    LastRow = Range("A1").End(xlDown).Row
    Call NashDataColumnHeader(ActiveSheet, dataArray(), LR, LC)
    For refIndex = LBound(dataArray) To UBound(dataArray)
        If dataArray(refIndex, 0) = "UNID" Then
            CurrCol = dataArray(refIndex, 1)
            Range(Cells(1, CurrCol), Cells(varLastRow, CurrCol)).Select
            Set rngStart = Selection
            startRange = rngStart.Column
            Exit For ' When variable is found no use going through the rest!
        End If
    Next refIndex

    ' Create a new pivot table
    MasterSheet.Activate
    pivotTableName = "PivotTable"
    Set pivotSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    pivotSheet.Name = pivotTableName
    newSrcData = "'" & MasterSheet.Name & "'!R1C" & startRange & ":R" & LastRow & "C" & startRange + 2
    tblDest = "'" & pivotSheet.Name & "'!R1C1"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=newSrcData, _
        Version:=xlPivotTableVersion12).CreatePivotTable _
            TableDestination:=tblDest, _
            TableName:=pivotTableName, _
            DefaultVersion:=xlPivotTableVersion12
    ActiveSheet.Move After:=Sheets(Sheets.Count)
    'ActiveSheet.Name = pivotTableName

    With ActiveSheet.PivotTables(pivotTableName).PivotFields("UNID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields("OBS"), "Average of OBS", xlAverage
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields("SIM"), "Average of SIM", xlAverage

    ' Copy Pivot Values into a new worksheet
    pivotSheet.Activate
    Columns("A:C").Select
    Selection.Copy
    Set tmpSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    tmpSheet.Name = "MonthlyStats"
    tmpSheet.Activate
    Columns("A:C").ColumnWidth = 17
    Columns("A:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LastRow = Range("B1").End(xlDown).Row
    tmpSheet.Cells(LastRow, 1).EntireRow.Delete
    If Range("A1").Offset(0, 1).Value = "Values" Then Rows(1).EntireRow.Delete
    Range("A1").Select

    ' New Last Row
    newLastRow = Range("A1").End(xlDown).Row
    MlyLastRow = newLastRow

End Function
'---------------------------------------------------------------------
' Date Created : February 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 22, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ColumnHeader
' Description  : This function sets up the column headers for pro-
'                cessing.
' Parameters   : Worksheet, String Array, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashDataColumnHeader(ByRef tmpSheet As Worksheet, ByRef headerArr(), _
ByRef varLastRow As Long, ByRef varLastColumn As Long)

    Dim rACells As Range, rLoopCells As Range
    Dim refIDCellValue As String
    Dim rowID As Integer
    Dim NewCol As Long
    Dim refIndex As Integer, colIndex As Integer
    Dim CurrentCol As Long

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Initialize Variables
    Call FindLastRowColumn(varLastRow, varLastColumn)
    If WorksheetFunction.CountA(Cells) > 0 Then Range(Cells(1, 1), Cells(1, lastCol)).Select
    Set rACells = Selection

    On Error Resume Next 'In case of NO text constants.

    ' Initialize Array
    NewCol = varLastColumn - 1 ' Because of the header on row one!
    ReDim headerArr(NewCol, 1)

    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If

    'Initializing values in the array to the present AB_ID."
    rowID = 0
    For Each rLoopCells In rACells
        refIDCellValue = rLoopCells.Value
        headerArr(rowID, 0) = UCase(Trim(refIDCellValue))
        CurrentCol = Range(rLoopCells.Address).Column
        headerArr(rowID, 1) = CurrentCol
        rowID = rowID + 1
    Next rLoopCells

End Function
