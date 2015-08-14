Attribute VB_Name = "NS_5_Calculations"
Const Stats1 As String = "(O-P)^2"
Const Stats2 As String = "(O-Oavg)^2"
Const Stats3 As String = "|O-P|"
Const Stats4 As String = "|O-Oavg|"
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 7, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashSetupCalculationsWorksheet
' Description  : This function sets up the calculations for both
'                daily and monthly worksheets.
' Parameters   : Worksheet, Long
' Returns      : -
'--------------------------------------------------------------------
Function NashSetupCalculationsWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpShtNum As Long)

    wbMaster.Activate

    ' Monthly, use Pivot Table Worksheet
    tmpShtNum = Sheets.Count
    Call NashCalculationsWorksheet(wbMaster, tmpShtNum)

    ' Delete Pivot Table
    Application.DisplayAlerts = False
    Worksheets(tmpShtNum - 1).Delete
    Application.DisplayAlerts = True

    ' Daily, Worksheet #3
    tmpShtNum = Sheets.Count - 1
    Call NashCalculationsWorksheet(wbMaster, tmpShtNum)

End Function
'---------------------------------------------------------------------
' Date Created : February 23, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 1, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashCalculationsWorksheet
' Description  : This function sets up the calculations for four
'                equations;  "(O-P)^2", "(O-Oavg)^2", "|O-P|" and
'                "|O-Oavg|".
' Parameters   : Worksheet, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashCalculationsWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpShtNum As Long)

    Dim tmpSheet As Worksheet
    Dim LastRow As Long, lastCol As Long
    Dim statsCol As Long, tmpCol As Long

    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate

    ' Average Calculations
    Call FindLastRowColumn(LastRow, lastCol)
    tmpCol = lastCol + 4 ' Do calculations  from last column + four more columns
    Columns("D:I").ColumnWidth = 17
    If tmpShtNum > 3 Then
        Range("A1").Offset(0, 0).Value = "UNID"
        Range("A1").Offset(0, 1).Value = "MON_AVE_OBS"
        Range("A1").Offset(0, 2).Value = "MON_AVE_SIM"
    End If
    Range("A1").Offset(0, tmpCol).Value = "TOT_AVE_OBS"
    Range("A1").Offset(1, tmpCol).FormulaR1C1 = "=AVERAGE(RC[-" & lastCol + 3 & "]:R[" & LastRow - 2 & "]C[-" & lastCol + 3 & "])"
    Range("A1").Offset(0, tmpCol + 1).Value = "TOT_AVE_SIM"
    Range("A1").Offset(1, tmpCol + 1).FormulaR1C1 = "=AVERAGE(RC[-" & lastCol + 3 & "]:R[" & LastRow - 2 & "]C[-" & lastCol + 3 & "])"

    ' Statistics
    ' Print Header then Values
    Range("A1").Offset(0, lastCol).Value = Stats1 ' Headers
    Range("A1").Offset(1, lastCol).FormulaR1C1 = "=(RC[-" & lastCol - 1 & "]-RC[-" & lastCol - 2 & "])^2"
    Range("A1").Offset(1, lastCol).AutoFill Destination:=Range(Cells(2, lastCol + 1), Cells(LastRow, lastCol + 1))

    Range("A1").Offset(0, lastCol + 1).Value = Stats2
    Range("A1").Offset(1, lastCol + 1).FormulaR1C1 = "=(RC[-" & lastCol & "]-R2C" & tmpCol + 1 & ")^2"
    Range("A1").Offset(1, lastCol + 1).AutoFill Destination:=Range(Cells(2, lastCol + 2), Cells(LastRow, lastCol + 2))

    Range("A1").Offset(0, lastCol + 2).Value = Stats3
    Range("A1").Offset(1, lastCol + 2).FormulaR1C1 = "=ABS(RC[-" & lastCol + 1 & "]-RC[-" & lastCol & "])"
    Range("A1").Offset(1, lastCol + 2).AutoFill Destination:=Range(Cells(2, lastCol + 3), Cells(LastRow, lastCol + 3))

    Range("A1").Offset(0, lastCol + 3).Value = Stats4
    Range("A1").Offset(1, lastCol + 3).FormulaR1C1 = "=ABS(RC[-" & lastCol + 2 & "]-R2C" & tmpCol + 1 & ")"
    Range("A1").Offset(1, lastCol + 3).AutoFill Destination:=Range(Cells(2, lastCol + 4), Cells(LastRow, lastCol + 4))

    ' Rearrange Columns
    Range(Columns(lastCol + 1), Columns(lastCol + 2)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    statsCol = Range("A1").Offset(0, tmpCol + 1).Column + 2
    Range(Columns(statsCol - 1), Columns(statsCol)).Cut
    Cells(1, lastCol + 1).Select
    ActiveSheet.Paste

End Function
