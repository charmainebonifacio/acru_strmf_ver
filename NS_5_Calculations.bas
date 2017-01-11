Attribute VB_Name = "NS_5_Calculations"
Const Stats1 As String = "(O-P)"
Const Stats2 As String = "(O-P)^2"
Const Stats3 As String = "(O-Oavg)^2"
Const Stats4 As String = "|O-P|"
Const Stats5 As String = "|O-Oavg|"
Const Stats6 As String = "|P-Oavg|"
Const Stats7 As String = "|P-Oavg|+|O-Oavg|"
Const Stats8 As String = "(|P-Oavg|+|O-Oavg|)^2"
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : November 19, 2015
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
    'Worksheets(tmpShtNum - 1).Delete
    Application.DisplayAlerts = True
        
    ' Daily, Worksheet #3
    tmpShtNum = Sheets.Count - 1
    Call NashCalculationsWorksheet(wbMaster, tmpShtNum)

End Function
'---------------------------------------------------------------------
' Date Created : February 23, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : November 22, 2015
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
    Dim LastRow As Long, LastCol As Long
    Dim statsCol As Long, tmpCol As Long
    
    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate
    
    ' Average Calculations
    Call FindLastRowColumn(LastRow, LastCol)
    tmpCol = LastCol + 8 ' Do calculations  from last column + four more columns
    If tmpShtNum > 4 Then
        Range("A1").Offset(0, 0).Value = "UNID"
        Range("A1").Offset(0, 1).Value = "MON_AVE_OBS"
        Range("A1").Offset(0, 2).Value = "MON_AVE_SIM"
    End If
    Range("A1").Offset(0, tmpCol).Value = "TOT_AVE_OBS"
    Range("A1").Offset(1, tmpCol).FormulaR1C1 = "=AVERAGE(RC[-" & LastCol + 7 & "]:R[" & LastRow - 2 & "]C[-" & LastCol + 7 & "])"
    Range("A1").Offset(0, tmpCol + 1).Value = "TOT_AVE_SIM"
    Range("A1").Offset(1, tmpCol + 1).FormulaR1C1 = "=AVERAGE(RC[-" & LastCol + 7 & "]:R[" & LastRow - 2 & "]C[-" & LastCol + 7 & "])"

    ' Statistics
    ' Print Header then Values
    Range("A1").Offset(0, LastCol).Value = Stats1
    Range("A1").Offset(1, LastCol).FormulaR1C1 = "=(RC[-" & LastCol - 1 & "]-RC[-" & LastCol - 2 & "])"
    Range("A1").Offset(1, LastCol).AutoFill Destination:=Range(Cells(2, LastCol + 1), Cells(LastRow, LastCol + 1))
    
    Range("A1").Offset(0, LastCol + 1).Value = Stats2 ' Headers
    Range("A1").Offset(1, LastCol + 1).FormulaR1C1 = "=(RC[-" & LastCol & "]-RC[-" & LastCol - 1 & "])^2"
    Range("A1").Offset(1, LastCol + 1).AutoFill Destination:=Range(Cells(2, LastCol + 2), Cells(LastRow, LastCol + 2))
    
    Range("A1").Offset(0, LastCol + 2).Value = Stats3
    Range("A1").Offset(1, LastCol + 2).FormulaR1C1 = "=(RC[-" & LastCol + 1 & "]-R2C" & tmpCol + 1 & ")^2"
    Range("A1").Offset(1, LastCol + 2).AutoFill Destination:=Range(Cells(2, LastCol + 3), Cells(LastRow, LastCol + 3))
    
    Range("A1").Offset(0, LastCol + 3).Value = Stats4
    Range("A1").Offset(1, LastCol + 3).FormulaR1C1 = "=ABS(RC[-" & LastCol + 2 & "]-RC[-" & LastCol + 1 & "])"
    Range("A1").Offset(1, LastCol + 3).AutoFill Destination:=Range(Cells(2, LastCol + 4), Cells(LastRow, LastCol + 4))
    
    Range("A1").Offset(0, LastCol + 4).Value = Stats5
    Range("A1").Offset(1, LastCol + 4).FormulaR1C1 = "=ABS(RC[-" & LastCol + 3 & "]-R2C" & tmpCol + 1 & ")"
    Range("A1").Offset(1, LastCol + 4).AutoFill Destination:=Range(Cells(2, LastCol + 5), Cells(LastRow, LastCol + 5))
    
    Range("A1").Offset(0, LastCol + 5).Value = Stats6
    Range("A1").Offset(1, LastCol + 5).FormulaR1C1 = "=ABS(RC[-" & LastCol + 3 & "]-R2C" & tmpCol + 1 & ")"
    Range("A1").Offset(1, LastCol + 5).AutoFill Destination:=Range(Cells(2, LastCol + 6), Cells(LastRow, LastCol + 6))
        
    Range("A1").Offset(0, LastCol + 6).Value = Stats7
    Range("A1").Offset(1, LastCol + 6).FormulaR1C1 = "=(RC[-" & LastCol - 1 & "]+RC[-" & LastCol - 2 & "])"
    Range("A1").Offset(1, LastCol + 6).AutoFill Destination:=Range(Cells(2, LastCol + 7), Cells(LastRow, LastCol + 7))

    Range("A1").Offset(0, LastCol + 7).Value = Stats8
    Range("A1").Offset(1, LastCol + 7).FormulaR1C1 = "=(RC[-" & LastCol & "]+RC[-" & LastCol - 1 & "])^2"
    Range("A1").Offset(1, LastCol + 7).AutoFill Destination:=Range(Cells(2, LastCol + 8), Cells(LastRow, LastCol + 8))

    ' Rearrange Columns
    Range(Columns(LastCol + 1), Columns(LastCol + 2)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    statsCol = Range("A1").Offset(0, tmpCol + 1).Column + 2
    Range(Columns(statsCol - 1), Columns(statsCol)).Cut
    Cells(1, LastCol + 1).Select
    ActiveSheet.Paste
    
    ' Change Column Width
    Columns("D:K").ColumnWidth = 15
    Columns("L:L").ColumnWidth = 20
    Columns("M:M").ColumnWidth = 25
    
End Function
