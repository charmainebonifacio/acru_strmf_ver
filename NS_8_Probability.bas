Attribute VB_Name = "NS_8_Probability"
Public lblArr() As String
Public nonlblArr() As String
'---------------------------------------------------------------------
' Date Created : March 10, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 10, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashProbabilityWorksheet
' Description  : This function sets up the Summary Statistics
'                worksheet of the workbook.
' Parameters   : Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashProbabilityWorksheet(ByRef wbMaster As Workbook, _
ByVal DlyLastRow As Long, ByVal MlyLastRow As Long)

    Dim shtIndex As Long

    ' Initialize Arrays
    Call LabelArray
    Call NonLabelArray

    ' Create Daily Probability Worksheet
    shtIndex = 1
    Call ProbabilitySheetLayout(wbMaster, shtIndex, DlyLastRow)

    ' Create Monthly Probability Worksheet
    shtIndex = 2
    Call ProbabilitySheetLayout(wbMaster, shtIndex, MlyLastRow)

End Function
'---------------------------------------------------------------------
' Date Created : March 10, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 10, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LabelArray
' Description  : This function initializes the lblArr array.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function LabelArray()

    ReDim lblArr(0 To 10)
    lblArr(0) = "0.001"
    lblArr(1) = "0.01"
    lblArr(2) = "0.05"
    lblArr(3) = "0.1"
    lblArr(4) = "0.2"
    lblArr(5) = "0.5"
    lblArr(6) = "0.8"
    lblArr(7) = "0.9"
    lblArr(8) = "0.95"
    lblArr(9) = "0.99"
    lblArr(10) = "0.999"

End Function
'---------------------------------------------------------------------
' Date Created : March 10, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 10, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NonLabelArray
' Description  : This function initializes the lblArr array.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function NonLabelArray()

    ReDim nonlblArr(0 To 7)
    nonlblArr(0) = "0.02"
    nonlblArr(1) = "0.03"
    nonlblArr(2) = "0.3"
    nonlblArr(3) = "0.4"
    nonlblArr(4) = "0.6"
    nonlblArr(5) = "0.7"
    nonlblArr(6) = "0.97"
    nonlblArr(7) = "0.98"

End Function
'---------------------------------------------------------------------
' Date Created : March 10, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 5, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ProbabilitySheetLayout
' Description  : This function sets up the Statistics
'                section of the Summary Statistics worksheet.
' Parameters   : Workbook, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function ProbabilitySheetLayout(ByRef wbMaster As Workbook, _
ByVal dataIndex As Long, ByVal lastrow As Long)

    Dim tmpSht As Worksheet
    Dim origSht As Worksheet

    ' Name Worksheet
    wbMaster.Activate
    Set tmpSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    If dataIndex = 1 Then dataName = "Daily Data Probability"
    If dataIndex = 2 Then dataName = "Monthly Data Probability"
    tmpSheet.Name = dataName

    ' Copy OBS/SIM Data
    Set origSht = Sheets(dataIndex + 2)
    origSht.Activate
    Range("B1").EntireColumn.Copy Destination:=tmpSheet.Range("A:A")
    Range("C1").EntireColumn.Copy Destination:=tmpSheet.Range("B:B")

    ' Setup and Cell Alignment
    tmpSheet.Activate
    Cells.Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Rows("1:1").Select
    With Selection
        .Font.Size = 12
        .Font.Bold = True
        .RowHeight = 30
        .WrapText = True
    End With
    Range("A:D, F:H, J:L").ColumnWidth = 14
    Range("E:E,I:I").ColumnWidth = 2

    ' Enter Texts and Format Layout
    If dataIndex = 2 Then
        Range("A1").Offset(0, 0).Value = "OBS"
        Range("A1").Offset(0, 1).Value = "SIM"
    End If
    Range("A1").Offset(0, 2).Value = "RANK"
    Range("A1").Offset(0, 3).Value = "XRANK"

    tmpSheet.Activate
    ActiveSheet.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A2:A" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Sort Values from Largest to Smallest Value (Descending Order)
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("B2:B" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Rank Values from Largest to Smallest Value (Descending Order)
    Range("C2").Value = "1"
    Range("C3").Value = "2"
    Range("C4").Value = "3"
    Range("C2:C4").Select
    Selection.AutoFill Destination:=Range("C2:C" & lastrow), Type:=xlFillDefault
    Range("D2").Select
    If Val(Application.Version) <= 12 Then Selection.FormulaR1C1 = "=NORMSINV((R[0]C[-1]-0.5)/COUNT(R2C2:R" & lastrow & "C2))"
    If Val(Application.Version) > 12 Then Selection.FormulaR1C1 = "=NORM.S.INV((R[0]C[-1]-0.5)/COUNT(R2C2:R" & lastrow & "C2))"
    Selection.AutoFill Destination:=Range("D2:D" & lastrow), Type:=xlFillDefault

    ' Label
    Range("F1").Offset(0, 0).Value = "LABEL"
    Range("F1").Offset(0, 1).Value = "Y-LABEL"
    Range("F1").Offset(0, 2).Value = "X-LABEL"
    For Row = LBound(lblArr) To UBound(lblArr)
        Debug.Print lblArr(Row)
        Range("F1").Offset(Row + 1, 0).Value = lblArr(Row)
        Range("F1").Offset(Row + 1, 1).Value = 0
    Next Row
    Range("H2").Select
    If Val(Application.Version) <= 12 Then Selection.FormulaR1C1 = "=NORMSINV(R[0]C[-2])"
    If Val(Application.Version) > 12 Then Selection.FormulaR1C1 = "=NORM.S.INV(R[0]C[-2])"
    Selection.AutoFill Destination:=Range("H2:H12"), Type:=xlFillDefault

    ' UnLabeled
    Range("J1").Offset(0, 0).Value = "UNLABELED"
    Range("J1").Offset(0, 1).Value = "X-UNLABELED"
    Range("J1").Offset(0, 2).Value = "Y-UNLABELED"
    For Row = LBound(nonlblArr) To UBound(nonlblArr)
        Range("J1").Offset(Row + 1, 0).Value = nonlblArr(Row)
        Range("J1").Offset(Row + 1, 1).Value = 0
    Next Row
    Range("L2").Select
    If Val(Application.Version) <= 12 Then Selection.FormulaR1C1 = "=NORMSINV(R[0]C[-2])"
    If Val(Application.Version) > 12 Then Selection.FormulaR1C1 = "=NORM.S.INV(R[0]C[-2])"
    Selection.AutoFill Destination:=Range("L2:L9"), Type:=xlFillDefault

End Function
