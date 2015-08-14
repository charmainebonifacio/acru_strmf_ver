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
' Parameters   : Workbook, Workbook, Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashProbabilityWorksheet(ByRef wbMaster As Workbook, _
ByRef macroWKBK As Workbook, ByRef macroSht As Worksheet, _
ByVal DlyLastRow As Long, ByVal MlyLastRow As Long)

    Dim dailySheetCount As Long
    Dim monthlySheetCount As Long

    ' Initialize Arrays
    Call LabelArray
    Call NonLabelArray
    dailySheetCount = 1
    monthlySheetCount = 2

    ' Create Daily Probability Worksheet then Monthly Probability Worksheet
    Call ProbabilitySheetLayout(wbMaster, dailySheetCount, DlyLastRow)
    Call ProbabilitySheetLayout(wbMaster, monthlySheetCount, MlyLastRow)

    ' Create Daily and Monthly Probability Graphs
    Call GraphProbabilitySheet(macroWKBK, macroSht, wbMaster, dailySheetCount)
    Call GraphProbabilitySheet(macroWKBK, macroSht, wbMaster, monthlySheetCount)

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
ByVal dataIndex As Long, ByVal LastRow As Long)

    Dim tmpSheet As Worksheet
    Dim origSheet As Worksheet

    ' Name Worksheet
    wbMaster.Activate
    Set tmpSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    If dataIndex = 1 Then dataName = "Daily Data Probability"
    If dataIndex = 2 Then dataName = "Monthly Data Probability"
    tmpSheet.Name = dataName

    ' Copy OBS/SIM Data
    Set origSheet = Sheets(dataIndex + 2)
    origSheet.Activate
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
        .SetRange Range("A2:A" & LastRow)
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
        .SetRange Range("B2:B" & LastRow)
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
    Selection.AutoFill Destination:=Range("C2:C" & LastRow), Type:=xlFillDefault
    Range("D2").Select
    If Val(Application.Version) <= 12 Then Selection.FormulaR1C1 = "=NORMSINV((R[0]C[-1]-0.5)/COUNT(R2C2:R" & LastRow & "C2))"
    If Val(Application.Version) > 12 Then Selection.FormulaR1C1 = "=NORM.S.INV((R[0]C[-1]-0.5)/COUNT(R2C2:R" & LastRow & "C2))"
    Selection.AutoFill Destination:=Range("D2:D" & LastRow), Type:=xlFillDefault

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
'---------------------------------------------------------------------
' Date Created : June 25, 2013
' Created By   : Jacob Palardy
'---------------------------------------------------------------------
' Date Edited  : March 21, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MinValue
' Description  : This program uses the minimum and maximum values
'                from a range to determine acceptable values for a
'                graph minimum and scale units
' Parameters   : Double, Double, String
' Returns      : -
'---------------------------------------------------------------------
Function MinValue(MinVal As Double, GraphUnits As Double, sheetName As String)

    Dim ColMin, ColMax
    Dim graphrange As Double

    ' Initialize Values
    ColMin = Application.WorksheetFunction.Min(Sheets(sheetName).Range("A2:B31"))
    ColMax = Application.WorksheetFunction.Max(Sheets(sheetName).Range("A2:B31"))
    graphrange = (ColMax - ColMin)

    ' Range
    If graphrange <= 10 Then GraphUnits = 1
    If graphrange > 10 And graphrange <= 20 Then GraphUnits = 2
    If graphrange > 20 And graphrange < 50 Then GraphUnits = 5
    If graphrange >= 50 And graphrange < 100 Then GraphUnits = 10
    If graphrange >= 100 And graphrange < 250 Then GraphUnits = 25
    If graphrange >= 250 And graphrange < 500 Then GraphUnits = 50
    If graphrange >= 500 And graphrange < 1000 Then GraphUnits = 100
    If graphrange >= 1000 And graphrange < 3000 Then GraphUnits = 250
    If graphrange >= 3000 Then GraphUnits = 500

    MinVal = (Round(ColMin / GraphUnits) * GraphUnits)
    If MinVal >= ColMin Then
        MinVal = (Round((ColMin - GraphUnits) / GraphUnits) * GraphUnits)
    Else
        MinVal = (Round(ColMin / GraphUnits) * GraphUnits)
    End If
    If GraphUnits Mod 2 = 0 Then
        MinVal = (Round((MinVal - GraphUnits) / (GraphUnits * 2)) * (GraphUnits * 2))
    Else
        If GraphUnits = 25 Then
            MinVal = (Round((MinVal - 25) / 25) * 25)
        End If
    End If
    If MinVal <= 0 Then MinVal = 0

End Function
'---------------------------------------------------------------------
' Date Created : June 25, 2013
' Created By   : Jacob Palardy
'---------------------------------------------------------------------
' Date Edited  : March 25, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : GraphProbabilitySheet
' Description  : This program copies template graph from the Macro
'                workbook and creates a graph for each probability
'                worksheets.
' Parameters   : Workbook, Worksheet, Workbook, Long
' Returns      : -
'---------------------------------------------------------------------
Function GraphProbabilitySheet(ByRef macroWKBK As Workbook, _
ByRef macroSht As Worksheet, ByRef wbMaster As Workbook, _
ByVal sheetCount As Long)

    Dim tmpSheet As Worksheet
    Dim WS_Count As Integer
    Dim xCount As Integer
    Dim i As Integer
    Dim sheetName As String
    Dim newName As String
    Dim MinVal As Double
    Dim GraphUnits As Double
    Dim axisName As String
    Dim LastRow As Integer

    macroWKBK.Activate
    Set tmpSheet = macroWKBK.Worksheets(2)
    tmpSheet.Activate
    ActiveSheet.ChartObjects(1).Copy

    wbMaster.Activate
    If sheetCount = 1 Then
        Sheets(8).Select
        LastRow = Sheets(8).Range("A2").End(xlDown).Row
        sheetName = Sheets(8).Name
        newName = "Daily Probability Graph"
    End If
    If sheetCount = 2 Then
        Sheets(9).Select
        ' Find last row
        LastRow = Sheets(9).Range("A2").End(xlDown).Row
        sheetName = Sheets(9).Name
        newName = "Monthly Probability Graph"
    End If

    Range("I19").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=newName
    Sheets(newName).Move After:=Sheets(Sheets.Count)

    ' For Obs Series
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$1"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$D$2:$D$" & LastRow
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$A$2:$A$" & LastRow
    ActiveChart.SeriesCollection(1).Format.Line.Weight = 1.75

    ' For Sim Series
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$B$1"
    ActiveChart.SeriesCollection(2).XValues = "='" & sheetName & "'!$D$2:$D$" & LastRow
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$2:$B$" & LastRow
    ActiveChart.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
    ActiveChart.SeriesCollection(2).Format.Line.Weight = 1.75

    ' Used to format chart for each unique graph
    MinValue MinVal, GraphUnits, sheetName

    ' Other graph elements
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = MinVal
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    ActiveChart.Axes(xlValue).MajorUnit = GraphUnits
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Streamflow (mm/day)"
    ActiveChart.Axes(xlValue).ScaleType = xlLogarithmic
    ActiveChart.Axes(xlValue).MinimumScale = 0.1

    ' Legend
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Select
    Selection.Position = xlTop
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Arial"
        .NameFarEast = "Arial"
        .Name = "Arial"
        .Bold = msoTrue
        .Size = 24
    End With

    ' Save Original
    macroWKBK.Activate
    Worksheets(1).Activate

End Function
