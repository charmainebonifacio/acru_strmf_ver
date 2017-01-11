Attribute VB_Name = "NS_92_TimeSeriesGraphs"
'---------------------------------------------------------------------
' Date Created : December 15, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 4, 2016
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SumMontlhyStreamflowWorksheet
' Description  : This function creates the monthly/annual streamflow
'                worksheet and graph that contains two streamflow
'                variables.
' Parameters   : Workbook
' Returns      : -
'---------------------------------------------------------------------
Function SumMontlhyStreamflowWorksheet(ByRef wbMaster As Workbook, _
ByVal StartYear As String, ByVal EndYear As String)

    Dim shtIndex As Long
    Dim MasterSheet As Worksheet, tmpSheet As Worksheet
    Dim LastRow As Long, LastCol As Long
    Dim headerMean() As String
    Dim pivotTableName As String
    Dim pivotSheet As Worksheet
    Dim newSrcData As String, tblDest As String
    Dim graphName As String
    Dim yAxis As String
    Dim items As Range
    Dim varLabel As String
    Dim SY As Long, EY As Long
    
    varLabel = "AVG_"
    yAxis = "Mean Monthly Streamflow (mm/day)"
    shtIndex = 9
    SY = CInt(StartYear)
    EY = CInt(EndYear)
    
    Set MasterSheet = wbMaster.Worksheets(shtIndex)
    MasterSheet.Activate
     
    ReDim headerMean(5) As String
    headerMean(0) = Range("A1").Value
    headerMean(1) = Range("B1").Value
    headerMean(2) = Range("C1").Value
    headerMean(3) = Range("D1").Value
    headerMean(4) = Trim(Range("E1").Value)
    headerMean(5) = Trim(Range("F1").Value)
    
    ' Change Headers
    Range("E1").Value = headerMean(4)
    Range("F1").Value = headerMean(5)
    
    ' START WITH SEASONALITY
    ' Find Relevant Columns
    Call FindLastRowColumn(LastRow, LastCol)
    Range(Cells(1, 1), Cells(LastRow, LastCol)).Select

    ' Set selected items as current selection
    Set items = Selection
    
    '==============================================================
    ' MONTHLY SEASONAL FLOWS
    ' Create a new pivot table
    pivotTableName = "PivotFlows"
    wbMaster.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=items, _
        Version:=xlPivotTableVersion12).CreatePivotTable _
            TableDestination:="", _
            TableName:=pivotTableName, _
            DefaultVersion:=xlPivotTableVersion12
    ActiveSheet.Name = pivotTableName
    Set pivotSheet = ActiveSheet
    ActiveSheet.Move After:=Sheets(14)

    ActiveSheet.PivotTables(pivotTableName).ManualUpdate = True
    With ActiveSheet.PivotTables(pivotTableName).PivotFields(headerMean(1))
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields(headerMean(4)), varLabel & headerMean(4), xlAverage
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields(headerMean(5)), varLabel & headerMean(5), xlAverage
    ActiveSheet.PivotTables(pivotTableName).ManualUpdate = False
    
   ' Copy Pivot Values into a new worksheet
    pivotSheet.Activate
    Range("A2:C14").Select
    Selection.Copy
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Insert Graph Here
    pivotSheet.Activate
    graphName = "Monthly_Seasonal_Flow_Graph"
    Range("F2:G14").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    With ActiveChart
        .SetSourceData Source:=Range(pivotSheet.Name & "!$F$2:$G$14")
        .Location Where:=xlLocationAsNewSheet, Name:=graphName
    End With
    ActiveChart.Move After:=ActiveWorkbook.Sheets(15)
    
    With ActiveChart.FullSeriesCollection(1)
        .XValues = "={""JAN"",""FEB"",""MAR"",""APR"",""MAY"",""JUN"",""JUL"",""AUG"",""SEP"",""OCT"",""NOV"",""DEC""}"
    End With
    
    With ActiveChart
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = yAxis
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 16
    End With
    
    ' Change Colour Format
    ActiveChart.SeriesCollection(1).Select
    With Selection
        .Name = "OBS"
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection
        .Name = "SIM"
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        '.Format.Line.Weight = 2.25
    End With
    
    ' Legend
    ActiveChart.Legend.Select
    Selection.Position = xlTop
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.Format.TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 18
    End With
    
    ' Chart Title Delete
    ActiveChart.ChartTitle.Select
    Selection.Delete
    
    '==============================================================
    ' ANNUAL FLOW

    pivotSheet.Activate
    ActiveSheet.PivotTables(pivotTableName).ManualUpdate = True
    ActiveSheet.PivotTables(pivotTableName).PivotFields(headerMean(1)).Orientation = _
        xlHidden
    ActiveSheet.PivotTables(pivotTableName).PivotFields(varLabel & headerMean(4)).Orientation = _
        xlHidden
    ActiveSheet.PivotTables(pivotTableName).PivotFields(varLabel & headerMean(5)).Orientation = _
        xlHidden
    varLabel = "SUM_"
    With ActiveSheet.PivotTables(pivotTableName).PivotFields(headerMean(0))
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields(headerMean(4)), varLabel & headerMean(4), xlSum
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields(headerMean(5)), varLabel & headerMean(5), xlSum
    ActiveSheet.PivotTables(pivotTableName).ManualUpdate = False
    
    yAxis = "Streamflow (mm/day)"
   ' Copy Pivot Values into a new worksheet
    pivotSheet.Activate
    Call FindLastRowColumn(LastRow, LastCol)
    Range(Cells(2, 1), Cells(LastRow - 1, 3)).Select

    'Range("A2:C13").Select
    Selection.Copy
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Insert Graph Here
    pivotSheet.Activate
    graphName = "Annual_Seasonal_Flow_Graph"
    LastRow = EY - SY + 1
    Range(Cells(2, 10), Cells(LastRow + 2, 11)).Select
    ActiveSheet.Shapes.AddChart2(227, xlColumnClustered).Select
    With ActiveChart
        .SetSourceData Source:=Range(pivotSheet.Name & "!$J$2:$K$" & LastRow + 2)
        .Location Where:=xlLocationAsNewSheet, Name:=graphName
    End With
    ActiveChart.Move After:=ActiveWorkbook.Sheets(16)

    With ActiveChart
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = yAxis
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 16
    End With
    
    With ActiveChart.FullSeriesCollection(1)
        .XValues = "=" & pivotSheet.Name & "!$I$3:$I$" & LastRow + 2
    End With
    
    ' Change Colour Format
    ActiveChart.SeriesCollection(1).Select
    With Selection
        .Name = "OBS"
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection
        .Name = "SIM"
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        '.Format.Line.Weight = 2.25
    End With
    With Selection
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Fill.Transparency = 0
        .Format.Fill.Solid
    End With
    
    ' Bar Chart
    ActiveChart.ChartGroups(1).Overlap = -25
    ActiveChart.ChartGroups(1).GapWidth = 250
    
    ' Legend
    ActiveChart.Legend.Select
    Selection.Position = xlTop
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.Format.TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 18
    End With
    
    ' Chart Title Delete
    ActiveChart.ChartTitle.Select
    Selection.Delete
    
Cancel:
Set MasterSheet = Nothing
Set pivotSheet = Nothing
End Function
