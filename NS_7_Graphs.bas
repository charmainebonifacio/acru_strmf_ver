Attribute VB_Name = "NS_7_Graphs"
'---------------------------------------------------------------------
' Date Created : March 8, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 8, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateStreamflowGraph
' Description  : This program creates an observed against simulated
'                graph using the data copied
' Parameters   : Workbook, Worksheet, Long, Long, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function CreateStreamflowGraph(ByRef wbMaster As Workbook, _
ByRef tmpSheet As Worksheet, ByVal tmpShtNum As Long, _
ByVal LastRow As Long, ByVal maxVal As Long, ByVal calIndex As Long)

    Dim txtText As String
    Dim gridunits As Integer
    Dim graphName As String
    Dim rng As Range
    Dim graphSheet As Worksheet
    Dim yAxis As String, xAxis As String
    Dim startRange As Range, endRange As Range
    Dim sR As String, eR As String
    
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate
    
    yAxis = "Simulated Streamflow"
    xAxis = "Observed Streamflow"
    
    Set startRange = Worksheets(tmpShtNum).Cells(2, 2)
    sR = startRange.Address
    Set endRange = Worksheets(tmpShtNum).Cells(LastRow, 3)
    eR = endRange.Address
    
    Range(Cells(2, 2), Cells(LastRow, 3)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    If calIndex = 1 Then
        graphName = "Daily Streamflow ScatterPlot"
     '   ActiveChart.SetSourceData Source:=Range("DailyStats!$B:$C")
    Else
        graphName = "Monthly Streamflow ScatterPlot"
     '   ActiveChart.SetSourceData Source:=Range("MonthlyStats!$B:$C")
    End If
    With ActiveChart
        .SetSourceData Source:=Range(tmpSheet.Name & "!" & sR & ":" & eR)
        .Location Where:=xlLocationAsNewSheet, Name:=graphName
    End With
    ActiveChart.Move After:=ActiveWorkbook.Sheets(Sheets.Count)
    
    ' Setup Axis
    With ActiveChart
        .ChartTitle.Delete
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).TickLabelPosition = xlLow
    End With
    
    ' Add Axis titles
    With ActiveChart
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.Font.Size = 16
        .Axes(xlValue, xlPrimary).AxisTitle.Text = yAxis
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 26
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Bold = msoTrue
        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .Axes(xlCategory).TickLabels.Font.Size = 16
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = xAxis
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 26
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = msoTrue
    End With
    
  ' Change Units to be the same for X & Y Axis
    ActiveChart.ChartArea.Select
    yMaxUnit = ActiveChart.Axes(xlValue).MaximumScale
    yMinUnit = ActiveChart.Axes(xlValue).MinimumScale
    xMaxUnit = ActiveChart.Axes(xlCategory).MaximumScale
    xMinUnit = ActiveChart.Axes(xlCategory).MinimumScale
    If yMaxUnit > xMaxUnit Then
        ActiveChart.Axes(xlCategory).MaximumScale = yMaxUnit
        ActiveChart.Axes(xlValue).MaximumScale = yMaxUnit
    Else
        ActiveChart.Axes(xlCategory).MaximumScale = xMaxUnit
        ActiveChart.Axes(xlValue).MaximumScale = xMaxUnit
    End If
    If xMinUnit < yMinUnit Then
        ActiveChart.Axes(xlCategory).MinimumScale = xMinUnit
        ActiveChart.Axes(xlValue).MinimumScale = xMinUnit
    Else
        ActiveChart.Axes(xlCategory).MinimumScale = yMinUnit
        ActiveChart.Axes(xlValue).MinimumScale = yMinUnit
    End If
    
    If yMaxUnit Mod 10 = 0 Then
      ActiveChart.Axes(xlCategory).MajorUnit = 10
      ActiveChart.Axes(xlValue).MajorUnit = 10
    Else
      ActiveChart.Axes(xlCategory).MajorUnit = 5
      ActiveChart.Axes(xlValue).MajorUnit = 5
    End If
    
   ' Change Box Layout
    ActiveChart.PlotArea.Select
    With ActiveChart.PlotArea
        .Left = 120
        .Top = 30
        .Height = 400
        .Width = 400
    End With
    With ActiveChart.PlotArea.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    ' Grid Clean Up
'    ActiveChart.ChartArea.Select
'    ActiveChart.Axes(xlValue).MajorGridlines.Select
'    Selection.Format.Line.Visible = msoFalse
'    ActiveChart.Axes(xlCategory).MajorGridlines.Select
'    Selection.Format.Line.Visible = msoFalse

    ' Series Information
    Dim seriesData As Series
    Set seriesData = ActiveChart.SeriesCollection(1)
    With seriesData
        .MarkerStyle = 8
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerSize = 7
        .Format.Fill.Visible = msoFalse
        .Trendlines.Add Type:=xlLinear
        With seriesData.Trendlines(1)
            .Border.LineStyle = xlDash
            .Format.Line.Visible = msoTrue
            .Format.Line.Weight = 1
            .Format.Line.DashStyle = msoLineDash
            .DisplayRSquared = True
            .DisplayEquation = True
        End With
    End With
    
    ' Add the Trendline
  '  ActiveChart.SeriesCollection.NewSeries
   ' Dim trendLine As Series
    'Set trendLine = ActiveChart.SeriesCollection(2)
    'trendLine.Name = "=""Trendline"""
    'trendLine.XValues = "={0,0}"
    'trendLine.Values = "={0,0}"
    'trendLine.XValues = "={0," & maxVal & "}"
    'trendLine.Values = "={0," & maxVal & "}"
   ' With trendLine
    '    .Border.Weight = xlMedium
     '   .Border.ColorIndex = 1
      '  .MarkerStyle = -4142
    'End With

    ' Add Trendline Information
    ActiveChart.ChartArea.Select
    txtText = ActiveChart.FullSeriesCollection(1).Trendlines(1).DataLabel.Text
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            500, 75, 200, 100).TextFrame.Characters.Text = _
            txtText & vbLf & "N = " & WorksheetFunction.Count(Worksheets(tmpShtNum).Range("B:B"))
'    ActiveChart.TextBoxes(1).Interior.Color = vbWhite
    ActiveChart.TextBoxes(1).Font.Size = 20
    seriesData.Trendlines(1).DataLabel.Delete
    
End Function


