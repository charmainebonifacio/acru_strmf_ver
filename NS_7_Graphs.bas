Attribute VB_Name = "NS_7_Graphs"
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 6, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateStreamflowGraph
' Description  : This function creates the daily or the monthly
'                streamflow graphs using the observed against
'                simulated data.
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

    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate

    Range("B:B").Select
    ActiveSheet.Shapes.AddChart.Select
    If calIndex = 1 Then
        graphName = "Daily Streamflow Graph"
        ActiveChart.SetSourceData Source:=Range("DailyStats!$B:$B")
    Else
        graphName = "Monthly Streamflow Graph"
        ActiveChart.SetSourceData Source:=Range("MonthlyStats!$B:$B")
    End If
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=graphName
    ActiveChart.Move After:=ActiveWorkbook.Sheets(Sheets.Count)
    With ActiveChart
        .ChartType = xlXYScatter
        .HasLegend = False
        .HasTitle = False
    End With

    ' Series Information
    Dim seriesData As Series
    Set seriesData = ActiveChart.SeriesCollection(1)
    seriesData.Name = "Data"
    seriesData.Values = "=" & tmpSheet.Name & "!C2:C" & LastRow
    seriesData.XValues = "=" & tmpSheet.Name & "!B2:B" & LastRow
    With seriesData
        .MarkerStyle = 8
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerSize = 3
        .Trendlines.Add Type:=xlLinear
        With seriesData.Trendlines(1)
            .Border.Weight = xlMedium
            .Border.LineStyle = xlDash
            .DisplayRSquared = True
            .DisplayEquation = True
        End With
    End With

    ' Set Axis
    'ActiveChart.Axes(xlCategory).MinimumScale = 0
    'ActiveChart.Axes(xlCategory).MaximumScale = maxVal
    'ActiveChart.Axes(xlValue).MinimumScale = 0
    'ActiveChart.Axes(xlValue).MaximumScale = maxVal
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    gridunits = ActiveChart.Axes(xlValue).MajorUnit
    ActiveChart.Axes(xlCategory).MajorUnit = gridunits

    ' Add axis titles
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue).TickLabels.Font.Size = 20
    ActiveChart.Axes(xlCategory).TickLabels.Font.Size = 20
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Simulated Streamflow"
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 28
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Observed Streamflow"
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 28

    ' Add the Trendline
    ActiveChart.SeriesCollection.NewSeries
    Dim trendLine As Series
    Set trendLine = ActiveChart.SeriesCollection(2)
    trendLine.Name = "=""Trendline"""
    trendLine.XValues = "={0," & maxVal & "}"
    trendLine.Values = "={0," & maxVal & "}"
    With trendLine
        .Border.Weight = xlMedium
        .Border.ColorIndex = 1
        .MarkerStyle = -4142
    End With

    ' Add Trendline Information
    txtText = seriesData.Trendlines(1).DataLabel.Text
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            110, 10, 250, 120).TextFrame.Characters.Text = _
            txtText & vbLf & "n = " & WorksheetFunction.Count(Worksheets(tmpShtNum).Range("B:B"))
    ActiveChart.TextBoxes(1).Interior.Color = vbWhite
    ActiveChart.TextBoxes(1).Font.Size = 26
    seriesData.Trendlines(1).DataLabel.Delete

End Function
