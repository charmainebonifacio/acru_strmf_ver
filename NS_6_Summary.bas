Attribute VB_Name = "NS_6_Summary"
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 4, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashSummaryWorksheet
' Description  : This function sets up the Summary Statistics
'                worksheet of the workbook.
' Parameters   : Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashSummaryWorksheet(ByRef wbMaster As Workbook, _
ByVal DlyLastRow As Long, ByVal MlyLastRow As Long)

    Dim tmpSht As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim statsCol As Long

    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    tmpSheet.Name = "SummaryStats"
    tmpSheet.Activate

    ' Change Background Colour
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

    ' Setup Column Width
    Rows("1:1").RowHeight = 40
    Columns("A").ColumnWidth = 30
    Columns("B:D").ColumnWidth = 20
    Columns("D").ColumnWidth = 5
    Columns("E").ColumnWidth = 40
    Columns("F:G").ColumnWidth = 20

    ' Enter Texts and Format Layout
    Call SummaryTextLayout(wbMaster, Sheets.Count)
    Call NashTextLayout(wbMaster, Sheets.Count)
    Call AdditonalTextLayout(wbMaster, Sheets.Count)

    ' Enter Calculations for Daily then Monthly
    Call SummaryCalculationsWorksheet(wbMaster, tmpSht, DlyLastRow, 1)
    Call SummaryCalculationsWorksheet(wbMaster, tmpSht, MlyLastRow, 2)

    ' Enter Nash Calculations for Daily then Monthly
    Call NashCalculationsWorksheet(wbMaster, tmpSht, DlyLastRow, 1)
    Call NashCalculationsWorksheet(wbMaster, tmpSht, MlyLastRow, 2)

    ' Number Format
    Cells.Select
    Selection.NumberFormat = "0.00"
    Range("B7:C9").NumberFormat = "General"
    Range("A1").Select

End Function
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 28, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SummaryTextLayout
' Description  : This function sets up the Statistics
'                section of the Summary Statistics worksheet.
' Parameters   : Worksheet, Long
' Returns      : -
'---------------------------------------------------------------------
Function SummaryTextLayout(ByRef wbMaster As Workbook, _
ByRef tmpShtNum As Long)

    Dim tmpSht As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim statsCol As Long

    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate

    ' Enter Header Formatting
    Columns("A:A").Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    Range("B1").Value = "DAILY"
    Range("C1").Value = "MONTHLY"
    Rows("1:1").Select
    With Selection
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' Enter Stats Header Information
    Range("A2").Value = "MEAN"
    Range("A3").Value = "MEAN OBS"
    Range("A4").Value = "MEAN SIM"

    Range("A6").Value = "N"
    Range("A7").Value = "OBS N"
    Range("A8").Value = "SIM N"
    Range("A9").Value = "% Difference"

    Range("A11").Value = "SUM of Q (mm)"
    Range("A12").Value = "Sum OBS Q"
    Range("A13").Value = "Sum SIM Q"
    Range("A14").Value = "MAQ"

    Range("A16").Activate
    With ActiveCell
        .Value = "VARIANCE (mm2)"
        .Characters(Start:=13, Length:=1).Font.Superscript = True
    End With
    Range("A17").Value = "OBS Variance"
    Range("A18").Value = "SIM Variance"
    Range("A19").Value = "% Difference"

    Range("A21").Value = "STANDARD DEVIATION (mm)"
    Range("A22").Value = "OBS STD"
    Range("A23").Value = "SIM STD"
    Range("A24").Value = "% Difference"

    Range("A26").Value = "*GOODNESS OF FIT"
    Range("A27").Value = "Slope of Line"
    Range("A28").Value = "R2"
    Range("A28").Characters(Start:=2, Length:=1).Font.Superscript = True

    ' Enter Formatting for Headers
    Range("A2:A2,A6:A6,A11:A11,A16:A16,A21:A21,A26:A26").Select
    With Selection
        .Font.Size = 8
        .Font.Italic = True
    End With

    ' Enter Formatting for calculations
    Range("B3:B4,B7:B9,B12:B14,B17:B19,B22:B24,B27:B28").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    Range("C3:C4,C7:C9,C12:C14,C17:C19,C22:C24,C27:C28").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0.7
        .PatternTintAndShade = 0
    End With

End Function
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 28, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashTextLayout
' Description  : This function sets up the Nash-Sutcliffe equation
'                section of the Summary Statistics worksheet.
' Parameters   : Worksheet, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashTextLayout(ByRef wbMaster As Workbook, _
ByRef tmpShtNum As Long)

    Dim tmpSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim statsCol As Long

    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate

    ' Enter Header Formatting
    Columns("E:E").Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    Selection.Font.Bold = True
    Range("F1").Value = "DAILY"
    Range("G1").Value = "MONTHLY"

    ' Enter Stats Header Information
    Range("E2").Value = "NASH-SUTCLIFFE COEFFICIENT"
    Range("E3").Value = "SUM of (O-P)^2"
    Range("E4").Value = "SUM of (O-Oavg)^2"
    Range("E5").Value = "1 - [SUM of (O-P)^2 / SUM of (O-Oavg)^2]"

    Range("E7").Value = "**MODIFIED NASH-SUTCLIFFE COEFFICIENT"
    Range("E8").Value = "SUM of |O-P|"
    Range("E9").Value = "SUM of |O-Oavg|"
    Range("E10").Value = "1 - [SUM of |O-P| / SUM of |O-Oavg|]"

    ' Enter Formatting for Headers
    Range("E2:E2,E7:E7").Select
    With Selection
        .Font.Size = 8
        .Font.Italic = True
    End With

    ' Enter Formatting for calculations
    Range("F3:F5,F8:F10").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    Range("G3:G5,G8:G10").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0.7
        .PatternTintAndShade = 0
    End With

End Function
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 1, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : AdditonalTextLayout
' Description  : This function sets up additonal texts associated
'                with the statistics and Nash-Sutcliffe equations.
' Parameters   : Worksheet, Long
' Returns      : -
'---------------------------------------------------------------------
Function AdditonalTextLayout(ByRef wbMaster As Workbook, _
ByRef tmpShtNum As Long)

    Dim tmpSheet As Worksheet
    Dim note1 As String, note2 As String

    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate

    note1 = "*Slope of Line and R2 values in table are correct but the values present on the graph sheet are not correct."
    Range("E13").Value = note1
    Range("E13:G17").Merge

    note2 = "**Modified Nash-Sutcliffe only to be used if regular Nash-Sutcliffe values are bad--See Legates, D.R., & McCabe, G.J. (1999). Evaluating the use of �goodness-of-fit� measures in hydrologic and hydroclimatic model validation. Water Resources Research, 35(1), 233-241."
    Range("E20").Value = note2
    Range("E20:G27").Merge

    Range("E13:G17,E20:G27").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 14
        .Font.Italic = True
        .Font.Bold = False
    End With

End Function
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 1, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SummaryCalculationsWorksheet
' Description  : This function sets up calculations for the
'                Statistics section of the Summary Statistics
'                worksheet.
' Parameters   : Workbook, Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function SummaryCalculationsWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpSht As Worksheet, ByVal shtLastRow As Long, _
ByVal calIndex As Long)

    Dim dlySht As Worksheet, monSht As Worksheet
    Dim startCol As Long
    Dim wkName As String

    ' Activate Appropriate Worksheet #
    Set tmpSht = wbMaster.Worksheets(Sheets.Count)
    Set dlySht = wbMaster.Worksheets(Sheets.Count - 2)
    Set monSht = wbMaster.Worksheets(Sheets.Count - 1)
    tmpSht.Activate

    ' Initialize Variables
    If calIndex = 1 Then
        startCol = 2
        wkName = dlySht.Name
    Else
        startCol = 3
        wkName = monSht.Name
    End If

    Range("A3").Offset(0, calIndex).Value = "=" & wkName & "!R2C4"
    Range("A4").Offset(0, calIndex).Value = "=" & wkName & "!R2C5"

    Range("A7").Offset(0, calIndex).Value = "=Count(" & wkName & "!B2:B" & shtLastRow & ")"
    Range("A8").Offset(0, calIndex).Value = "=Count(" & wkName & "!C2:C" & shtLastRow & ")"
    Range("A9").Offset(0, calIndex).Value = "=(R7C" & startCol & "/R8C" & startCol & "*100)-100"

    Range("A12").Offset(0, calIndex).Value = "=SUM(" & wkName & "!B2:B" & shtLastRow & ")"
    Range("A13").Offset(0, calIndex).Value = "=SUM(" & wkName & "!C2:C" & shtLastRow & ")"
    Range("A14").Offset(0, calIndex).Value = "=(R12C" & startCol & "/R13C" & startCol & "*100)-100"

    Range("A17").Offset(0, calIndex).Value = "=VAR(" & wkName & "!B2:B" & shtLastRow & ")"
    Range("A18").Offset(0, calIndex).Value = "=VAR(" & wkName & "!C2:C" & shtLastRow & ")"
    Range("A19").Offset(0, calIndex).Value = "=(R17C" & startCol & "/R18C" & startCol & "*100)-100"

    Range("A22").Offset(0, calIndex).Value = "=STDEV(" & wkName & "!B2:B" & shtLastRow & ")"
    Range("A23").Offset(0, calIndex).Value = "=STDEV(" & wkName & "!C2:C" & shtLastRow & ")"
    Range("A24").Offset(0, calIndex).Value = "=(R22C" & startCol & "/R23C" & startCol & "*100)-100"

    Range("A27").Offset(0, calIndex).Value = "=SLOPE(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"
    Range("A28").Offset(0, calIndex).Value = "=RSQ(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"

End Function
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 1, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NashCalculationsWorksheet
' Description  : This function sets up calculations for the
'                Nash-Sutcliffe equation section of the Summary
'                Statistics worksheet.
' Parameters   : Workbook, Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function NashCalculationsWorksheet(ByRef wbMaster As Workbook, _
ByRef tmpSht As Worksheet, ByVal shtLastRow As Long, _
ByVal calIndex As Long)

    Dim dlySht As Worksheet, monSht As Worksheet
    Dim startCol As Long
    Dim wkName As String

    ' Activate Appropriate Worksheet #
    Set tmpSht = wbMaster.Worksheets(Sheets.Count)
    Set dlySht = wbMaster.Worksheets(Sheets.Count - 2)
    Set monSht = wbMaster.Worksheets(Sheets.Count - 1)
    tmpSht.Activate

    ' Initialize Variables
    If calIndex = 1 Then
        startCol = 6
        wkName = dlySht.Name
    Else
        startCol = 7
        wkName = monSht.Name
    End If

    Range("E3").Offset(0, calIndex).Value = "=SUM(" & wkName & "!F2:F" & shtLastRow & ")"
    Range("E4").Offset(0, calIndex).Value = "=SUM(" & wkName & "!G2:G" & shtLastRow & ")"
    Range("E5").Offset(0, calIndex).Value = "=1-(R3C" & startCol & "/R4C" & startCol & ")"

    Range("E8").Offset(0, calIndex).Value = "=SUM(" & wkName & "!H2:H" & shtLastRow & ")"
    Range("E9").Offset(0, calIndex).Value = "=SUM(" & wkName & "!I2:I" & shtLastRow & ")"
    Range("E10").Offset(0, calIndex).Value = "=1-(R8C" & startCol & "/R9C" & startCol & ")"

End Function
