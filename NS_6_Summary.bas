Attribute VB_Name = "NS_6_Summary"
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : November 22, 2015
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
    Dim LastRow As Long, LastCol As Long
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
    Columns("A").ColumnWidth = 51
    Columns("B:D").ColumnWidth = 15
    Columns("D").ColumnWidth = 5
    Columns("E").ColumnWidth = 50
    Columns("F:G").ColumnWidth = 15

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
    Range("B3:C4").NumberFormat = "General"
    Range("A1").Select
    
End Function
'---------------------------------------------------------------------
' Date Created : February 27, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : November 19, 2015
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
    Dim LastRow As Long, LastCol As Long
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
    Range("A2").Value = "N"
    Range("A3").Value = "OBS N"
    Range("A4").Value = "SIM N"
    
    Range("A6").Value = "MEAN"
    Range("A7").Value = "MEAN OBS"
    Range("A8").Value = "MEAN SIM"
    Range("A9").Value = "% DIFFERENCE"
    
    Range("A11").Value = "SUM OF Q (mm)"
    Range("A12").Value = "SUM OBS Q"
    Range("A13").Value = "SUM SIM Q"
    Range("A14").Value = "% DIFFERENCE"
    
    Range("A16").Activate
    With ActiveCell
        .Value = "VARIANCE (mm2)"
        .Characters(start:=13, Length:=1).Font.Superscript = True
    End With
    Range("A17").Value = "OBS VARIANCE"
    Range("A18").Value = "SIM VARIANCE"
    Range("A19").Value = "% DIFFERENCE"

    Range("A21").Value = "STANDARD DEVIATION (mm)"
    Range("A22").Value = "OBS STD"
    Range("A23").Value = "SIM STD"
    Range("A24").Value = "% DIFFERENCE"
    
    Range("A26").Value = "*STD REGRESSION (GOODNESS OF FIT)"
    Range("A27").Value = "SLOPE OF LINE"
    Range("A27").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The slope indicates the relative relationship between simluated and observed values. A slope of 1 indicates the model perfectly reproduces the magnitudes of observed data."
    End With
    Range("A28").Value = "Y-INTERCEPT"
    Range("A28").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The Y-Intercept indicates the presence of a lag or lead between simulated and observed data. The y-intercept of 0 indicates the model perfectly reproduced the magnitude of the observed data."
    End With
    Range("A29").Value = "COEFFICIENT OF DETERMINATION (R2)"
    Range("A29").Characters(start:=32, Length:=1).Font.Superscript = True
    Range("A29").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The coefficient of determination, r2, als describe the degree of collinearity betweeb simulated and observed data. It ranges from 0 to 1, where high values means less error variance. It is overly-sensitive to extreme high values and the insensitive to the additive and proportional differences between model predictions."
    End With
    Range("A30").Value = "PEARSON CORRELATION COEFFICIENT (r)"
    Range("A30").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The correlation coefficient, r, describe the degree of collinearity between simulated and observed data. It ranges from -1 to 1, where r = 0 means no linear relationship exists, r = 1 or -1 means a perfect positive or negative linear relationship exists. It is overly-sensitive to extreme high values and the insensitive to the additive and proportional differences between model predictions."
    End With
    
    Range("A32").Value = "*ERROR INDEX: RMSE-observed standard ratio (RSR) "
    Range("A32").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The RSR optimal value of 0 indicates zero RMSE or residual variation or a perfect model simulation. The lower the RSR value, the lower the RMSE and indicates a better model performance."
    End With
    Range("A33").Value = "OBSERVED STANDARD DEVIATION (STDobs)"
    Range("A34").Value = "ROOT MEAN SQUARE ERROR (RMSE)"
    Range("A35").Value = "RMSE / STDobs"
    
    Range("A37").Value = "*ERROR INDEX: PERCENT BIAS (PBIAS)"
    Range("A37").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The percent bias measures the average tendency of simulated data to be larger or smaller than the observed data. A value of 0 means accurate model simulation, where positive values indicate model underestimation bias and negative values indicate model overestimation bias."
    End With
    Range("A38").Value = "SUM of OBS"
    Range("A39").Value = "SUM of (O-P) * 100"
    Range("A40").Value = "SUM of (O-P) * 100 / SUM of OBS"
    
    ' Enter Formatting for Headers
    Range("A2:A2,A6:A6,A11:A11,A16:A16,A21:A21,A26:A26,A32:A32,A37:A37").Select
    With Selection
        .Font.Size = 8
        .Font.Italic = True
    End With
    
    ' Enter Formatting for calculations FOR DAILY
    Range("B3:B4,B7:B9,B12:B14,B17:B19,B22:B24,B27:B30,B33:B35,B38:B40").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    
    ' Enter Formatting for calculations FOR MONTHLY
    Range("C3:C4,C7:C9,C12:C14,C17:C19,C22:C24,C27:C30,C33:C35,C38:C40").Select
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
' Date Edited  : November 19, 2015
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
    Dim LastRow As Long, LastCol As Long
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
    Range("E2").Value = "**NASH-SUTCLIFFE EFFICIENCY INDEX"
    Range("E2").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="The Nash-Sutcliffe Efficiency Index is a normalized standard measure that determines the relative magnitude of the residual variance and compared to the measured data variance. Values of 1 is the optimal value; 0 to 1 are acceptable where negative values means unacceptable model performance."
    End With
    Range("E3").Value = "SUM OF (O-P)^2"
    Range("E4").Value = "SUM OF (O-Oavg)^2"
    Range("E5").Value = "1 - [SUM OF (O-P)^2 / SUM OF (O-Oavg)^2]"
    
    Range("E7").Value = "**MODIFIED NASH-SUTCLIFFE EFFICIENCY INDEX"
    Range("E8").Value = "SUM of |O-P|"
    Range("E9").Value = "SUM of |O-Oavg|"
    Range("E10").Value = "1 - [SUM of |O-P| / SUM of |O-Oavg|]"

    Range("E12").Value = "**INDEX OF AGREEMENT"
    Range("E12").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Willmott(1981) developed index of agreement, d, as a standard measure of the degree of model prediction eror. The values range from 0 to 1, where 1 indicates a perfect agreement between the simulated and observed data and 0 indicates no agreement at all."
    End With
    Range("E13").Value = "SUM OF (O-P)^2"
    Range("E14").Value = "SUM OF (|P-Oavg|+|O-Oavg|)^2"
    Range("E15").Value = "1 - [SUM OF (O-P)^2/ SUM OF (|P-Oavg|+|O-Oavg|)^2]"

    Range("E17").Value = "**MODIFIED INDEX OF AGREEMENT"
    Range("E18").Value = "SUM OF |O-P|"
    Range("E19").Value = "SUM OF |P-Oavg|+|O-Oavg|"
    Range("E20").Value = "1 - [SUM OF |O-P|/ SUM OF |P-Oavg|+|O-Oavg|]"

    ' Enter Formatting for Headers
    Range("E2:E2,E7:E7,E12:E12,E17:E17").Select
    With Selection
        .Font.Size = 8
        .Font.Italic = True
    End With
    
    ' Enter Formatting for calculations
    Range("F3:F5,F8:F10,F13:F15,F18:F20").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    Range("G3:G5,G8:G10,G13:G15,G18:G20").Select
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
' Date Edited  : November 19, 2015
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
    Dim note1 As String, note2 As String, note3 As String
    
    ' Activate Appropriate Worksheet #
    wbMaster.Activate
    Set tmpSheet = wbMaster.Worksheets(tmpShtNum)
    tmpSheet.Activate
    
    note1 = "*Statistics are summarized in Moriasi, D., Arnold, J., Van Liew, M., Bingner, R., Harmel, R., & Veith, T. (2007). Model evaluation guidelines for systematic quantification of accuracy in watershed simulations. Trans. ASABE, 50(3), 885-900."
    Range("E22").Value = note1
    Range("E22:G27").Merge
    
    note2 = "**Modified Index of Agreement and Nash Sutcliffe Index --See Krause, P., Boyle, D. P., & Bäse, F. (2005). Comparison of different efficiency criteria for hydrological model assessment. Adv. Geosci., 5, 89-97. doi:10.5194/adgeo-5-89-2005."
    Range("E28").Value = note2
    Range("E28:G33").Merge
    
    note3 = "***Modified Nash-Sutcliffe only to be used if regular Nash-Sutcliffe values are bad--See Legates, D.R., & McCabe, G.J. (1999). Evaluating the use of “goodness-of-fit” measures in hydrologic and hydroclimatic model validation. Water Resources Research, 35(1), 233-241."
    Range("E34").Value = note3
    Range("E34:G39").Merge
    
    Range("E22:G27,E28:G33,E34:G39").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium
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
' Date Edited  : December 13, 2016
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
    
    ' OBS N
    Range("A3").Offset(0, calIndex).Value = "=Count(" & wkName & "!B2:B" & shtLastRow & ")"
    ' SIM N
    Range("A4").Offset(0, calIndex).Value = "=Count(" & wkName & "!C2:C" & shtLastRow & ")"
    ' MEAN OBS
    Range("A7").Offset(0, calIndex).Value = "=" & wkName & "!R2C4"
    ' MEAN SIM
    Range("A8").Offset(0, calIndex).Value = "=" & wkName & "!R2C5"
    ' % DIFFERENCE
    Range("A9").Offset(0, calIndex).Value = "=((100*(R7C" & startCol & "-R8C" & startCol & ")/R7C))"
    ' SUM OBS Q
    Range("A12").Offset(0, calIndex).Value = "=SUM(" & wkName & "!B2:B" & shtLastRow & ")"
    ' SUM SIM Q
    Range("A13").Offset(0, calIndex).Value = "=SUM(" & wkName & "!C2:C" & shtLastRow & ")"
    ' MAQ
    Range("A14").Offset(0, calIndex).Value = "=((100*(R12C" & startCol & "-R13C" & startCol & ")/R12C))"
    ' OBS VARIANCE
    If val(Application.Version) <= 12 Then Range("A17").Offset(0, calIndex).Value = "=VAR(" & wkName & "!B2:B" & shtLastRow & ")"
    If val(Application.Version) > 12 Then Range("A17").Offset(0, calIndex).Value = "=VAR.S(" & wkName & "!B2:B" & shtLastRow & ")"
    ' OBS VARIANCE
    If val(Application.Version) <= 12 Then Range("A18").Offset(0, calIndex).Value = "=VAR(" & wkName & "!C2:C" & shtLastRow & ")"
    If val(Application.Version) > 12 Then Range("A18").Offset(0, calIndex).Value = "=VAR.S(" & wkName & "!C2:C" & shtLastRow & ")"
    ' % DIFFERENCE
    Range("A19").Offset(0, calIndex).Value = "=((100*(R17C" & startCol & "-R18C" & startCol & ")/R17C))"
    ' OBS STD
    If val(Application.Version) <= 12 Then Range("A22").Offset(0, calIndex).Value = "=STDEV(" & wkName & "!B2:B" & shtLastRow & ")"
    If val(Application.Version) > 12 Then Range("A22").Offset(0, calIndex).Value = "=STDEV.S(" & wkName & "!B2:B" & shtLastRow & ")"
    ' SIM STD
    If val(Application.Version) <= 12 Then Range("A23").Offset(0, calIndex).Value = "=STDEV(" & wkName & "!C2:C" & shtLastRow & ")"
    If val(Application.Version) > 12 Then Range("A23").Offset(0, calIndex).Value = "=STDEV.S(" & wkName & "!C2:C" & shtLastRow & ")"
    ' % DIFFERENCE
    Range("A24").Offset(0, calIndex).Value = "=((100*(R22C" & startCol & "-R23C" & startCol & ")/R22C))"
    ' Slope of Line
    Range("A27").Offset(0, calIndex).Value = "=SLOPE(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"
    ' Y-INTERCEPT
    Range("A28").Offset(0, calIndex).Value = "=INTERCEPT(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"
    ' R Squared
    Range("A29").Offset(0, calIndex).Value = "=RSQ(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"
    ' Pearson Correlation Coefficient
    Range("A30").Offset(0, calIndex).Value = "=CORREL(" & wkName & "!C2:C" & shtLastRow & "," & wkName & "!B2:B" & shtLastRow & ")"
    Range("A1").Select

End Function
'---------------------------------------------------------------------
' Date Created : March 1, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : November 22, 2015
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
    Dim startCol As Long, errCol As Long
    Dim wkName As String
    
    ' Activate Appropriate Worksheet #
    Set tmpSht = wbMaster.Worksheets(Sheets.Count)
    Set dlySht = wbMaster.Worksheets(Sheets.Count - 2)
    Set monSht = wbMaster.Worksheets(Sheets.Count - 1)
    tmpSht.Activate

    ' Initialize Variables
    If calIndex = 1 Then
        startCol = 6
        errCol = 2
        wkName = dlySht.Name
    Else
        startCol = 7
        errCol = 3
        wkName = monSht.Name
    End If
    
    ' DIMENSIONLESS STANDARD REGRESSION STATISTICS
    ' Nash Sutcliffe E
    ' SUM of (O-P)^2
    Range("E3").Offset(0, calIndex).Value = "=SUM(" & wkName & "!G2:G" & shtLastRow & ")"
    ' SUM of(O - Oavg) ^ 2
    Range("E4").Offset(0, calIndex).Value = "=SUM(" & wkName & "!H2:H" & shtLastRow & ")"
    ' 1 - [SUM of (O-P)^2 / SUM of (O-Oavg)^2]
    Range("E5").Offset(0, calIndex).Value = "=1-(R3C" & startCol & "/R4C" & startCol & ")"
    
    ' Modified Nash Sutcliffe E1
    ' SUM of |O-P|
    Range("E8").Offset(0, calIndex).Value = "=SUM(" & wkName & "!I2:I" & shtLastRow & ")"
    ' SUM of |O-Oavg|
    Range("E9").Offset(0, calIndex).Value = "=SUM(" & wkName & "!J2:J" & shtLastRow & ")"
    ' 1 - [SUM of |O-P| / SUM of |O-Oavg|]
    Range("E10").Offset(0, calIndex).Value = "=1-(R8C" & startCol & "/R9C" & startCol & ")"
    
    ' Index of Agreement d
    ' SUM of (O-P)^2
    Range("E13").Offset(0, calIndex).Value = "=SUM(" & wkName & "!G2:G" & shtLastRow & ")"
    ' SUM OF (|P-Oavg|+|O-Oavg|)^2
    Range("E14").Offset(0, calIndex).Value = "=SUM(" & wkName & "!M2:M" & shtLastRow & ")"
    ' 1 - [SUM of (O-P)^2 / SUM of (|P-Oavg|+|O-Oavg|)^2]
    Range("E15").Offset(0, calIndex).Value = "=1-(R13C" & startCol & "/R14C" & startCol & ")"
    
    ' Modified Index of Agreement d1
    ' SUM of |O-P|
    Range("E18").Offset(0, calIndex).Value = "=SUM(" & wkName & "!I2:I" & shtLastRow & ")"
    ' SUM of |P-Oavg|+|O-Oavg|
    Range("E19").Offset(0, calIndex).Value = "=SUM(" & wkName & "!L2:L" & shtLastRow & ")"
    ' 1 - [SUM of |O-P| / SUM of |P-Oavg|+|O-Oavg|]
    Range("E20").Offset(0, calIndex).Value = "=1-(R18C" & startCol & "/R19C" & startCol & ")"
    
    ' ERROR INDEX STATISTICS
    ' Observed Standard Deviation
    Range("A33").Offset(0, calIndex).Value = "=(R22C)"
    ' RMSE
    Range("A34").Offset(0, calIndex).Value = "=SQRT(R3C" & startCol & "/R3C" & errCol & ")"
    ' RSR
    Range("A35").Offset(0, calIndex).Value = "=(R34C" & errCol & "/R33C" & errCol & ")"
    
    ' PERCENT BIAS
    ' SUM of OBS
    Range("A38").Offset(0, calIndex).Value = "=SUM(" & wkName & "!B2:B" & shtLastRow & ")"
    ' SUM of (O-P) * 100
    Range("A39").Offset(0, calIndex).Value = "=SUM(" & wkName & "!F2:F" & shtLastRow & ")*100"
    ' SUM of (O-P) * 100 / SUM of OBS
    Range("A40").Offset(0, calIndex).Value = "=R39C/R38C"

End Function
