Attribute VB_Name = "NS_1_MAIN"
Public HRUNUM As String
Public outRUNVAL As String
Public OutPath As String
Public OutDate As String
Public OutName() As String
Public HRU As Integer
Public Counter As Integer
Public HRUarr() As String
Public headerArray() As String
Public areaArray() As Double
Public varLastRow As Long
Public varLastColumn As Long
'---------------------------------------------------------------------
' Date Created : February 21, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 21, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : UserFormInitialize
' Description  : This function will initialize the userform.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Sub UserFormInitialize()

    UserForm1.Show

End Sub
'---------------------------------------------------------------------
' Date Created : February 21, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 4, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : NASHSUTCLIFF_MAIN
' Description  : This function will run an area weight analysis on a
'                specific variable from ACRU output file.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function NASHSUTCLIFF_MAIN(ByVal outRUNVAL As String)

    Dim wbMaster As Workbook, MasterSheet As Worksheet
    Dim tmpSheet As Worksheet, tmpSheetNum As Long
    Dim MasterFile As String, OutFileName As String
    Dim StartYear As String, EndYear As String
    Dim DailyLastRow As Long, MonthlyLastRow As Long
    Dim inputMaxAxis As Long
    Dim valResult As Boolean, acruFileResult As Boolean
    Dim wbResult As Boolean, wbExists As Boolean
    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String

    UserForm1.Hide
    Application.ScreenUpdating = False

    acruFileResult = True
    SummaryTitle = "Nash SutCliff POST-PROCESSING Summary: "
    start_time = Now()

    '-------------------------------------------------------------
    ' Initialize variables to find into an array
    '-------------------------------------------------------------
    Call InitVarArray

    '-------------------------------------------------------------
    ' Validate User Input
    '-------------------------------------------------------------
    valResult = ValidateDirectory
    If valResult = False Then GoTo Cancel

    '-------------------------------------------------------------
    ' New Workbook (wbMaster) will be added that contains one
    ' worksheet names "Original" (Mastersheet). Worksheet #1
    ' Setup the selected ACRU file. Then find & copy values for
    ' two particular variables: CELRUN, STRMFL
    '-------------------------------------------------------------
    acruFileResult = Analyze_Multi_ACRU_Out_xxxx(wbMaster, MasterSheet, MasterFile)
    '  Application.StatusBar = "Finished post-processing the selected ACRU output files."
    If acruFileResult = False Then GoTo Cancel

    '-------------------------------------------------------------
    ' Process Original Data. Worksheet #2
    ' Find the next Start Year and the associated row number
    '-------------------------------------------------------------
    Call NashOrigSetupWorksheet(wbMaster, tmpSheet, StartYear, EndYear, DailyLastRow)
    ' Application.StatusBar = "Finished processing the original ACRU data for Nash SutCliff calculations."
    Set tmpSheet = Nothing

    '-------------------------------------------------------------
    ' Copy Original Data but only keep DATE, OBS and SIM. Worksheet #3
    '-------------------------------------------------------------
    Call NashDataSetupWorksheet(wbMaster, tmpSheet)
    Set tmpSheet = Nothing

    '-------------------------------------------------------------
    ' Create Pivot Table for Monthly Calculations. Worksheet #4
    ' Find the next Start Year and the associated row number
    '-------------------------------------------------------------
    Call NashPivotSetupWorksheet(wbMaster, tmpSheet, StartYear, EndYear, MonthlyLastRow)
    ' Application.StatusBar = "Finished creating pivot table on ACRU data for Nash SutCliff calculations."
    Set tmpSheet = Nothing

    '-------------------------------------------------------------
    ' Start Monthly and Daily Calculations. Worksheet #5 & #3
    ' Delete Pivot Table for simpler calculations
    '-------------------------------------------------------------
    Call NashSetupCalculationsWorksheet(wbMaster, tmpSheetNum)

    '-------------------------------------------------------------
    ' Summarize Calculations with Statistics: Worksheet #7
    ' Daily and Monthly Nash Sutcliff
    '-------------------------------------------------------------
    Call NashSummaryWorksheet(wbMaster, DailyLastRow, MonthlyLastRow)

    '-------------------------------------------------------------
    ' Create Streamflow Graphs
    ' Daily and Monthly (Worksheet #6, #7)
    '-------------------------------------------------------------
    inputMaxAxis = 0 ' InputBox("Set Axis Maximums")
    Call CreateStreamflowGraph(wbMaster, tmpSheet, 3, _
        DailyLastRow, _
        inputMaxAxis, 1)

    Call CreateStreamflowGraph(wbMaster, tmpSheet, 4, _
        MonthlyLastRow, _
        inputMaxAxis, 2)

    '-------------------------------------------------------------
    ' Create Daily Data Probability Worksheet and Graph
    ' Worksheet #8, #9
    '-------------------------------------------------------------

    '-------------------------------------------------------------
    ' Create Monthly Data Probability Worksheet and Graph
    ' Worksheet #10, #11
    '-------------------------------------------------------------

    '-------------------------------------------------------------
    ' Save Workbook and all the progress as follows:
    ' NS_HRU####_RUN##_MMDDYYYY.xlsx
    '-------------------------------------------------------------
    wbResult = IsFileOpen(MasterFile)
    If wbResult = True Then CheckWorkBook (MasterFile)
    wbExists = CheckFileExists(OutPath, MasterFile, ".xlsx")
    If wbExists = True Then MasterFile = ChangeName(wbExists, OutPath, MasterFile) ' Change MasterFile
    OutFileName = SaveReturnXLSX(wbMaster, MasterSheet, OutPath, MasterFile)
    wbMaster.Close SaveChanges:=False

    end_time = Now()
    ProcessingTime = DateDiff("n", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime, OutFileName)
    MsgBox MessageSummary, vbOKOnly, SummaryTitle

    Application.StatusBar = False
    Application.ScreenUpdating = True

Cancel:
    If acruFileResult = False Then
        MessageSummary = MacroCancel(3)
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
    End If
    Set wbMaster = Nothing
    Set MasterSheet = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : July 31, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 31, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : IsFileOpen
' Description  : This function will check the status of the file.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function IsFileOpen(ByVal MasterFile As String) As Boolean

    Dim iFilenum As Long
    Dim iErr As Long
    Dim wbTMP As Workbook

    On Error Resume Next

    iFilenum = FreeFile()
    Open FileName For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = Err

    Set wbTMP = Workbooks(MasterFile)

    On Error GoTo 0

    Select Case iErr
        Case 0:  IsFileOpen = False ' Closed
        Case 70: IsFileOpen = True ' Opened
        Case 75: IsFileOpen = True ' Read Only
        Case Else: Error iErr
    End Select

End Function
'---------------------------------------------------------------------------------------
' Date Created : July 31, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 31, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : WorkBookCheck
' Description  : This function will re-open only if it is already open. Otherwise, this
'                function will not be invoked at all.
' Parameters   : String
' Returns      : -
'---------------------------------------------------------------------------------------
Function CheckWorkBook(ByVal MasterFile As String)

    Dim WbookCheck As Workbook

    On Error Resume Next
    Set WbookCheck = Workbooks.Open(MasterFile)
    On Error GoTo 0

    If WbookCheck Is Nothing Then 'Closed
        Debug.Print "Closed"
    ElseIf Application.ActiveWorkbook.Name = WbookCheck.Name Then
        WbookCheck.Close SaveChanges:=True
    End If

End Function
