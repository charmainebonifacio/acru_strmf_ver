Attribute VB_Name = "NS_3_Setup"
'---------------------------------------------------------------------
' Date Created : February 21, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 21, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : InitVarArray
' Description  : This function initializes the two variables that is
'                required to run the Nash SutClif calculations.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function InitVarArray()

    Dim refIndex As Integer
    Dim varNum As Integer

    varNum = 1
    ReDim OutName(0 To varNum)
    OutName(0) = "STRMFL"
    OutName(1) = "CELRUN"

    For refIndex = LBound(OutName) To UBound(OutName)
        Debug.Print OutName(refIndex)
    Next refIndex

End Function
'---------------------------------------------------------------------
' Date Created : July 29, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 29, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Analyze_Multi_ACRU_Out_xxxx
' Description  : This function processes all the ACRU_Out.XXXX files
'                and parses out the HRU number for reference. This
'                function is able to specify if the user would like
'                a different directory for the output file. If not,
'                it uses the input directory.
' Parameters   : Workbook, Worksheet
' Returns      : Boolean
'---------------------------------------------------------------------
Function Analyze_Multi_ACRU_Out_xxxx(ByRef wbMaster As Workbook, _
ByRef MasterSheet As Worksheet, ByRef MasterFile As String) As Boolean

    Dim refIndex As Integer, FileCount As Integer, WrongFileCount As Integer
    Dim vTempVer As String, colExists As Boolean
    Dim formatDate As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Analyze_Multi_ACRU_Out_xxxx = False

    FileCount = 0
    WrongFileCount = 0

    ' Create a Master Workbook
    Set wbMaster = Workbooks.Add(1)
    Set MasterSheet = wbMaster.Worksheets(1)
    MasterSheet.Name = "OriginalData"

    ' Call each file...
    For refIndex = LBound(HRUarr) To UBound(HRUarr)
        FileCount = FileCount + 1
        HRUNUM = HRUarr(refIndex)
        colExists = Setup_ACRU_OUT_XXXX(FileCount, MasterSheet)
        If colExists = False Then WrongFileCount = WrongFileCount + 1
    Next refIndex

    ' Check the variable specified by the user
    If WrongFileCount = 0 Then
        Analyze_Multi_ACRU_Out_xxxx = True
        formatDate = Format(Date, "mm/dd/yyyy")
        OutDate = Replace(formatDate, "/", "")
        MasterFile = "NS_HRU" & HRUNUM & "_RUN" & outRUNVAL & "_" & OutDate
    Else:
        wbMaster.Close SaveChanges:=False
        Set wbMaster = Nothing
        Set MasterSheet = Nothing
    End If

End Function
'---------------------------------------------------------------------
' Date Created : July 29, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 29, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Setup_ACRU_OUT_XXXX
' Description  : This function sets up the ACRU output file. It finds
'                a specific column and copies the values within it.
' Parameters   : Integer, String, Worksheet
' Returns      : Boolean
'---------------------------------------------------------------------
Function Setup_ACRU_OUT_XXXX(ByVal FileCount As Integer, _
ByRef MasterSheet As Worksheet) As Boolean

    Dim wbACRU As Workbook, wsACRU As Worksheet
    Dim refIndex As Integer
    Dim tmpFile As String
    Dim resultVar As Boolean
    Dim arrayText()
    Dim VarToFind As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Setup_ACRU_OUT_XXXX = True

    tmpFile = OutPath & "ACRU_Out." & HRUNUM
    arrayText = Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), _
            Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
            Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), _
            Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), _
            Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
            Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), _
            Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
            Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), _
            Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), _
            Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), _
            Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), _
            Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1), _
            Array(50, 1), Array(51, 1), Array(52, 1))
    Application.StatusBar = "Post-processing File: " & tmpFile
    If FileCount Mod 5 = 0 Then DoEvents
    Workbooks.OpenText _
            FileName:=tmpFile, _
            Origin:=437, _
            StartRow:=1, _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=True, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=arrayText, _
            TrailingMinusNumbers:=True

    Set wbACRU = ActiveWorkbook
    Set wsACRU = wbACRU.Worksheets(1)

    ' Process Headers
    If FileCount = 1 Then Call ColumnHeader(wsACRU, headerArray(), varLastRow, varLastColumn)
    ' Once validation goes through, then setup ACRU files
    If FileCount = 1 Then Call CopyDate(wsACRU, MasterSheet, varLastRow)
    For varInd = LBound(OutName) To UBound(OutName)
        VarToFind = Trim(UCase(OutName(varInd)))
        Call CopyValues(wsACRU, MasterSheet, VarToFind, varLastRow)
    Next varInd
    Range("A1").Select

    ' Save excel spreadsheet
    wbACRU.Close SaveChanges:=False
    Application.StatusBar = False

End Function
'---------------------------------------------------------------------
' Date Created : July 29, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 29, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ColumnHeader
' Description  : This function sets up the column headers for pro-
'                cessing.
' Parameters   : Worksheet, String Array, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function ColumnHeader(ByRef tmpSheet As Worksheet, ByRef headerArray() As String, _
ByRef varLastRow As Long, ByRef varLastColumn As Long)

    Dim rACells As Range, rLoopCells As Range
    Dim refIDCellValue As String
    Dim rowID As Integer
    Dim NewCol As Long
    Dim refIndex As Integer, colIndex As Integer
    Dim CurrentCol As Long

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Initialize Variables
    Call FindLastRowColumn(varLastRow, varLastColumn)
    If WorksheetFunction.CountA(Cells) > 0 Then Range(Cells(1, 1), Cells(1, lastCol)).Select
    Set rACells = Selection

    On Error Resume Next 'In case of NO text constants.

    ' Initialize Array
    NewCol = varLastColumn - 1 ' Because of the header on row one!
    ReDim headerArray(NewCol, 1)

    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If

    'Initializing values in the array to the present AB_ID."
    rowID = 0
    For Each rLoopCells In rACells
        refIDCellValue = rLoopCells.Value
        headerArray(rowID, 0) = UCase(Trim(refIDCellValue))
        CurrentCol = Range(rLoopCells.Address).Column
        headerArray(rowID, 1) = CStr(CurrentCol)
        rowID = rowID + 1
    Next rLoopCells

End Function
'---------------------------------------------------------------------
' Date Created : July 29, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 29, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyDate
' Description  : This function only copies the Y, M and D columns
' Parameters   : Worksheet, Worksheet, Long
' Returns      : -
'---------------------------------------------------------------------
Function CopyDate(SourceSht As Worksheet, DestSht As Worksheet, _
ByVal varLastRow As Long)

    Dim RngSelect
    Dim PasteSelect
    Dim CurrCol As Integer

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Activate Source Worksheet.
    SourceSht.Activate

    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    For refIndex = LBound(headerArray) To UBound(headerArray)
        If headerArray(refIndex, 0) = "YEAR" Then
            CurrCol = CInt(headerArray(refIndex, 1))
            Range(Cells(1, CurrCol), Cells(varLastRow, CurrCol + 2)).Select
            RngSelect = Selection.Address
            Range(RngSelect).Copy

            ' Activate Destination Worksheet.
            DestSht.Activate
            Call ColumnCheck(DestSht)
            PasteSelect = Selection.Address
            Range(PasteSelect).Select
            DestSht.Paste
        Else: Exit For
        End If
    Next refIndex

    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False

End Function
'---------------------------------------------------------------------
' Date Created : July 29, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 29, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyValues
' Description  : This function copies the column that matches the
'                specified variable.
' Parameters   : Worksheet, Worksheet, String, Long
' Returns      : -
'---------------------------------------------------------------------
Function CopyValues(SourceSht As Worksheet, DestSht As Worksheet, _
ByVal VarToFind As String, ByVal varLastRow As Long)

    Dim RngSelect
    Dim PasteSelect
    Dim CurrCol As Integer

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Activate Source Worksheet.
    SourceSht.Activate

    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    For refIndex = LBound(headerArray) To UBound(headerArray)
        If headerArray(refIndex, 0) = VarToFind Then
            CurrCol = CInt(headerArray(refIndex, 1))
            Range(Cells(1, CurrCol), Cells(varLastRow, CurrCol)).Select
            RngSelect = Selection.Address
            Range(RngSelect).Copy

            ' Activate Destination Worksheet.
            DestSht.Activate
            Call ColumnCheck(DestSht)
            PasteSelect = Selection.Address
            Range(PasteSelect).Select
            DestSht.Paste
            Exit For ' When variable is found no use going through the rest!
        End If
    Next refIndex

    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False

End Function
