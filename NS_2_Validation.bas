Attribute VB_Name = "NS_2_Validation"
'---------------------------------------------------------------------
' Date Created : August 2 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 14, 2015
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ValidateDirectory
' Description  : This function validates user input in order to find
'                the reference table before setting up the files.
' Parameters   : -
' Returns      : Boolean
'---------------------------------------------------------------------
Function ValidateDirectory() As Boolean

    Dim openFile As Variant
    Dim fileCheck As Boolean
    Dim filePath As String, vResponse As String
    Dim Sep As String
    Dim tableResult As Boolean
    Dim MessageSummary As String, SummaryTitle As String

    Application.ScreenUpdating = False
    ValidateDirectory = True
    tableResult = True
    fileCheck = True
    SummaryTitle = "ACRU POST-PROCESSING TOOL: EXPORT VARIABLE"

    '-------------------------------------------------------------
    ' Select Multiple ACRU Output files to be processed.
    '-------------------------------------------------------------
    openFile = Application.GetOpenFilename( _
        filefilter:="ACRU OUTPUT (*.*), *.*", _
        Title:="Open ACRU OUTPUT Files", MultiSelect:=True)
    If TypeName(openFile) = "Boolean" Then GoTo Cancel '"User has cancelled."

    '-------------------------------------------------------------
    ' Setup ACRU output files. If user selected non-ACRU output,
    ' then end function
    '-------------------------------------------------------------
    fileCheck = ACRU_OUTFILE_Selection(openFile, filePath, HRUarr()) ' Sorted array!
    If fileCheck = False Then GoTo Cancel
    If fileCheck = True Then
        ' Setup InPath and OutPath
        InPath = ReturnFolder(filePath)
        vResponse = MsgBox("Would you like to select an output directory?", vbYesNo)
        If vResponse = vbYes Then
            OutPath = GetFolder
            OutPath = ReturnFolder(OutPath)
        Else: OutPath = InPath ' Simply use the current file directory
        End If
    End If

Cancel:
    If TypeName(openFile) = "Boolean" Then
        ValidateDirectory = False
    End If
    If fileCheck = False Then
        ValidateDirectory = fileCheck
        MessageSummary = MacroCancel(1)
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
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
' Title        : ACRU_OUTFILE_Selection
' Description  : This function allows the user to select multiple
'                ACRU_Out files. It parses out the HRU number that
'                will be used later. If user selects a non-ACRU
'                file then the function will return a False.
' Parameters   : Variant, String Array
' Returns      : Boolean
'---------------------------------------------------------------------
Function ACRU_OUTFILE_Selection(ByRef openFile As Variant, _
ByRef filePath As String, ByRef HRUarr() As String) As Boolean

    Dim Txtfile As String
    Dim ACRU As String, HRUname As String, ACRUOUT As String
    Dim Sep As String
    Dim Top As Integer, Bottom As Integer
    Dim FileCounter As Integer, refIndex As Integer

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    ACRU_OUTFILE_Selection = False

    ACRUOUT = "ACRU_Out"
    If TypeName(openFile) = "Variant()" Then FileCounter = 0

    '-------------------------------------------------------------
    ' Extract the appropriate file names...
    '-------------------------------------------------------------
    FileCounter = 0
    For Each ACRUFile In openFile
        FileCounter = FileCounter + 1
        Sep = InStrRev(ACRUFile, "\")
        filePath = Left(ACRUFile, Sep - 1)
        Txtfile = Mid(ACRUFile, Sep + 1)
        Sep = InStrRev(Txtfile, ".")
        ACRU = Left(Txtfile, Sep - 1)
        If ACRU = ACRUOUT Then
            HRUname = Mid(Txtfile, Sep + 1)
            ReDim Preserve HRUarr(FileCounter - 1)
            HRUarr(FileCounter - 1) = HRUname
        Else: Exit Function ' Don't sort array
        End If
    Next

    ' Sort Array!
    Top = UBound(HRUarr)
    Bottom = LBound(HRUarr)
    Call QuickSort(HRUarr(), Bottom, Top)
    ACRU_OUTFILE_Selection = True

Cancel:
End Function
'---------------------------------------------------------------------
' Date Acquired: August 2, 2013
' Source       : http://www.blueclaw-db.com/quick-sort.htm
'---------------------------------------------------------------------
' Date Edited  : August 2, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : QuickSort
' Description  : This function sorts the array that contains all the
'                HRU
' Parameters   : String Array, Integer, Integer
' Returns      : -
'---------------------------------------------------------------------
Function QuickSort(strArray() As String, intBottom As Integer, intTop As Integer)

    Dim strPivot As String, strTemp As String
    Dim intBottomTemp As Integer, intTopTemp As Integer

    intBottomTemp = intBottom
    intTopTemp = intTop

    strPivot = strArray((intBottom + intTop) \ 2)

    While (intBottomTemp <= intTopTemp)
        '  comparison of the values is a descending sort
        While (strArray(intBottomTemp) < strPivot And intBottomTemp < intTop)
            intBottomTemp = intBottomTemp + 1
        Wend
        While (strPivot < strArray(intTopTemp) And intTopTemp > intBottom)
            intTopTemp = intTopTemp - 1
        Wend
        If intBottomTemp < intTopTemp Then
            strTemp = strArray(intBottomTemp)
            strArray(intBottomTemp) = strArray(intTopTemp)
            strArray(intTopTemp) = strTemp
        End If
        If intBottomTemp <= intTopTemp Then
            intBottomTemp = intBottomTemp + 1
            intTopTemp = intTopTemp - 1
        End If
    Wend

    'the function calls itself until everything is in good order
    If (intBottom < intTopTemp) Then QuickSort strArray, intBottom, intTopTemp
    If (intBottomTemp < intTop) Then QuickSort strArray, intBottomTemp, intTop

End Function
