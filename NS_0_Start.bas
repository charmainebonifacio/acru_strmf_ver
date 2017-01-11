Attribute VB_Name = "NS_0_Start"
'---------------------------------------------------------------------
' Date Created : September 7, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : September 7, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Start_Here
' Description  : The purpose of function is to initialize the userform.
'---------------------------------------------------------------------
Sub Start_Here()
       
    Dim myForm As UserForm1
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String
    Dim strLabel3 As String, strLabel4 As String
    Dim strLabel5 As String, strLabel6 As String
    Dim strLabel7 As String, strLabel8 As String
    Dim frameLabel1 As String, frameLabel2 As String, frameLabel3 As String
    Dim userFormCaption As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Set myForm = UserForm1
    
    ' Label Strings
    userFormCaption = "  KIENZLE LAB TOOLS"
    button1 = "SELECT ACRU OUTPUT FILE"
    frameLabel2 = "TOOL GUIDE"
    frameLabel3 = "HELP SECTION"
    
    strLabel1 = "THE ACRU STREAMFLOW VERIFICATION MACRO"
    strLabel2 = "STEP 1."
    strLabel3 = "Run ACRU Model. Find the specific ACRU OUT FILE." & vbLf
    strLabel4 = "STEP 3."
    strLabel5 = "For more information, hover mouse over button."
    strLabel7 = "STEP 2."
    strLabel8 = "Enter RUN Number (##):" & vbLf
    
    ' UserForm Initialize
    myForm.Caption = userFormCaption
    myForm.Frame2.Caption = frameLabel2
    myForm.Frame5.Caption = frameLabel3
    myForm.Frame2.Font.Bold = True
    myForm.Frame5.Font.Bold = True
    myForm.Label1.Caption = strLabel1
    myForm.Label1.Font.Size = 21
    myForm.Label1.Font.Bold = True
    myForm.Label1.TextAlign = fmTextAlignCenter
    
    myForm.Label2 = strLabel2
    myForm.Label2.Font.Size = 13
    myForm.Label2.Font.Bold = True
    myForm.Label7 = strLabel7
    myForm.Label7.Font.Size = 13
    myForm.Label7.Font.Bold = True
        
    myForm.Label3 = strLabel3
    myForm.Label3.Font.Size = 11
    myForm.Label8 = strLabel8
    myForm.Label8.Font.Size = 11
    myForm.Label4 = strLabel4
    myForm.Label4.Font.Size = 13
    myForm.Label4.Font.Bold = True
    myForm.CommandButton1.Caption = button1
    myForm.CommandButton1.Font.Size = 11
    
    ' Help File
    myForm.Label5 = strLabel5
    myForm.Label5.Font.Size = 8
    myForm.Label5.Font.Italic = True
    
    Application.StatusBar = "Macro has been initiated."
    myForm.Show

End Sub
'---------------------------------------------------------------------------------------
' Date Created : September 7, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : September 7, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HELPFILE
' Description  : This function will feed the help tip section depending on the button
'                that has been activated.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function HELPFILE(ByVal Notification As Integer) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case 1
            NotifyUser = "TITLE: ACRU MENU PARAMETERIZATION MACRO" & vbLf
            NotifyUser = NotifyUser & "DESCRIPTION: This macro will save the ACRU OUT " & _
                "file as .XLSX file. " & vbLf
            NotifyUser = NotifyUser & "INPUT: Directory and ACRU OUT Files." & vbLf
            NotifyUser = NotifyUser & "OUTPUT: LOG_RUN file, ACRU OUT Files in .XLSX format" & vbLf
    End Select
    
    HELPFILE = NotifyUser
    
End Function
'---------------------------------------------------------------------
' Date Created : September 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : September 22, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveLogFile
' Description  : This function saves file as .TXT.
'                When new file is named after an existing file, the
'                same name is used with an number attached to it.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveLogFile(ByVal fileDir As String, _
ByVal fileName As String, ByVal fileExt As String) As String

    Dim saveFile As String
    Dim formatDate As String
    Dim saveDate As String
    Dim saveName As String
    Dim sPath As String

    ' Date
    formatDate = Format(Date, "MM/dd/yyyy")
    saveDate = Replace(formatDate, "/", "")
    
    ' Save information as Temp, which can then be renamed later..
    sPath = fileDir
    If Right(fileDir, 1) <> "\" Then sPath = fileDir & "\"
    saveName = fileName & "_" & saveDate & fileExt
    
    ' Rename existing file
    i = 1
    If CheckFileExists(sPath, saveName) = True Then
        If Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) <> "" Then
            Do Until Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) = ""
                i = i + 1
            Loop
            saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        Else: saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        End If
    Else: saveFile = sPath & fileName & "_" & saveDate & fileExt
    End If
    
    SaveLogFile = saveFile
    
End Function

