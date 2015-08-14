Attribute VB_Name = "NS_WARNING"
'---------------------------------------------------------------------------------------
' Date Created : July 31, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 10, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MacroTimer
' Description  : This function will notify user how much time has elapsed to complete
'                the entire procedure.
' Parameters   : Long
' Returns      : String
'---------------------------------------------------------------------------------------
Function MacroTimer(ByVal TimeElapsed As Long, ByVal MasterFile As String) As String

    Dim NotifyUser As String
    
    NotifyUser = "MACRO RUN IS SUCCESSFUL!" & vbCrLf
    NotifyUser = NotifyUser & vbCrLf
    NotifyUser = NotifyUser & "The macro has finished processing ACRU output files. "
    NotifyUser = NotifyUser & "Processing took a total of " & TimeElapsed & " seconds." & vbCrLf
    NotifyUser = NotifyUser & vbCrLf
    NotifyUser = NotifyUser & "Your OUTPUT file and directory can be found here: "
    NotifyUser = NotifyUser & MasterFile

    MacroTimer = NotifyUser
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 13, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : February 21, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MacroCancel
' Description  : This function will notify user how much time has elapsed to complete
'                the entire procedure.
' Parameters   : Long
' Returns      : String
'---------------------------------------------------------------------------------------
Function MacroCancel(ByVal notificationIndex As Integer) As String

    Dim Message As String, NotifyUser As String
    
    Select Case notificationIndex
        Case 1
            Message = "The selected ACRU output file is invalid. Please select another file."
        Case 2
            Message = "The workbook already exists. The tool will save to a new file with the same run number."
        Case 3
            Message = "The user specified variable name does not exist in one or more of the selected ACRU output files."
    End Select
    
    NotifyUser = "The macro run has been cancelled. "
    NotifyUser = NotifyUser & Message
    MacroCancel = NotifyUser
    
End Function
