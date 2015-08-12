VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ACRU Input File Path and Output File Name"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' Date Created : February 21, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : February 21, 2014
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Commandbutton1_Click
' Description  : This function will pass on the values entered by the
'                user in order to create an output file in a specific
'                naming format: NS_HRU####_MMDDYYYY.xlsx
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Private Sub Commandbutton1_Click()

    outRUNVAL = UserForm1.FormRun.Value
    Call NASHSUTCLIFF_MAIN(outRUNVAL)

End Sub
