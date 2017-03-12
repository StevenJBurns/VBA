VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusInputForm 
   Caption         =   "Status Monkey Input"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   OleObjectBlob   =   "StatusInputForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "StatusInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExitStatusMonkey_Click()
    ActivePresentation.Save
    ActivePresentation.Close
    Application.Quit
    Unload Me
End Sub

Private Sub StatusITS_Change()
    

    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    StatusITS.AddItem "Normal"
    StatusITS.AddItem "Caution"
    StatusITS.AddItem "Extreme"
    
End Sub
