VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} window 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "control.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim running As Boolean

Private Sub UserForm_Activate()
    window.Left = -20
    window.Top = -35
    window.Width = 1200
    window.Height = 720
    
    running = True: While (running)
        DoEvents
        window.BackColor = RGB(Int((25 * Rnd) + 1), Int((25 * Rnd) + 1), Int((25 * Rnd) + 1))
    Wend
End Sub

Private Sub UserForm_Terminate()
    running = False
End Sub
