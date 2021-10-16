VERSION 5.00
Begin VB.Form ShapedForm 
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3180
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ShapedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The following code is only used to allow form draging
'from any part of it

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
