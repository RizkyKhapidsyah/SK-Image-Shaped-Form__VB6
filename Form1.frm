VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Shaped Form Example"
   ClientHeight    =   3210
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   4710
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mChangeBackgroundPicture 
         Caption         =   "Change Background &Picture"
      End
      Begin VB.Menu mInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables must be declared
Option Explicit

'This is the variable that will keep the memory
'address of the region
Private hRgn As Long

'Constants declaration needed for the CommonDialog
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100

Private Sub Form_Load()
    Me.Move 0, 0, 4000, 690
'Set the transparent color to White, Create the region
'and modify the Forms Shape with it
    CommonDialog1.Color = vbWhite
    SetRegion
'Show the Shaped Form
    ShapedForm.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Free the used memory by the Region and unload the Shaped
'Form
    If hRgn Then DeleteObject hRgn
    Unload ShapedForm
End Sub

Private Sub mChangeBackgroundPicture_Click()
    On Error Resume Next
    Err.Clear
    With CommonDialog1
'Set the CommonDialog Open File options
        .DialogTitle = "Please Select a Picture"
        .Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_LONGNAMES + OFN_NONETWORKBUTTON + OFN_PATHMUSTEXIST
        .Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur|Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|Icons (*.ico;*.cur)|*.ico;*.cur|All Files (*.*)|*.*"
        .ShowOpen
'Check if Cancel was pressed
        If Err.Number = 32755 Then Exit Sub
'Set the CommonDialog Color Select options
        .Flags = CC_FULLOPEN + CC_SOLIDCOLOR + CC_RGBINIT + CC_ANYCOLOR
        .ShowColor
'Check if Cancel was pressed
        If Err.Number = 32755 Then Exit Sub
        On Error GoTo erro
'Make the Shaped Form invisible
        ShapedForm.Visible = False
        DoEvents
'Change the Forms Background Picture, Width and Height
'It's necessary that the forms dimensions are equal or
'bigger that the Picture ones.
        ShapedForm.Picture = LoadPicture(.FileName)
        ShapedForm.Width = ShapedForm.Picture.Width
        ShapedForm.Height = ShapedForm.Picture.Height
'Set it's new Shape based on it's Background Picture and
'Transparent Color
        SetRegion
    End With
erro:
'Error handler
    If Err.Number <> 0 Then MsgBox "Error Number " & Err.Number & " : " & Err.Description, vbApplicationModal + vbCritical
'Make the Shaped Form visible
    ShapedForm.Visible = True
End Sub

Private Sub mExit_Click()
'Unload the Menu Form
    Unload Me
End Sub

Private Sub SetRegion()
'Free the memory allocated by the previous Region
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region
    hRgn = GetBitmapRegion(ShapedForm.Picture, CommonDialog1.Color)
'Set the Forms new Region
    SetWindowRgn ShapedForm.hwnd, hRgn, True
End Sub

Private Sub mInstructions_Click()
'Show a message box with a simple explanation
    Dim Texto As String
    Texto = "This is what really happens:" & vbCrLf & vbCrLf
    Texto = Texto & "The Background Picture of the Form and a particular colour is passed to a function. Then, the Image is scanned and all pixels that have equal colour to the Transparent Colour are removed from the Image, creating a new virtual Image (a Region, to be exact) that will be used by the form. The smaller the picture is, the faster it is scanned." & vbCrLf & vbCrLf & vbCrLf
    Texto = Texto & "Programmed by Pedro Lamas" & vbCrLf & "Copyright ©1997-1999 Underground Software" & vbCrLf & vbCrLf
    Texto = Texto & "Home-Page (Dedicated to VB): www.terravista.pt/portosanto/3723/" & vbCrLf & "E-Mail: sniper@hotpop.com"
    MsgBox Texto, vbApplicationModal + vbInformation, "Instructions"
End Sub
