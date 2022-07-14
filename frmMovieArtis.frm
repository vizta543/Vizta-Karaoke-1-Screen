VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmMovieArtis 
   Caption         =   "frmMovieArtis"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPesan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006D6D6D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMovieArtis.frx":0000
      Top             =   2040
      Width           =   7455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmMovieArtis.frx":0006
      Top             =   0
   End
End
Attribute VB_Name = "frmMovieArtis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" _
   (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
    On Error Resume Next
    lokasi = App.Path

    Skin1.LoadSkin lokasi + "\skin\artis.skn"
    Skin1.ApplySkinByName hWnd, "artis"
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    
    frmRoom.Enabled = False
    LockRoom
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmRoom.Enabled = True
    LockRoom
End Sub

Private Sub Skin1_Click(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next
If Source.GetName = "btnclose" Then
    Unload Me
End If
End Sub

