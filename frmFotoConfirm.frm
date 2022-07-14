VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmFotoConfirm 
   BorderStyle     =   0  'None
   Caption         =   "frmFotoConfirm"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   2820
      TabIndex        =   0
      Top             =   1920
      Width           =   9855
      Begin VB.Image Image1 
         Height          =   6855
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmFotoConfirm.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmFotoConfirm"
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
    Skin1.LoadSkin lokasi + "\skin\foto.skn"
    Skin1.ApplySkinByName hWnd, "fotoForm"

    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
End Sub

Private Sub Skin1_Click(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next
    If Source.GetName = "ButtonDelete" Then
        deletefile
        Unload Me
    ElseIf Source.GetName = "ButtonSave" Then
        Unload Me
    End If
End Sub

Sub deletefile()
    On Error Resume Next
    Dim strPath As String
    strPath = frmRoom.picfile
    Kill strPath
End Sub
