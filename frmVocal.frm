VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmVocal 
   Caption         =   "frmVocal"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmVocal.frx":0000
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmVocal"
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
    
Public VocalAktif As Integer

Private Sub Form_Load()
    On Error Resume Next
    vpbfrmVocal = True
    
    lokasi = App.Path
    Text1.Width = 0
    Text1.Height = 0
    Select Case VocalAktif
    Case 0
        Skin1.LoadSkin lokasi + "\skin\sknvocal.skn"
        Skin1.ApplySkinByName hWnd, "vocalon"
    Case 1
        Skin1.LoadSkin lokasi + "\skin\sknvocal.skn"
        Skin1.ApplySkinByName hWnd, "vocaloff"
    Case 2
        Skin1.LoadSkin lokasi + "\skin\sknvocal.skn"
        Skin1.ApplySkinByName hWnd, "scoreon"
    Case 3
        Skin1.LoadSkin lokasi + "\skin\sknvocal.skn"
        Skin1.ApplySkinByName hWnd, "scoreoff"
    Case Else
        Unload Me
        Exit Sub
    End Select
       
    frmRoom.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    vpbfrmVocal = False
    frmRoom.Enabled = True
    
    If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 2) Then
              If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                  sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
              ElseIf (frmRoom.vpointer = 3) And (frmRoom.lstPlaylist.Visible = True) Then
                  If (frmRoom.lstPlaylist.ListItems.Count > 0) Then
                      frmRoom.txtUser.SetFocus
                      frmRoom.lstPlaylist.SetFocus
                  Else
                      frmRoom.lstAll.Visible = True
                      frmRoom.lstAll.SetFocus
                      sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                  End If
             End If
    End If
End Sub

