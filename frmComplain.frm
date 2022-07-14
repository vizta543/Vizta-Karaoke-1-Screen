VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmSaran 
   BorderStyle     =   0  'None
   Caption         =   "frmSaran"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPesan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2655
      Left            =   1320
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3120
      Width           =   4935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmComplain.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmSaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
Dim vvideotemp As Integer
    
Private Sub cmdpesanCancel_Click()
    On Error Resume Next
    Unload Me
    Select Case frmRoom.vpointer
    Case 1
        frmRoom.txtSearch.SetFocus
    Case 2
        frmRoom.txtSearch.SetFocus
    Case 3
        frmRoom.lstPlaylist.SetFocus
    Case 4
        frmRoom.txtChat.SetFocus
    Case 5
        frmRoom.lstTV.SetFocus
    End Select
End Sub

Sub prcOK()
    On Error Resume Next
    If txtPesan.text <> "" Then
        '---- Input ----
        Dim Sql As String
        Dim myrs As MYSQL_RS
        Sql = "INSERT INTO complain (room, pesan, nama, tanggal) VALUES ('" & mysql_escape_string(frmRoom.txtCompName.text) & "' , '" & mysql_escape_string(txtPesan.text) & "', '" & frmRoom.txtUser.text & "', '" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "-" & " " & Time$ & "')"
        Set myrs = MyConn.Execute(Sql)
    End If
    
    Unload Me
    
    Select Case frmRoom.vpointer
    Case 1
        frmRoom.txtSearch.SetFocus
    Case 2
        frmRoom.txtSearch.SetFocus
    Case 3
        frmRoom.lstPlaylist.SetFocus
    Case 4
        frmRoom.txtChat.SetFocus
    Case 5
        frmRoom.lstTV.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim lokasi As String
    lokasi = App.Path

        Skin1.LoadSkin lokasi + "\skin\sknsaran.skn"
        Skin1.ApplySkinByName hWnd, "sknsaran"
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    frmRoom.Enabled = False
    
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 10
    vpbfrmSaran = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmRoom.Enabled = True
    frmRoom.vVideo = vvideotemp
    vpbfrmSaran = False
    LockRoom
End Sub

Private Sub Skin1_Click(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next
If Source.GetName = "btnok" Then
    prcOK
ElseIf Source.GetName = "btncancel" Then
    cmdpesanCancel_Click
End If
End Sub

