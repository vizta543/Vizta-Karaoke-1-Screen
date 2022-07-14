VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmCall 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmCall"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   Picture         =   "frmCall.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1680
      Top             =   1080
   End
   Begin VB.TextBox txtPesan 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      IMEMode         =   3  'DISABLE
      Left            =   7560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   6120
      Width           =   3855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmCall.frx":27AFC
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsAnim 
      Height          =   4200
      Left            =   1560
      TabIndex        =   1
      Top             =   1650
      Visible         =   0   'False
      Width           =   3315
      _cx             =   5856
      _cy             =   7408
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\Vizta\VOD7\OneScreen\One Screen Remote\potongan\Source\effect\effect all.swf"
      Src             =   "D:\Project\VOD\Vizta\VOD7\OneScreen\One Screen Remote\potongan\Source\effect\effect all.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
End
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Penampakan
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
    
Public vpengirim As Integer
'0=addmemberplaylist
Dim gagal As Boolean
Dim vvideotemp As Integer

Private Sub Form_Activate()
    On Error Resume Next
    txtPesan.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    vpbFrmCall = True
    gagal = False
    Dim lokasi As String
    lokasi = App.Path
    If vpengirim = 0 Then 'member
        Skin1.LoadSkin lokasi + "\skin\sknaddmember.skn"
        Skin1.ApplySkinByName hWnd, "sknadd"
    ElseIf vpengirim = 1 Then 'delete member
        Skin1.LoadSkin lokasi + "\skin\sknaddmember.skn"
        Skin1.ApplySkinByName hWnd, "skndelete"
    ElseIf vpengirim = 2 Then 'add member
        Skin1.LoadSkin lokasi + "\skin\sknaddmember.skn"
        Skin1.ApplySkinByName hWnd, "sknclear"
    ElseIf vpengirim = 3 Then 'CALL
        Skin1.LoadSkin lokasi + "\skin\skncall.skn"
        Skin1.ApplySkinByName hWnd, "skncall"
    ElseIf vpengirim = 4 Then 'effect mic
        Skin1.LoadSkin lokasi + "\skin\skneffect.skn"
        Skin1.ApplySkinByName hWnd, "skneffect"
        flsAnim.Movie = lokasi + "\picture\anim\effectampli"
        flsAnim.Visible = True
        flsAnim.SetVariable "vtulisan", vEffectAmpli
    End If

    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    frmRoom.Enabled = False
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmRoom.vVideo = vvideotemp
    vpbFrmCall = False
End Sub

Public Sub prcOK()
    Dim pesan As String
    Dim waktu As String
    
    On Error Resume Next
    
    Timer1.Enabled = False
    
    If vpengirim = 0 Or vpengirim = 1 Or vpengirim = 2 Or vpengirim = 4 Then
        gagal = True 'untuk melewati panggilan
    End If
    
    If gagal = False Then
        If txtPesan.text = "" Then
            pesan = "Panggilan"
        Else
            pesan = txtPesan.text
            pesan = Replace$(pesan, "'", "''")
        End If
    
        'MASUKKAN MESSAGE MSSQL
        Dim rs As New ADODB.Recordset
        waktu = Str$(year(Now)) & "-" & Trim$(Str$(month(Now))) & "-" & Trim$(Str$(day(Now))) & " " & hour(Now) & ":" & minute(Now) & ":" & second(Now)
    
        KoneksiAdoDBVizta.Execute "INSERT INTO Tmessage " & _
                "            (tanggal, pesan, dari, status) " & _
                "VALUES      ('" & waktu & "','" & pesan & "','" & frmRoom.txtCompName & "',0)"
    End If
    
    frmRoom.Enabled = True
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

Private Sub Timer1_Timer()
    On Error Resume Next
    prcOK
End Sub

Private Sub txtPesan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        prcOK
    End If
End Sub


