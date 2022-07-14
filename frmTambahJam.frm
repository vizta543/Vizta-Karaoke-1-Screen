VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmTambahJam 
   BorderStyle     =   0  'None
   Caption         =   "frmTambahJam"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmTambahJam.frx":0000
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsjawab 
      Height          =   960
      Left            =   3720
      TabIndex        =   1
      Top             =   4740
      Width           =   3330
      _cx             =   5874
      _cy             =   1693
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\Vizta\VOD 4\potongan\main\yes.swf"
      Src             =   "D:\Project\VOD\Vizta\VOD 4\potongan\main\yes.swf"
      WMode           =   "Window"
      Play            =   0   'False
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmTambahJam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpbtjWaktuHabis As Date
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
    
Dim vok As Integer
Dim lokasi As String
Dim vvideotemp As Integer

Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    vpbfrmTambahJam = True
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    lokasi = App.Path
    
        Skin1.LoadSkin lokasi + "\skin\skntambahjam.skn"
        Skin1.ApplySkinByName hWnd, "skntambahjam"

    flsjawab.Movie = lokasi + "\picture\anim\yes"
    vok = 1
    
    frmRoom.Enabled = False
    
    LockRoom
    
'    frmRoom.hotnon
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 16
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    vpbfrmTambahJam = False
    frmRoom.vVideo = vvideotemp
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        prcOK
    End If
    If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) _
    Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
        If vok = 1 Then
            vok = 2
            flsjawab.Movie = lokasi + "\picture\anim\no"
        ElseIf vok = 2 Then
            vok = 1
            flsjawab.Movie = lokasi + "\picture\anim\yes"
        Else
            vok = 3
        End If
    End If
End Sub

Public Sub prcOK()
    On Error Resume Next
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim jam As Integer
    Dim waktu As String
    Dim pesan As String
    
    frmRoom.vTambahWaktu = False
    If vok = 1 Then
        frmRoom.vTambahWaktu = True
        
        'MASUKKAN MESSAGE MSSQL
        waktu = Str$(year(Now)) & "-" & Trim$(Str$(month(Now))) & "-" & Trim$(Str$(day(Now))) & " " & hour(Now) & ":" & minute(Now) & ":" & second(Now)
        pesan = "Tambah Jam"
        
        KoneksiAdoDBVizta.Execute "INSERT INTO Tmessage " & _
                "            (tanggal, pesan, dari, status) " & _
                "VALUES      ('" & waktu & "','" & pesan & "','" & frmRoom.txtCompName & "',0)"
    End If
    
    Unload Me

    If vpbFrmCountry = True Then
        frmCountry.Show
    ElseIf vpbFrmCategory = True Then
        frmCategory.Show
    ElseIf vpbFrmConfirmasi = True Then
        frmConfirmasi.Show
    ElseIf vpbFrmCall = True Then
        frmCall.vpengirim = 3
        frmCall.Show
        frmCall.Timer1.Enabled = True
    ElseIf vpbfrmSaran = True Then
        frmSaran.Show
    Else
        frmRoom.Enabled = True
        Select Case frmRoom.vpointer
        Case 1
            frmRoom.txtSearch.SetFocus
        Case 2
            frmRoom.txtSearch.SetFocus
        Case 4
            frmRoom.txtChat.SetFocus
        Case 5
            frmRoom.lstTV.SetFocus
        End Select
        LockRoom
    End If
End Sub

