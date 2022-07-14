VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmScore 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "frmScore"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScore 
      Interval        =   6000
      Left            =   3000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmScore.frx":0000
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsScore 
      Height          =   4200
      Left            =   1040
      TabIndex        =   2
      Top             =   2680
      Width           =   3315
      _cx             =   5856
      _cy             =   7408
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\I-Sing\Source\potongan\Score\100.swf"
      Src             =   "D:\Project\VOD\I-Sing\Source\potongan\Score\100.swf"
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
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      BackColor       =   &H005F1426&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   80.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1860
      Left            =   4920
      TabIndex        =   0
      Text            =   "99"
      Top             =   2400
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpScore 
      Height          =   585
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   1320
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   51
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2328
      _cy             =   1032
   End
End
Attribute VB_Name = "frmScore"
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
    
Private Sub Form_Activate()
    On Error Resume Next
        If vpbFrmCountry = True Then
            frmCountry.lstCountry.SetFocus
        ElseIf vpbFrmCategory = True Then
            frmCategory.Text1.SetFocus
        ElseIf vpbFrmConfirmasi = True Then
            frmConfirmasi.Text2.SetFocus
        ElseIf vpbFrmCall = True Then
            frmCall.txtPesan.SetFocus
        ElseIf vpbfrmTambahJam = True Then
            frmTambahJam.Text1.SetFocus
        ElseIf vpbfrmSaran = True Then
            frmSaran.txtPesan.SetFocus
        ElseIf vpbfrmNew = True Then
            frmNew.lstTop.SetFocus
        ElseIf vpbfrmPopuler = True Then
            frmPopuler.lstTop.SetFocus
        ElseIf vpbfrmAbout = True Then
            frmabout.flsLogo.SetFocus
        ElseIf vpbfrmHelp = True Then
            frmHelp.flsLogo.SetFocus
        ElseIf vpbfrmVocal = True Then
            frmVocal.Text1.SetFocus
        Else
            If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 5) Then
                If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                ElseIf frmRoom.vpointer = 3 Then
                    frmRoom.lstPlaylist.SetFocus
                End If
            ElseIf vVideo = 3 Then
                sMakeCaret frmRoom.txtChat, frmRoom.caretLebar, frmRoom.caretTinggi
            ElseIf vVideo = 7 Then  'TV
                lstTV.SetFocus
            End If
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim nilai As Integer
    Dim lokasi As String
    lokasi = App.Path
    
    Skin1.LoadSkin lokasi + "\skin\sknscore.skn"
    Skin1.ApplySkinByName hWnd, "sknscore"
        
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    
    wmpScore.Width = 0
    wmpScore.Left = 0
    Randomize
    nilai = Int((30 * Rnd) + 1) + 69
    txtScore.text = Str$(nilai)
    
    wmpScore.settings.volume = 30
    If nilai > 90 Then
        wmpScore.URL = lokasi + "\picture\music\music5.wav"
        flsScore.Movie = lokasi + "\picture\music\score 100.swf"
    ElseIf nilai > 80 Then
         wmpScore.URL = lokasi + "\picture\music\music4.wav"
         flsScore.Movie = lokasi + "\picture\music\score 90.swf"
    ElseIf nilai > 70 Then
         wmpScore.URL = lokasi + "\picture\music\music3.wav"
         flsScore.Movie = lokasi + "\picture\music\score 80.swf"
    ElseIf nilai > 60 Then
         wmpScore.URL = lokasi + "\picture\music\music2.wav"
         flsScore.Movie = lokasi + "\picture\music\score 70.swf"
    ElseIf nilai >= 50 Then
         wmpScore.URL = lokasi + "\picture\music\music1.wav"
         flsScore.Movie = lokasi + "\picture\music\score 60.swf"
    Else
         flsScore.Movie = lokasi + "\picture\music\score 80.swf"
    End If
    wmpScore.Controls.play
    LockRoom
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    LockRoom
End Sub


Private Sub tmrScore_Timer()
    On Error Resume Next
    Unload Me
End Sub

