VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmCategory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmCategory"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "frmCategory.frx":0000
      Top             =   2040
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsMovieCategory 
      Height          =   5565
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7050
      _cx             =   12435
      _cy             =   9807
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\I-Sing\Source\potongan\Movie\Anim\call"
      Src             =   "D:\Project\VOD\I-Sing\Source\potongan\Movie\Anim\call"
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
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

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
    
Dim pvCategory As Integer
Dim lokasi As String
Dim vvideotemp As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        prcOK
        Exit Sub
    End If
    
    If KeyCode = vbKeyUp Then
        Select Case pvCategory
        Case 0
           flsMovieCategory.Movie = lokasi + "\picture\anim\ccartoon"
           pvCategory = 5
        Case 1
            flsMovieCategory.Movie = lokasi + "\picture\anim\call"
            pvCategory = 0
        Case 2
            flsMovieCategory.Movie = lokasi + "\picture\anim\cdrama"
            pvCategory = 1
        Case 3
            flsMovieCategory.Movie = lokasi + "\picture\anim\ccomedy"
            pvCategory = 2
        Case 4
            flsMovieCategory.Movie = lokasi + "\picture\anim\caction"
            pvCategory = 3
        Case 5
            flsMovieCategory.Movie = lokasi + "\picture\anim\choror"
            pvCategory = 4
        End Select
    End If
    If KeyCode = vbKeyDown Then
        Select Case pvCategory
        Case 0
           flsMovieCategory.Movie = lokasi + "\picture\anim\cdrama"
           pvCategory = 1
        Case 1
            flsMovieCategory.Movie = lokasi + "\picture\anim\ccomedy"
            pvCategory = 2
        Case 2
            flsMovieCategory.Movie = lokasi + "\picture\anim\caction"
            pvCategory = 3
        Case 3
            flsMovieCategory.Movie = lokasi + "\picture\anim\choror"
            pvCategory = 4
        Case 4
            flsMovieCategory.Movie = lokasi + "\picture\anim\ccartoon"
            pvCategory = 5
        Case 5
            flsMovieCategory.Movie = lokasi + "\picture\anim\call"
            pvCategory = 0
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    vpbFrmCategory = True
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 14

    lokasi = App.Path
    
    Skin1.LoadSkin lokasi + "\skin\sknMovieCategory.skn"
    Skin1.ApplySkinByName hWnd, "skncategory"
    
    flsMovieCategory.Movie = lokasi + "\picture\anim\call"
        
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    frmRoom.Enabled = False
'    frmRoom.hotnon
    
    pvCategory = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmRoom.vVideo = vvideotemp
    vpbFrmCategory = False
    'frmRoom.hot
End Sub

Public Sub prcOK()
    On Error Resume Next
    frmRoom.Enabled = True
    frmRoom.vMovieKategori = pvCategory
    frmRoom.moviecari
    
        Select Case pvCategory
        Case 0
           frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\all"
        Case 1
            frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\drama"
        Case 2
            frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\comedy"
        Case 3
            frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\action"
        Case 4
            frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\horor"
        Case 5
            frmRoom.flsMovieCategory.Movie = lokasi + "\picture\anim\cartoon"
        End Select
        
    Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        prcOK
        Exit Sub
    End If
End Sub


