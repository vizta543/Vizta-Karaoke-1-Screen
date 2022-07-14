VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "frmabout"
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsLogo 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      _cx             =   9922
      _cy             =   5424
      FlashVars       =   ""
      Movie           =   "c:\new.swf"
      Src             =   "c:\new.swf"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsTombol 
      Height          =   1020
      Left            =   8280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10320
      Width           =   960
      _cx             =   1693
      _cy             =   1799
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\Source\Source\potongan\Help\exit.swf"
      Src             =   "D:\Project\VOD\Source\Source\potongan\Help\exit.swf"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "frmabout"
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
    
Private Sub flsTombol_GotFocus()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    flsLogo.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    vpbfrmAbout = True
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    If frmUser.settingScreenResolution = "S-SD" Then
        flsLogo.Top = 0
        flsLogo.Left = 0
        flsLogo.Width = Screen.Width
        flsLogo.Height = Screen.Height
        flsTombol.ScaleMode = 1
        flsTombol.Left = 7000
        flsTombol.Top = 10600
        flsTombol.ZOrder 0
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        flsLogo.ScaleMode = 2
        flsLogo.Top = 0
        flsLogo.Left = 0
        flsLogo.Width = Screen.Width
        flsLogo.Height = Screen.Height
        flsTombol.Left = 9000
        flsTombol.Top = 9999
        flsTombol.ZOrder 0
    Else
        flsLogo.ScaleMode = 2
        flsLogo.Top = 0
        flsLogo.Left = 0
        flsLogo.Width = Screen.Width
        flsLogo.Height = Screen.Height
        flsTombol.Left = 13500
        flsTombol.Top = 15300
        flsTombol.ZOrder 0
    End If

    Dim lokasi As String
    lokasi = App.Path + "\Picture\about\"
    flsLogo.Movie = lokasi + "about.swf"
    flsTombol.Movie = lokasi + "exit.swf"
    
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 11
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    vpbfrmAbout = False
    frmRoom.vVideo = vvideotemp
End Sub

Sub prcOK()
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
