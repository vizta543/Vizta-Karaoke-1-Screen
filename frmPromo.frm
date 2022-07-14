VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmPromo 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "frmPromo"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsPromo 
      Height          =   11520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _cx             =   27093
      _cy             =   20320
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\I-Sing\Source\potongan\main\logo.swf"
      Src             =   "D:\Project\VOD\I-Sing\Source\potongan\main\logo.swf"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ExactFit"
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
Attribute VB_Name = "frmPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub flsPromo_FSCommand(ByVal command As String, ByVal args As String)
  
  On Error Resume Next
  
  If command = "stop" Then
    loadPromo
  End If
  
  If Err.Number <> 0 Then
    LogError Me.Name, "flsPromo_FSCommand"
  End If
End Sub

Private Sub Form_Load()
    
  On Error Resume Next
  
      'added by Andi 22-01-2021
    flsPromo.ScaleMode = 2
    flsPromo.Width = Screen.Width
    flsPromo.Height = Screen.Height
    flsPromo.Visible = True
    'added by Andi 22-01-2021
  
  loadPromo
  
  Form2.Visible = False
  
  If Err.Number <> 0 Then
    LogError Me.Name, "Form_Load"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Form2.Visible = True
End Sub

Sub loadPromo()

  On Error Resume Next
  
  Dim promoAnimationPath As String
  
  promoAnimationPath = frmRoom.promoAnimationCollection(frmRoom.promoAnimationCollectionCurrent)
  
  If frmRoom.FileExists("\\" & vpbServerUtama & promoAnimationPath) Then
    promoAnimationPath = "\\" & vpbServerUtama & promoAnimationPath
  ElseIf frmRoom.FileExists("\\" & vpbServerBackup & promoAnimationPath) Then
    promoAnimationPath = "\\" & vpbServerBackup & promoAnimationPath
  Else
    promoAnimationPath = ""
  End If

  If promoAnimationPath <> "" Then
    flsPromo.Movie = promoAnimationPath
  End If
  
  frmRoom.promoAnimationCollectionCurrent = frmRoom.promoAnimationCollectionCurrent + 1
  If frmRoom.promoAnimationCollectionCurrent > frmRoom.promoAnimationCollection.Count Then
    frmRoom.promoAnimationCollectionCurrent = 1
  End If
  
  
  If Err.Number <> 0 Then
    LogError Name, "loadPromo"
  End If
End Sub

