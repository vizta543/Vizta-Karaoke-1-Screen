VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmVideo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmVideo"
   ClientHeight    =   11520
   ClientLeft      =   2265
   ClientTop       =   1920
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillStyle       =   5  'Downward Diagonal
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   11520
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
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
      volume          =   50
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
      _cx             =   27093
      _cy             =   20320
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pvPlaystate As Long
Public pbDurasi As Double
Public pbDurasiAkhir As Double
Public pbTempo As Double

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
            If (frmRoom.vVideo = 0) Then
                If vpbfrmCamera Then
                    frmCamera.Show
                End If
                frmTransparent.Show
                If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                ElseIf frmRoom.vpointer = 3 Then
                    frmRoom.lstPlaylist.SetFocus
                End If
            ElseIf frmRoom.vVideo = 5 Then
                If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                ElseIf frmRoom.vpointer = 3 Then
                    frmRoom.lstPlaylist.SetFocus
                End If
            ElseIf frmRoom.vVideo = 3 Then
                sMakeCaret frmRoom.txtChat, frmRoom.caretLebar, frmRoom.caretTinggi
            ElseIf frmRoom.vVideo = 7 Then  'TV
                 frmRoom.lstTV.SetFocus
            End If
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    LockRoom
    frmRoom.vVideoAktif = True
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.WindowState = 2
    
    WindowsMediaPlayer1.Top = 0
    WindowsMediaPlayer1.Left = 0
    WindowsMediaPlayer1.stretchToFit = True
    'edited by Andi 22-01-2021
    WindowsMediaPlayer1.Width = Screen.Width
    WindowsMediaPlayer1.Height = Screen.Height
    'edited by Andi 22-01-2021
    WindowsMediaPlayer1.network.bufferingTime = 16000
    
    Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    LockRoom
    frmRoom.vVideoAktif = False
    vpbBolehMainVocal = True
End Sub

Public Sub Timer1_Timer()

    On Error Resume Next
    
    If pvPlaystate = 3 Then
        pbDurasiAkhir = pbDurasiAkhir + (1 + (pbTempo * 1))
    End If
If Err.Number <> 0 Then
  LogError Name, "Timer1_Timer 1"
End If
    
    If (pvPlaystate <> 1) And (ScoreValid = True) And (ScoreSetup = True) Then
        If (WindowsMediaPlayer1.currentMedia.duration <> 0) And (pbDurasiAkhir >= WindowsMediaPlayer1.currentMedia.duration - 7) Then
            frmScore.Show
            ScoreValid = False
        End If
    End If
If Err.Number <> 0 Then
  LogError Name, "Timer1_Timer 2"
End If
    
    If (WindowsMediaPlayer1.currentMedia.duration <> 0) And (pbDurasiAkhir >= WindowsMediaPlayer1.currentMedia.duration - 2) Then
        vpbBolehMainVocal = False
    Else
        vpbBolehMainVocal = True
    End If
If Err.Number <> 0 Then
  LogError Name, "Timer1_Timer 3"
End If
    
    If (pvPlaystate = 1) Or (pbDurasiAkhir >= WindowsMediaPlayer1.currentMedia.duration) Then
        Timer1.Enabled = False
        frmRoom.cmdStop_Click
    End If
    
    
    If Err.Number <> 0 Then
      LogError Name, "Timer1_Timer"
    End If
End Sub


Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal newState As Long)
    
    On Error Resume Next
    
    If newState < 32000 Then
        pvPlaystate = newState
    End If
    
    'PLAY
    If newState = 3 Then
        Unload frmPromo
        Timer1.Enabled = True
    End If
    
    'STOP
    If newState = 1 Then
    '   If frmRoom.vrekamstate Then
    '        frmRoom.btnRecStop_Click
    '   End If
    End If
    
    'CONNECTING
    If newState = 10 Then 'stop
       pbDurasi = WindowsMediaPlayer1.currentMedia.duration
       pbDurasiAkhir = 0
       pbTempo = 0
       SetPitch 0
       SetTempo 1
       frmRoom.vTempo = 0
    End If
    
    LockRoom
    
    If Err.Number <> 0 Then
      LogError Name, "WindowsMediaPlayer1_PlayStateChange"
    End If
End Sub
