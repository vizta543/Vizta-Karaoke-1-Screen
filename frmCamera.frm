VERSION 5.00
Object = "{F15158C8-31F4-4D02-A18E-FFDF0FFFE433}#1.0#0"; "videocap.ocx"
Begin VB.Form frmCamera 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmCamera"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   1095
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VIDEOCAPLib.VideoCap VideoCap1 
      Height          =   11520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _Version        =   65536
      AspectRatio     =   0   'False
      TVMute          =   -1  'True
      TextFontName    =   ""
      LicenseKey      =   "videocapproB02032004"
      _ExtentX        =   27093
      _ExtentY        =   20320
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCamera"
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
            If frmRoom.Enabled Then
                If (frmRoom.vVideo = 0) Then
                    If frmRoom.lstPlaylist.Visible Then
                        frmRoom.txtUser.SetFocus
                    End If
                    If frmRoom.lstAll.Visible Then
                        frmRoom.lstAll.SetFocus
                    End If
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
        End If
        
End Sub

Private Sub Form_Load()
    On Error Resume Next
    frmCamera.VideoCap1.LicenseKey = "videocapproB02032004"

'    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
'    SWP_NOMOVE + SWP_NOSIZE
'    Me.Move (Screen.Width - Me.Width), 0
'    Me.Top = 1050
    
    vpbfrmCamera = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    VideoCap1.stop
    vpbfrmCamera = False
End Sub

