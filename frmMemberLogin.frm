VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmConfirmasi 
   BorderStyle     =   0  'None
   Caption         =   "frmConfirmasi"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAktif 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsAnim 
      Height          =   4200
      Left            =   1560
      TabIndex        =   4
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
   Begin VB.TextBox text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   3000
      TabIndex        =   1
      Text            =   "Load Member Playlist ?"
      Top             =   3860
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   0
      Top             =   3420
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmMemberLogin.frx":0000
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsjawab 
      Height          =   975
      Left            =   3660
      TabIndex        =   3
      Top             =   4760
      Width           =   3435
      _cx             =   6059
      _cy             =   1729
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\I-Sing\Source\potongan\Konfirmasi\yes"
      Src             =   "D:\Project\VOD\I-Sing\Source\potongan\Konfirmasi\yes"
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
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3600
      TabIndex        =   2
      Text            =   "Welcome"
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "frmConfirmasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpengirim As Integer '6=sedang recording 7=vcd 8=cd 9=diajak chat 10=HP 11=dvd
'Form penampakan
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
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
Dim lokasi As String
Dim vok As Integer
Dim vvideotemp As Integer
Dim vfokus As Integer
Dim ScreenSaverAktifConfirmasi As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ScreenSaverAktifConfirmasi = 0
    If KeyCode = vbKeyReturn Then
        frmRoom.Enabled = True
        If vok = 1 Then
            prcYes
        ElseIf vok = 2 Then
            prcNo
        ElseIf vok = 3 Then
            prcOK
        End If
        Exit Sub
    End If
    
    If vpengirim = 12 Then 'effect ampli
        If KeyCode = vbKeyDown Then
            Select Case vEffectAmpli
                Case 1
                    vEffectAmpli = 2
                Case 2
                    vEffectAmpli = 3
                Case 3
                    vEffectAmpli = 4
            End Select
        End If
        If KeyCode = vbKeyUp Then
            Select Case vEffectAmpli
                Case 2
                    vEffectAmpli = 1
                Case 3
                    vEffectAmpli = 2
                Case 4
                    vEffectAmpli = 3
            End Select
        End If
        flsAnim.SetVariable "vtulisan", vEffectAmpli
    
    ElseIf vpengirim = 14 Then
      If KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Then
          Select Case vEffectAmpli
            Case 1
              vEffectAmpli = 5
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
            Case 2
              vEffectAmpli = 6
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
            Case 3
              vEffectAmpli = 7
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
            Case 4
              vEffectAmpli = 8
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
            Case 5
            Case 6
            Case 7
            Case 8
            Case 9
              vEffectAmpli = 1
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 10
              vEffectAmpli = 2
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 11
              vEffectAmpli = 3
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 12
              vEffectAmpli = 4
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          End Select
      ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyUp Then
          Select Case vEffectAmpli
            Case 1
              vEffectAmpli = 9
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
            Case 2
              vEffectAmpli = 10
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
            Case 3
              vEffectAmpli = 11
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
            Case 4
              vEffectAmpli = 12
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
            Case 5
              vEffectAmpli = 1
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 6
              vEffectAmpli = 2
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 7
              vEffectAmpli = 3
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 8
              vEffectAmpli = 4
              frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
            Case 9
            Case 10
            Case 11
            Case 12
          End Select
      End If
    End If
    
    If (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
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

Private Sub Form_Load()
    On Error Resume Next
'----0=load member playlist
'----1=clear member playlist
'----5=order
'----6=vcd/cd sedang merekam
'----9=chat
'---10=hp

    Dim bytOpacity As Byte

    vpbFrmConfirmasi = True
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 15
    
    ScreenSaverAktifConfirmasi = 0
    
    Text1.Locked = True
    Text2.Locked = True
    text3.Locked = True
    
    lokasi = App.Path
    
    If (vpengirim = 0) Then
        Text2.PasswordChar = "*"
        Text1.Locked = False
        Text2.Locked = False
        
        Skin1.LoadSkin lokasi + "\skin\sknloadmember.skn"
        Skin1.ApplySkinByName hWnd, "sknloadmember"
        flsjawab.Movie = lokasi + "\picture\anim\yes"
        vok = 1
        tmrAktif.Enabled = True
    ElseIf (vpengirim = 5) Or (vpengirim = 7) Or (vpengirim = 8) Or (vpengirim = 9) Or (vpengirim = 10) Or (vpengirim = 11) Then
        Skin1.LoadSkin lokasi + "\skin\sknkonfirmasi.skn"
        Skin1.ApplySkinByName hWnd, "sknkonfirmasi"
        flsjawab.Movie = lokasi + "\picture\anim\yes"
        vok = 1
        tmrAktif.Enabled = True
    ElseIf vpengirim = 12 Then
        Skin1.LoadSkin lokasi + "\skin\skneffect.skn"
        Skin1.ApplySkinByName hWnd, "skneffect"
        flsAnim.Movie = lokasi + "\picture\anim\effectampli"
        If vEffectAmpli = 0 Then vEffectAmpli = 2
        Select Case vEffectAmpli
          Case 1
          Case 2
          Case 3
          Case 4
          Case 5
            vEffectAmpli = 1
          Case 6
            vEffectAmpli = 2
          Case 7
            vEffectAmpli = 3
          Case 8
            vEffectAmpli = 4
          Case 9
            vEffectAmpli = 1
          Case 10
            vEffectAmpli = 2
          Case 11
            vEffectAmpli = 3
          Case 12
            vEffectAmpli = 4
        End Select
        flsAnim.SetVariable "vtulisan", vEffectAmpli
        flsAnim.Visible = True
        
        flsjawab.Visible = False
        Text1.Width = 0
        Text2.Width = 0
        text3.Width = 0
        
        vok = 3
        tmrAktif.Enabled = True
        
    ElseIf vpengirim = 14 Then
        
        Width = 9450
        Height = 3660
        
        Skin1.LoadSkin lokasi + "\skin\micvol.skn"
        Skin1.ApplySkinByName hWnd, "micvol"
        
        flsAnim.Movie = lokasi + "\picture\anim\micvol"
        flsAnim.ZOrder vbBringToFront
        flsAnim.Visible = True
        flsAnim.Top = 0
        flsAnim.Left = 885
        flsAnim.Width = 5790
        flsAnim.Height = 3660
        
        If vEffectAmpli = 0 Then
          vEffectAmpli = 2
        End If
        If vEffectAmpli = 5 Or vEffectAmpli = 6 Or vEffectAmpli = 7 Or vEffectAmpli = 8 Then
          flsAnim.SetVariable "vtulisan", 1
        ElseIf vEffectAmpli = 1 Or vEffectAmpli = 2 Or vEffectAmpli = 3 Or vEffectAmpli = 4 Then
          flsAnim.SetVariable "vtulisan", 2
        ElseIf vEffectAmpli = 9 Or vEffectAmpli = 10 Or vEffectAmpli = 11 Or vEffectAmpli = 12 Then
          flsAnim.SetVariable "vtulisan", 3
        End If
        
        flsjawab.Visible = False
        
        Text1.Width = 0
        Text2.Width = 0
        text3.Width = 0
        
        vok = 3
        
        tmrAktif.Enabled = True

    ElseIf vpengirim = 15 Then
        
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        
        bytOpacity = 205
        
        Call SetLayeredWindowAttributes(hWnd, 0, bytOpacity, LWA_ALPHA)

        Width = 7305
        Height = 5400

        flsAnim.Movie = lokasi + "\picture\anim\dadu"
        flsAnim.ZOrder vbBringToFront
        flsAnim.Visible = True
        flsAnim.Top = 0
        flsAnim.Left = 0
        flsAnim.Width = 7305
        flsAnim.Height = 5400

        vok = 3
        
        Text1.Width = 0
        Text2.Width = 0
        text3.Width = 0
        
    Else
        Skin1.LoadSkin lokasi + "\skin\sknok.skn"
        Skin1.ApplySkinByName hWnd, "sknok"
        flsjawab.Movie = lokasi + "\picture\anim\ok"
        vok = 3
    End If
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    frmRoom.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmRoom.vVideo = vvideotemp

    vpbFrmConfirmasi = False
    frmRoom.Enabled = True
    If vpengirim = 9 Then
       Exit Sub
    End If
    
    If frmRoom.vVideo = 3 Then
      frmRoom.txtChat.SetFocus
      Exit Sub
    End If
    
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

Private Sub Skin1_Click(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next
    
    frmRoom.Enabled = True

    If Source.GetName = "btnyes" Then
        prcYes
    ElseIf Source.GetName = "btnno" Then
        prcNo
    ElseIf Source.GetName = "btnok" Then
        prcOK
    End If
End Sub

Private Sub chatConfirm(Data As String)
    On Error Resume Next
    If frmRoom.chat(1).State = sckConnected Then
        frmRoom.chat(1).SendData Data
    End If
End Sub

Private Sub prcYes()
    On Error Resume Next

    If vpengirim = 0 Then ' member
        prcLoginMember
    ElseIf vpengirim = 5 Then 'call
        frmRoom.Enabled = True
        If frmRoom.Socket(frmRoom.iSockets).State = sckConnected Then
            Dim Data As String
            Data = frmUser.txtRoom.text
            frmRoom.Socket(frmRoom.iSockets).SendData Data
            Unload Me
        Else
            Text1.text = "  Ada Kesalahan Teknis"
            Text2.text = "Gunakan Panggilan Manual"
            text3.text = "   Di Meja Anda "
            
            lokasi = App.Path
            Skin1.LoadSkin lokasi + "\skin\sknok.skn"
            Skin1.ApplySkinByName hWnd, "sknok"
            Exit Sub
        End If
        frmRoom.Show
    ElseIf vpengirim = 7 Then 'rekam vcd
        frmRoom.Enabled = True
        Unload Me
        DoEvents
        frmRoom.VideocapVCD
        frmRoom.Show
    ElseIf vpengirim = 8 Then 'rekam cd
        frmRoom.Enabled = True
        Unload Me
        DoEvents
        frmRoom.VideocapCD
        frmCamera.Width = 0
        frmRoom.Show
    ElseIf vpengirim = 9 Then 'Chat
        frmRoom.Enabled = True
        frmRoom.chatbuka
        chatConfirm ("'Terkoneksi !")
        Unload Me
        frmRoom.Show
    ElseIf vpengirim = 10 Then 'rekam HP
        frmRoom.Enabled = True
        Unload Me
        DoEvents
        frmRoom.VideocapHP
        frmRoom.Show
    ElseIf vpengirim = 11 Then 'rekam DVD
        frmRoom.Enabled = True
        Unload Me
        DoEvents
        frmRoom.VideocapDVD
        frmRoom.Show
    End If

    
End Sub

Private Sub prcNo()
    On Error Resume Next
    frmRoom.Enabled = True
    
    If vpengirim = 9 Then
        chatConfirm ("'No, thank you")
    ElseIf vpengirim = 0 Then
        vpbMember = ""
    End If

    Unload Me
End Sub

Private Sub prcOK()
    On Error Resume Next
    frmRoom.Enabled = True

    Unload Me
    If vpengirim = 1 Then
        frmRoom.txtLogin.SetFocus
    ElseIf vpengirim = 2 Then
        frmRoom.txtPass.text = ""
        frmRoom.txtPass.SetFocus
    ElseIf vpengirim = 12 Then
        Dim i As Integer
        Select Case vEffectAmpli
            Case 1
                For i = 1 To 5
                    RemoteAmpli "Preset1", 1
                    Sleep 3
                Next i
            Case 2
                For i = 1 To 5
                    RemoteAmpli "Preset2", 1
                    Sleep 3
                Next i
            Case 3
                For i = 1 To 5
                    RemoteAmpli "Preset3", 1
                    Sleep 3
                Next i
            Case 4
                For i = 1 To 5
                    RemoteAmpli "Preset4", 1
                    Sleep 3
                Next i
        End Select
        
    ElseIf vpengirim = 14 Then
    
        If vEffectAmpli = 1 Or vEffectAmpli = 2 Or vEffectAmpli = 3 Or vEffectAmpli = 4 Or vEffectAmpli = 5 Or vEffectAmpli = 6 Or vEffectAmpli = 7 Or vEffectAmpli = 8 Then
          For i = 1 To 5
            RemoteAmpli "Preset" & vEffectAmpli, 1
            Sleep 3
          Next
        ElseIf vEffectAmpli = 9 Or vEffectAmpli = 10 Or vEffectAmpli = 11 Or vEffectAmpli = 12 Then
          For i = 1 To 12
            RemoteAmpli "MICUP", 1
            Sleep 3
          Next
        End If
    ElseIf vpengirim = 15 Then
      frmDice2.tmr.Enabled = False
      Unload frmDice2
    End If
End Sub

Public Sub prcKonfirm()
    On Error Resume Next
    If vok = 1 Then
        prcYes
    ElseIf vok = 2 Then
        prcNo
    ElseIf vok = 3 Then
        prcOK
    End If
    Exit Sub
End Sub

Sub prcLoginMember()
    On Error Resume Next
    
    If vpengirim = 0 Then
        If vfokus = 1 Then
            Text2.SetFocus
            Exit Sub
        End If
    End If

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT     id, nama From tmember WHERE     (id = '" & Trim(Text1.text) & "')AND(nama = '" & Trim(Text2.text) & "')", _
       KoneksiAdoDBVizta, adOpenKeyset, adLockOptimistic
           
    If rs.recordCount > 0 Then
        frmRoom.Enabled = True
        vpbMember = Trim(rs.Fields(0))
        rs.Close
        Set rs = Nothing
                           
        'frmRoom.lstPlaylist.ListItems.Clear
        Dim LV As ListItem
        Dim Sql As String
        Dim myrs As MYSQL_RS
            Sql = "select masters.title, masters.singer, mp.idmusic, masters.PATH, masters.ANALOG, masters.VOL from memberplaylist mp inner join masters " & _
                  "on mp.idmusic = masters.idmusic " & _
                  "WHERE mp.idmember = '" & vpbMember & " ';"
            Set myrs = MyConn.Execute(Sql)
            
            Do Until myrs.EOF
                Set LV = frmRoom.lstPlaylist.ListItems.add(, , myrs.Fields(0).value)
                LV.SubItems(1) = myrs.Fields(1).value
                LV.SubItems(2) = myrs.Fields(2).value
                LV.SubItems(3) = myrs.Fields(3).value
                LV.SubItems(4) = myrs.Fields(4).value
                LV.SubItems(5) = myrs.Fields(5).value
                myrs.MoveNext
            Loop
        myrs.CloseRecordset
        Set myrs = Nothing
        
        Unload Me
        
        frmRoom.ScreenSaverAktif = 0
        frmRoom.Maksimal
        DoEvents
        frmRoom.Show
        frmRoom.pencet_btnplaylist
        DoEvents
        frmRoom.prcHilangkanScrollbarPlaylist
    Else
        text3.text = "Login gagal"
        rs.Close
        Set rs = Nothing
        Text1.SetFocus
    End If
End Sub


Private Sub Text1_GotFocus()
    On Error Resume Next
    vfokus = 1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
End Sub

Private Sub Text2_GotFocus()
    On Error Resume Next
    vfokus = 2
    Text2.SelStart = 0
    Text2.SelLength = Len(Text1.text)
End Sub

Private Sub tmrAktif_Timer()
    On Error Resume Next

    ScreenSaverAktifConfirmasi = ScreenSaverAktifConfirmasi + 1

    If vpengirim = 14 Then
    
        If ScreenSaverAktifConfirmasi = 3 Then
          
            tmrAktif.Enabled = False
            prcKonfirm
            frmRoom.Enabled = True
            Unload Me
        End If
        
    ElseIf vpengirim = 15 Then
    
        If ScreenSaverAktifConfirmasi = 5 Then
            tmrAktif.Enabled = False
            frmRoom.Enabled = True
            Unload Me
        End If
        
    Else
    
        If ScreenSaverAktifConfirmasi > 100 Then ScreenSaverAktifConfirmasi = 100
        
        If ScreenSaverAktifConfirmasi = 5 Then
            tmrAktif.Enabled = False
            frmRoom.Enabled = True
            Unload Me
        End If
        
    End If
End Sub
