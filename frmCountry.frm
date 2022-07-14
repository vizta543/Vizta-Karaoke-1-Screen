VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmCountry 
   Appearance      =   0  'Flat
   BackColor       =   &H0000244C&
   BorderStyle     =   0  'None
   Caption         =   "frmCountry"
   ClientHeight    =   10905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAktif 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin MSComctlLib.ListView lstCountry 
      Height          =   6435
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   9292
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmCountry.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TRANSPARAN FORM
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hWnd As Long, _
                    ByVal nIndex As Long) As Long
     
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hWnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long
                    
    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                    ByVal hWnd As Long, _
                    ByVal crKey As Long, _
                    ByVal bAlpha As Byte, _
                    ByVal dwFlags As Long) As Long
     
    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
'-----------------------------

'DISABLE SCROLLBAR
    Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, _
        ByVal wBar As Long, ByVal bShow As Long) As Long
    Private Const SB_HORZ = 0
    Private Const SB_VERT = 1
    Private Const SB_BOTH = 3
'-----------------------------

'Repaint Objek Form
    Private Declare Function LockWindowUpdate Lib _
        "user32" (ByVal hWndLock As Long) As Long
'-----------------------------

'Posisi Form
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
'------------------------------
    
Dim vvideotemp As Integer
Dim ScreenSaverAktifCountry As Integer
Dim lErr As Long

Private Sub Form_Load()
    On Error Resume Next
    
    vpbFrmCountry = True
    ScreenSaverAktifCountry = 0

    Dim lokasi As String
    lokasi = App.Path
    
    'edited by Andi 07/07/2022
    If frmUser.settingScreenResolution = "S-SD" Then
        If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Then
            lstCountry.Left = 1000
            lstCountry.Height = 4995
            lstCountry.Top = 1860
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf (frmRoom.vVideo = 5) Then
            lstCountry.Left = 1000
            lstCountry.Height = 4995
            lstCountry.Top = 1860
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf frmRoom.vVideo = 3 Then
            Skin1.LoadSkin lokasi + "\skin\sknchatuser.skn"
            Skin1.ApplySkinByName hWnd, "sknchatuser"
            lstCountry.Picture = LoadPicture(lokasi + "\Picture\chat\layar.jpg")
        End If
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Then
            lstCountry.Left = 1000
            lstCountry.Height = 3650
            lstCountry.Top = 1890
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf (frmRoom.vVideo = 5) Then
            lstCountry.Left = 1000
            lstCountry.Height = 4995
            lstCountry.Top = 1860
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf frmRoom.vVideo = 3 Then
            Skin1.LoadSkin lokasi + "\skin\sknchatuser.skn"
            Skin1.ApplySkinByName hWnd, "sknchatuser"
            lstCountry.Picture = LoadPicture(lokasi + "\Picture\chat\layar.jpg")
        End If
    ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
        If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Then
            lstCountry.Left = 1000
            lstCountry.Height = 4995
            lstCountry.Top = 1860
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf (frmRoom.vVideo = 5) Then
            lstCountry.Left = 1000
            lstCountry.Height = 4995
            lstCountry.Top = 1860
            lstCountry.ForeColor = &HFFFFFF
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
            tmrAktif.Enabled = True
        ElseIf frmRoom.vVideo = 3 Then
            Skin1.LoadSkin lokasi + "\skin\sknchatuser.skn"
            Skin1.ApplySkinByName hWnd, "sknchatuser"
            lstCountry.Picture = LoadPicture(lokasi + "\Picture\chat\layar.jpg")
        End If
    End If
    'edited by Andi 07/07/2022
    
'    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
'    SWP_NOMOVE + SWP_NOSIZE
    
    frmRoom.Enabled = False
    
    isiListCountry
    
    vvideotemp = frmRoom.vVideo
    frmRoom.vVideo = 13
End Sub

Private Sub isiListCountry()
    On Error Resume Next
    
    Dim sqlt As String
    Dim MyRst As MYSQL_RS
    Dim chm As ColumnHeader
    Dim LV As ListItem
    
    If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Or (frmRoom.vVideo = 5) Then
        If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Then
            sqlt = "SELECT TYPE, TYPENAME FROM kategori order by type"
        ElseIf (frmRoom.vVideo = 5) Then
            sqlt = "SELECT id, negara FROM filmnegara order by id"
        End If
        
        Set MyRst = MyConn.Execute(sqlt)
        lstCountry.ColumnHeaders.Clear
        lstCountry.ListItems.Clear
        
        If frmUser.settingScreenResolution = "S-SD" Then
            '
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            If frmRoom.vVideo = 0 Then
                lstCountry.Font.Size = 14
                lstCountry.Top = 2600
            ElseIf frmRoom.vVideo = 5 Then
                lstCountry.Font.Size = 14
                lstCountry.Top = 2900
            End If
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            lstCountry.Font.Size = 24
            lstCountry.Height = 5500
        End If
        
        Set chm = lstCountry.ColumnHeaders.add(, , , 10)
        Set chm = lstCountry.ColumnHeaders.add(, , , 3400)
        
        lstCountry.ColumnHeaders(2).Alignment = lvwColumnCenter
        
        Set LV = lstCountry.ListItems.add(, , ("0"))
        LV.SubItems(1) = "ALL"
            
        Do Until MyRst.EOF
            'Isi list data
            Set LV = lstCountry.ListItems.add(, , (MyRst.Fields(0).value))
            LV.SubItems(1) = MyRst.Fields(1).value
            MyRst.MoveNext
        Loop
    ElseIf frmRoom.vVideo = 3 Then
        sqlt = "SELECT idroom, userroom FROM room where status = 'chekin' order by userroom"
        Set MyRst = MyConn.Execute(sqlt)
        lstCountry.ColumnHeaders.Clear
        lstCountry.ListItems.Clear
        Set chm = lstCountry.ColumnHeaders.add(, , , 10)
        Set chm = lstCountry.ColumnHeaders.add(, , , 3400)
        
        lstCountry.ColumnHeaders(2).Alignment = lvwColumnCenter
        Do Until MyRst.EOF
            If Not MyRst.Fields(0).value = frmRoom.txtCompName.Tag Then
                Set LV = lstCountry.ListItems.add(, , (MyRst.Fields(0).value))
                LV.SubItems(1) = MyRst.Fields(1).value
            End If
            MyRst.MoveNext
        Loop
    End If
End Sub

Public Sub prcOK()
    On Error Resume Next
    
    frmRoom.vVideo = vvideotemp
    frmRoom.Enabled = True
    If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 1) Or (frmRoom.vVideo = 5) Then
        frmRoom.picKategori.Visible = False
        If lstCountry.selectedItem.index = 0 Then
            frmRoom.cbokategori.ListIndex = 0
        Else
            frmRoom.cbokategori.ListIndex = CInt(lstCountry.selectedItem.text)
        End If
        Select Case frmRoom.vVideo
            Case 0
                frmRoom.Maksimal
            Case 1
                frmRoom.midiKategori
            Case 5
                frmRoom.Maksimal
                frmRoom.moviecari
        End Select
        frmRoom.txtCategory.text = frmRoom.cbokategori.text
    ElseIf frmRoom.vVideo = 3 Then
        If lstCountry.ListItems.Count > 0 Then
            frmRoom.txtChatAktif.text = lstCountry.selectedItem.SubItems(1)
            frmRoom.txtChatAktif.Tag = CInt(lstCountry.selectedItem.text)
        End If
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    vpbFrmCountry = False
End Sub

Private Sub lstCountry_GotFocus()
    On Error Resume Next
    ShowScrollBar lstCountry.hWnd, SB_VERT, False
End Sub

Private Sub lstCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    ScreenSaverAktifCountry = 0
    lErr = LockWindowUpdate(lstCountry.hWnd)
    ShowScrollBar lstCountry.hWnd, SB_VERT, False
End Sub

Private Sub tutup()
    On Error Resume Next
    frmRoom.vVideo = vvideotemp
    frmRoom.Enabled = True
    Unload Me
    frmRoom.Minimal
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

Private Sub lstCountry_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ShowScrollBar lstCountry.hWnd, SB_VERT, False
    lErr = LockWindowUpdate(0)
End Sub

Private Sub tmrAktif_Timer()
    On Error Resume Next

    ScreenSaverAktifCountry = ScreenSaverAktifCountry + 1
    If ScreenSaverAktifCountry > 100 Then ScreenSaverAktifCountry = 100
    
    If ScreenSaverAktifCountry = 7 Then
        tmrAktif.Enabled = False
        tutup
    End If
End Sub
