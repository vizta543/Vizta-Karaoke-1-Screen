VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Begin VB.Form frmWelcome 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmWelcome"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsTombol 
      Height          =   2100
      Left            =   10560
      TabIndex        =   1
      Top             =   9000
      Width           =   4335
      _cx             =   7655
      _cy             =   3704
      FlashVars       =   ""
      Movie           =   "D:\Project\VOD\Source\Source\potongan\Intro\tombol.swf"
      Src             =   "D:\Project\VOD\Source\Source\potongan\Intro\tombol.swf"
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsLogo 
      Height          =   3075
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      _cx             =   9922
      _cy             =   5424
      FlashVars       =   ""
      Movie           =   "c:\new.swf"
      Src             =   "c:\new.swf"
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
      Left            =   12480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   10200
      Width           =   1215
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub flsLogo_GotFocus()

    On Error Resume Next
    
    Text1.SetFocus

    If Err.Number <> 0 Then
      LogError Name, "flsLogo_GotFocus"
    End If
End Sub

Private Sub flsTombol_GotFocus()

    On Error Resume Next
    
    StartKaraoke

    If Err.Number <> 0 Then
      LogError Name, "flsTombol_GotFocus"
    End If
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    vpbfrmWelcome = True
    frmRoom.vVideo = 8
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    
    flsLogo.Top = 0
    flsLogo.Left = 0
    flsLogo.Width = Me.Width
    flsLogo.Height = Me.Height
    
    frmRoom.Enabled = False
    
    Dim lokasi As String
    lokasi = App.Path + "\Picture\intro\"
    flsLogo.Movie = lokasi + "intro.swf"
    flsTombol.Movie = lokasi + "tombol.swf"
    
    frmUser.Timer1.Enabled = False
    SetCursorPos Screen.Width, 0


    If Err.Number <> 0 Then
      LogError Name, "Form_Load"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    frmRoom.vVideo = 0
    vpbfrmWelcome = False
    frmRoom.Enabled = True
    frmRoom.vHabisWaktu = False
    
    
    If Err.Number <> 0 Then
      LogError Name, "Form_Unload"
    End If
End Sub

Sub StartKaraoke()

    On Error Resume Next
    
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim waktu As String
    Dim WaktuCheckIn As String
    Dim idorder As String
    Dim PERINTAHDOS As String
    Dim remoteName As String
    Dim USER As String
    
    vpbMember = ""
    'Check Status Room
    Sql = "SELECT apa from room where ROOMNAME = '" & frmRoom.txtCompName.text & "'"
    Set myrs = MyConn.Execute(Sql)
    DoEvents
    
    If (myrs.Fields(0).value = "buka") Or (myrs.Fields(0).value = "welcome") Then
        waktu = Str$(year(Now)) & "-" & Trim$(Str$(month(Now))) & "-" & Trim$(Str$(day(Now))) & " " & hour(Now) & ":" & minute(Now) & ":" & second(Now)
    
        'UPDATE WAKTU MYSQL
        Sql = "UPDATE room SET WKTSTART = '" & waktu & "', apa = 'buka'  WHERE ROOMNAME ='" & frmRoom.txtCompName & "'"
        Set myrs = MyConn.Execute(Sql)
        DoEvents
    
        frmRoom.Enabled = True
    
        'UPDATE DURASI MSSQL
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT     MAX(torder_room.tglStart) AS checkin FROM torder_room INNER JOIN  troom ON torder_room.idorder = troom.idorder" & _
                " WHERE     (troom.namaroom = '" & Trim$(frmRoom.txtCompName) & "')", KoneksiAdoDBVizta, adOpenKeyset, adLockOptimistic
        DoEvents
        If IsNull(rs.Fields(0).value) = False Then
          WaktuCheckIn = Str$(year(rs.Fields(0).value)) & "-" & Trim$(Str$(month(rs.Fields(0).value))) & "-" & Trim$(Str$(day(rs.Fields(0).value))) & " " & hour(rs.Fields(0).value) & ":" & minute(rs.Fields(0).value) & ":" & second(rs.Fields(0).value)
          KoneksiAdoDBVizta.Execute "UPDATE    torder_room SET  tglStart = '" & Trim$(waktu) & "'   WHERE     (tglStart = '" & WaktuCheckIn & "')"
          DoEvents
          KoneksiAdoDBVizta.Execute "UPDATE    troom SET  waktu = '" & Trim$(waktu) & "' WHERE     (namaroom = '" & Trim$(frmRoom.txtCompName) & "')"
          DoEvents
        End If
        rs.Close

        'UPDATE SERVER BACKUP
        Dim vpServerTemp As String
        If AktifServerStatus = 1 Then
            vpServerTemp = vpbServerBackup
        Else
            vpServerTemp = vpbServerUtama
        End If
        
        Dim ECHO As ICMP_ECHO_REPLY
        Call Ping(vpServerTemp, ECHO)
        If ECHO.status = 0 Then
            If MyConnBackup.State = MY_CONN_OPEN Then
                Set myrs = MyConnBackup.Execute(Sql)
            Else
                MyConnBackup.OpenConnection vpServerTemp, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
                If MyConnBackup.State = MY_CONN_OPEN Then
                    Set myrs = MyConnBackup.Execute(Sql)
                End If
            End If
            DoEvents
        End If
        
        Unload Me
    
        frmRoom.Enabled = True
        frmRoom.txtSearch.SetFocus
        ScoreSetup = True
    End If
    
lblEnd:
    
    If Err.Number <> 0 Then
      LogError Me.Name, "StartKaraoke"
    End If
    
End Sub

Sub BukaKaraoke()

    On Error Resume Next
    
    Unload Me
    
    frmRoom.Enabled = True
    frmRoom.txtSearch.SetFocus
    ScoreSetup = True
    
lblEnd:
    
    If Err.Number <> 0 Then
      LogError Me.Name, "BukaKaraoke"
    End If
    
End Sub
