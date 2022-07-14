Attribute VB_Name = "ModKaraoke"
Option Explicit

'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'Global Koneksi
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function getfullscreen Lib "wmp.dll" () As Long
Public Declare Function API Lib "chronoapi.dll" () As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function addNetworkConnection Lib "Vizta Library.dll" (ByVal remoteName As String, ByVal remoteNameLength As Long, ByVal USER As String, ByVal userLength As Long, ByVal password As String, ByVal passwordLength As Long) As Long
Public Declare Function cancelNetworkConnection Lib "Vizta Library.dll" (ByVal remoteName As String, ByVal remoteNameLength As Long) As Long
Public Declare Function setAudioEndPointVolumeMasterVolumeLevelPercent Lib "Vizta Library.dll" (ByVal volumeLevelPercent As Long) As Long
Public Declare Function setAudioEndPointVolumeMute Lib "Vizta Library.dll" (ByVal mute As Long) As Long
Public Declare Function setTime Lib "Vizta Library.dll" (ByVal year As Long, ByVal month As Long, ByVal day As Long, ByVal hour As Long, ByVal minute As Long, ByVal second As Long) As Long

'For Slide Show
Declare Function SystemParametersInfo& Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByVal lpvParam As Any, ByVal fuWinIni&)
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'FREE DISK SPACE
    Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
    (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector _
    As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
'----------------------

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFF

Public Const REG_SZ = 1
Public Const HKEY_CURRENT_USER = &H80000001
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2

Global MyConn As New MYSQL_CONNECTION
Global MyConnBackup As New MYSQL_CONNECTION
Private Const DB_NAME As String = "karaoke"

Public AktifServerStatus As Integer
Public ScoreValid As Boolean
Public ScoreSetup As Boolean
Public vpbServer As String
Public vpbServerUtama As String
Public vpbServerBackup As String
Public vpbServerKeyWindows As String
Public vpbServerKeyMySQL As String
Public vpbBlackBox As Integer '0=off 1=on 2=auto

Public katajalan() As String
Public katajalantotal As Integer
Public katajalanx As Integer

'Variabel Aktif Form
Public vpbFrmCountry As Boolean
Public vpbFrmCategory As Boolean
Public vpbFrmConfirmasi As Boolean
Public vpbFrmCall As Boolean
Public vpbfrmTambahJam As Boolean
Public vpbfrmWelcome As Boolean
Public vpbfrmSaran As Boolean
Public vpbfrmNew As Boolean
Public vpbfrmPopuler As Boolean
Public vpbfrmAbout As Boolean
Public vpbfrmHelp As Boolean
Public vpbfrmVocal As Boolean
Public vpbfrmCamera As Boolean
Public vpbfrmLoading As Boolean


'variabel  status karaoke
Public vpbRoomStatus As Integer '0=tutup, 1=welcome 3=buka
'Variabel Umum
Public vpbMute As Boolean
Public vpbNamaKomputer As String
Public vpbHits As Boolean
Public vpbPopuler As Boolean
Public vpbNew As Boolean
Public vpbBolehMainVocal As Boolean
Public vTitleArtisPlaylist As Boolean

Public vpbAmpliAuto As Integer
Public vpbAmpliMinus As Integer
Public vEffectAmpli As Integer
Public vpbMember As String
Public vpbProsesEksekusi As Integer
Public vpbRemoteStatus As Integer
Private vPitch As Integer

Public KoneksiAdoDB As ADODB.Connection
Public KoneksiAdoDBVizta As ADODB.Connection

Public Sub KonekServer()

    On Error Resume Next
    
    
    If Not KoneksiAdoDB Is Nothing Then
      KoneksiAdoDB.Close
      Set KoneksiAdoDB = Nothing
    End If
    
    Set KoneksiAdoDB = New ADODB.Connection
    KoneksiAdoDB.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
            & "SERVER=" & vpbServer & ";" _
            & "DATABASE=karaoke;" _
            & "UID=karaoke;" _
            & "PWD=" & vpbServerKeyMySQL & ";" _
            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    KoneksiAdoDB.CursorLocation = adUseClient
    KoneksiAdoDB.Open
    
    
    If Not KoneksiAdoDBVizta Is Nothing Then
      KoneksiAdoDBVizta.Close
      Set KoneksiAdoDBVizta = Nothing
    End If
    
    Set KoneksiAdoDBVizta = New ADODB.Connection
    KoneksiAdoDBVizta.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
            & "SERVER=" & vpbServer & ";" _
            & "DATABASE=vizta;" _
            & "UID=karaoke;" _
            & "PWD=" & vpbServerKeyMySQL & ";" _
            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    KoneksiAdoDBVizta.CursorLocation = adUseClient
    KoneksiAdoDBVizta.Open
    
    
    If Err.Number <> 0 Then
        LogError "ModKaraoke", "KonekServer"
    End If
End Sub

'Use database
Private Sub UseDb()
    On Error Resume Next
    Dim Sql As String
    Sql = "USE " & DB_NAME & ";"
    MyConn.Execute Sql
End Sub

'Mengecek apakah database exist
Private Function Dbada() As Boolean
    On Error Resume Next
    Dim Sql As String, RcdAffected As Long
    Dim RcdSet As MYSQL_RS
    Dbada = False
    Sql = "SHOW DATABASES;"
    Set RcdSet = MyConn.Execute(Sql, RcdAffected)
    If RcdSet.recordCount = 0 Then Exit Function
    RcdSet.MoveFirst
    Do Until RcdSet.EOF
        If RcdSet.Fields(0).value = DB_NAME Then
            Dbada = True
            Exit Do
        End If
        RcdSet.MoveNext
    Loop
    RcdSet.CloseRecordset
    Set RcdSet = Nothing
End Function

'pesan error apabila ada kesalahan koneksi
Public Sub ShowError()
    On Error Resume Next
    MsgBox "Error " & MyConn.Error.Number & ": (" & MyConn.Error.Description & ")", vbCritical, "ADA KESALAHAN"
End Sub

Public Function tempPitch(deltaCents As Integer)
    On Error Resume Next
    Dim acc As Object
    Set acc = CreateObject("chronoapi.Access")
    
    Dim oAPI As Object
    Set oAPI = acc.API
    
    vPitch = oAPI.RateToCents(oAPI.Pitch)
    If (vPitch > -50) And (vPitch < 50) Then
        vPitch = 0
    End If
    If (vPitch > 50) And (vPitch < 150) Then
        vPitch = 100
    End If
    If (vPitch > 150) And (vPitch < 250) Then
        vPitch = 200
    End If
    If (vPitch > 250) And (vPitch < 350) Then
        vPitch = 300
    End If
    If (vPitch > 350) And (vPitch < 450) Then
        vPitch = 400
    End If
    If (vPitch > 450) And (vPitch < 550) Then
        vPitch = 500
    End If
    If (vPitch > 550) And (vPitch < 650) Then
        vPitch = 600
    End If
    If (vPitch > -150) And (vPitch < -50) Then
        vPitch = -100
    End If
    If (vPitch > -250) And (vPitch < -150) Then
        vPitch = -200
    End If
    If (vPitch > -350) And (vPitch < -250) Then
        vPitch = -300
    End If
    If (vPitch > -450) And (vPitch < -350) Then
        vPitch = -400
    End If
    If (vPitch > -550) And (vPitch < -450) Then
        vPitch = -500
    End If
    If (vPitch > -650) And (vPitch < -550) Then
        vPitch = -600
    End If
    
    oAPI.CentsToRate (deltaCents)
End Function

'Function Untuk Key Up (#) dan Key Down (b)
Public Function incPitch(deltaCents As Integer)
'   Dim acc As New chronoapi.CAccess
'   On Error Resume Next
'   Set API = acc.API
'   If API Is Nothing Then
'     Exit Function
'   Else
'     API.Pitch = API.CentsToRate(API.RateToCents(API.Pitch) + deltaCents)
'   End If

    'Win7
    On Error Resume Next
    Dim acc As Object
    Set acc = CreateObject("chronoapi.Access")
    
    Dim oAPI As Object
    Set oAPI = acc.API

    oAPI.Pitch = oAPI.CentsToRate(vPitch + deltaCents)
End Function

'Fungsi Reset Pitch
Public Function SetPitch(cents As Integer)

    On Error Resume Next
    
    Dim acc As Object
    Dim oAPI As Object
    
    Set acc = CreateObject("chronoapi.Access")
    If Not acc Is Nothing Then
        Set oAPI = acc.API
        If Not oAPI Is Nothing Then
          oAPI.Pitch = oAPI.CentsToRate(cents)
        End If
    End If
    
    
    If Err.Number <> 0 Then
      LogError "ModKaraoke", "SetPitch"
    End If
End Function

'Fungsi Tempo Up dan Tempo Down
Public Function IncTempo(deltaRate)
'    Dim acc As New chronoapi.CAccess
'    On Error Resume Next
'    Set API = acc.API
'        If API Is Nothing Then
'            Exit Function
'        Else
'            API.Tempo = API.Tempo + deltaRate
'        End If
'    frmVideo.pbTempo = frmVideo.pbTempo + deltaRate
    On Error Resume Next
    
    Dim acc As Object
    Set acc = CreateObject("chronoapi.Access")
    
    Dim oAPI As Object
    Set oAPI = acc.API
    oAPI.Tempo = oAPI.Tempo + deltaRate
    frmVideo.pbTempo = frmVideo.pbTempo + deltaRate
End Function

'Reset Tempo
Public Function SetTempo(rateSet)
    
    On Error Resume Next
    
    Dim acc As Object
    Dim oAPI As Object
    
    
    Set acc = CreateObject("chronoapi.Access")
    If Not acc Is Nothing Then
        Set oAPI = acc.API
        If Not oAPI Is Nothing Then
          oAPI.Tempo = rateSet
        End If
    End If
    
    
    If Err.Number <> 0 Then
      LogError "ModKaraoke", "SetTempo"
    End If
End Function
    
Public Function setFullScreen()
    On Error Resume Next
    If frmVideo.WindowsMediaPlayer1.openState = wmposEndCodecAcquisition Then
        frmVideo.WindowsMediaPlayer1.fullScreen = True
    ElseIf frmVideo.WindowsMediaPlayer1.openState = wmposMediaOpen Then
        frmVideo.WindowsMediaPlayer1.fullScreen = True
    ElseIf frmVideo.WindowsMediaPlayer1.openState = wmposMediaChanging Then
        frmVideo.WindowsMediaPlayer1.fullScreen = True
    End If
End Function

Public Sub Path()
    On Error Resume Next
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Sql = "select pathserv from path"
    Set myrs = MyConn.Execute(Sql)
End Sub

Public Sub HilangkanDriveServer()
    
  On Error Resume Next
  
  Dim remoteName As String
  
  remoteName = "\\" & vpbServerUtama
  cancelNetworkConnection remoteName, Len(remoteName)
  
  remoteName = "\\" & vpbServerBackup
  cancelNetworkConnection remoteName, Len(remoteName)
End Sub

Public Sub prcTutupAktifForm()
    On Error Resume Next
        'Tutup aktif form
        If vpbFrmCountry = True Then
            Unload frmCountry
        End If
        If vpbFrmCategory = True Then
            Unload frmCategory
        End If
        If vpbFrmConfirmasi = True Then
            Unload frmConfirmasi
        End If
        If vpbFrmCall = True Then
            Unload frmCall
        End If
        If vpbfrmSaran = True Then
            Unload frmSaran
        End If
        If vpbfrmTambahJam = True Then
            Unload frmTambahJam
            frmRoom.hotnon
        End If
        If vpbfrmNew = True Then
            Unload frmNew
        End If
        If vpbfrmPopuler = True Then
            Unload frmPopuler
        End If
        If vpbfrmVocal = True Then
            Unload frmVocal
        End If
        If vpbfrmHelp = True Then
            Unload frmHelp
        End If
End Sub

Public Sub prcTitleList()

    On Error Resume Next
    
    vTitleArtisPlaylist = False
    If frmRoom.vpointer = 3 Then
        vpbHits = False
        vpbNew = False
        vpbPopuler = False
        If vpbNew = True Or vpbHits = True Or vpbPopuler = True Then
            frmRoom.CariTitle
        End If
        frmRoom.vpointer = 1
        frmRoom.cmdSong_Click
        vTitleArtisPlaylist = True
    End If
    frmRoom.txtSearch.SetFocus
    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    
    If vTitleArtisPlaylist = False Then
        If frmUser.settingScreenResolution = "S-SD" Then
            If frmRoom.vpointer = 1 Then
                DoEvents
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\abr-title"
                frmRoom.vpointer = 7
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\title"
                frmRoom.vpointer = 1
            End If
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            If frmRoom.vpointer = 1 Then
                DoEvents
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\titlesingkatan"
                frmRoom.vpointer = 7
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\titlev2"
                frmRoom.vpointer = 1
            End If
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            If frmRoom.vpointer = 1 Then
                DoEvents
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\titlesingkatan"
                frmRoom.vpointer = 7
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\titlev2"
                frmRoom.vpointer = 1
            End If
        End If
        frmRoom.txtSearch_Change
    End If
    frmRoom.txtSearch.SelStart = 0
    frmRoom.txtSearch.SelLength = Len(frmRoom.txtSearch.text)
    vTitleArtisPlaylist = False
    
    If Err.Number <> 0 Then
      LogError "ModKaraoke", "prcTitleList"
    End If
End Sub

Public Sub prcSingerList()

    On Error Resume Next
    
    vTitleArtisPlaylist = False
    If frmRoom.vpointer = 3 Then
        vpbHits = False
        vpbNew = False
        vpbPopuler = False
        If vpbNew = True Or vpbHits = True Or vpbPopuler = True Then
            frmRoom.CariTitle
        End If
        frmRoom.vpointer = 2
        frmRoom.cmdSong_Click
        vTitleArtisPlaylist = True
    End If
    frmRoom.txtSearch.SetFocus
    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    If vTitleArtisPlaylist = False Then
        If frmUser.settingScreenResolution = "S-SD" Then
            If frmRoom.vpointer = 2 Then
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\abr-artist"
                frmRoom.vpointer = 8
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artist"
                frmRoom.vpointer = 2
            End If
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            If frmRoom.vpointer = 2 Then
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artistsingkatanv2"
                frmRoom.vpointer = 8
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artistv2"
                frmRoom.vpointer = 2
            End If
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            If frmRoom.vpointer = 2 Then
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artistsingkatanv2"
                frmRoom.vpointer = 8
            Else
                frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artistv2"
                frmRoom.vpointer = 2
            End If
        End If
        frmRoom.txtSearch_Change
    End If
    frmRoom.txtSearch.SelStart = 0
    frmRoom.txtSearch.SelLength = Len(frmRoom.txtSearch.text)
    vTitleArtisPlaylist = False
    
    If Err.Number <> 0 Then
      LogError "ModKaraoke", "prcSingerList"
    End If
End Sub

Public Sub prcSonglist()
    On Error Resume Next
    If vpbNew = True Or vpbHits = True Or vpbPopuler = True Then
        vpbHits = False
        vpbNew = False
        vpbPopuler = False
        frmRoom.CariTitle
    End If
    vpbHits = False
    vpbNew = False
    vpbPopuler = False
    frmRoom.cmdSong_Click
End Sub

Public Sub RemoteAmpliAktif()
    
    On Error Resume Next
    
    If frmUser.settingAmpliCOM = 0 Then
      If vpbAmpliAuto > 0 Then
          Dim strLocalIP As String
          frmRoom.WskRemote.Close
          strLocalIP = "127.0.0.1"
          frmRoom.WskRemote.Protocol = sckTCPProtocol
          frmRoom.WskRemote.RemoteHost = strLocalIP
          frmRoom.WskRemote.RemotePort = 8765
          frmRoom.WskRemote.Connect
      End If
    End If
    
hell:
    
    If Err.Number <> 0 Then
      logThisError "RemoteAmpliAktif"
    End If
    
End Sub

Public Sub RemoteAmpli(tombol As String, times_to_repeat As String)
    
    On Error Resume Next
    
    If frmUser.settingAmpliCOM = 0 Then
      
      Dim sendstr As String
      Dim password, remoteName, buttonname As String
      
      If (vpbAmpliAuto = 0) Then
          Exit Sub
      End If
      
      On Error GoTo aktifkan
      
      password = "kancut"
      remoteName = "REMOTEAMPLI"
      buttonname = tombol
      'times_to_repeat = "9"
      sendstr = (password & " " & remoteName & " " & buttonname & " " & times_to_repeat & vbLf)
      frmRoom.WskRemote.SendData (sendstr)
      
      Exit Sub
      
aktifkan:

      RemoteAmpliAktif
      
    Else
      frmUser.sendUSBAudioCommand tombol
    End If
    
End Sub

Public Sub resetAmpli()

    On Error Resume Next

    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim i, j, k, l As Integer
    
    If vpbAmpliAuto > 0 Then
        If frmRoom.WskRemote.State = sckClosed Then
            RemoteAmpliAktif
            DoEvents
        End If
        
        If vpbAmpliAuto = 1 Then
            If MyConn.State = MY_CONN_CLOSED Then
                If AktifServerStatus = 1 Then
                    frmRoom.konekServer1
                Else
                   frmRoom.konekServer2
                End If
                DoEvents
            End If
            
            Sql = "SELECT remotevol, remotemic, remoteecho FROM room where ROOMNAME = '" & frmRoom.txtCompName.text & "'"
            Set myrs = MyConn.Execute(Sql)
            DoEvents
            
            'NOLKAN SEMUA VOLUME
            For i = 1 To vpbAmpliMinus
                RemoteAmpli "VOLDN", 0
                DoEvents
            Next i
            DoEvents
    
            For i = 1 To vpbAmpliMinus
                RemoteAmpli "MICDN", 0
                DoEvents
            Next i
            DoEvents
            For i = 1 To vpbAmpliMinus
                RemoteAmpli "ECHODN", 0
                Sleep 5
                DoEvents
            Next i
            DoEvents
                   
            'SESUAIKAN DENGAN SETTINGAN
            j = myrs.Fields(0).value
            For i = 1 To j
                RemoteAmpli "VOLUP", 0
                Sleep 5
                DoEvents
            Next i
            DoEvents
            k = myrs.Fields(1).value
            For i = 1 To k
                RemoteAmpli "MICUP", 0
                DoEvents
            Next i
            DoEvents
            l = myrs.Fields(2).value
            For i = 1 To l
                RemoteAmpli "ECHOUP", 0
                Sleep 5
                DoEvents
            Next i
        ElseIf vpbAmpliAuto = 2 Then
            Sleep 5
            For i = 1 To 5
                RemoteAmpli "ECHODN", 0
                Sleep 5
                DoEvents
            Next i
            For i = 1 To 5
                RemoteAmpli "Preset2", 1
                Sleep 3
                DoEvents
            Next i
        ElseIf vpbAmpliAuto = 3 Then
            For i = 1 To 5
                RemoteAmpli "ECHOUP", 0
                Sleep 5
                DoEvents
            Next i
            
            For i = 1 To 5
                RemoteAmpli "PRESET0", 1
                Sleep 3
                DoEvents
            Next i
            vEffectAmpli = 1
        End If
        DoEvents
        Unload frmLoading
        frmRoom.tmrLstAllLoad_Timer
    Else
        Unload frmLoading
        frmRoom.tmrLstAllLoad_Timer
    End If
    
End Sub

Private Sub logThisError(procedureName As String)
  LogError "ModKaraoke", procedureName
End Sub
