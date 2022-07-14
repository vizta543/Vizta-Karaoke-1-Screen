VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmUser"
   ClientHeight    =   3525
   ClientLeft      =   4200
   ClientTop       =   2190
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm usbRCCom 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock wsPassword 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cmdShutDwn"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRoomId 
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtIdRoom 
      Height          =   285
      Left            =   7080
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRoom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "frmUser.frx":000C
      Top             =   120
   End
   Begin MSCommLib.MSComm usbAudioCom 
      Left            =   840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Inp Lib "inpout32.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32.dll" _
Alias "Out32" (ByVal PortAddress As Integer, ByVal value As Integer)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Sub usb_relay_device_close Lib "usb_relay_device.dll" (ByVal handle As Long)
Private Declare Function usb_relay_device_close_all_relay_channel Lib "usb_relay_device.dll" (ByVal handle As Long) As Long
Private Declare Function usb_relay_device_close_one_relay_channel Lib "usb_relay_device.dll" (ByVal handle As Long, ByVal channel As Long) As Long
Private Declare Function usb_relay_device_enumerate Lib "usb_relay_device.dll" () As Long
Private Declare Function usb_relay_device_open Lib "usb_relay_device.dll" (ByVal usbrdi As Long) As Long
Private Declare Function usb_relay_device_open_all_relay_channel Lib "usb_relay_device.dll" (ByVal handle As Long) As Long
Private Declare Function usb_relay_device_open_one_relay_channel Lib "usb_relay_device.dll" (ByVal handle As Long, ByVal channel As Long) As Long
Private Declare Function usb_relay_exit Lib "usb_relay_device.dll" () As Long
Private Declare Function usb_relay_init Lib "usb_relay_device.dll" () As Long

Const VK_ESCAPE = &H1B
Const VK_F4 = &H73
Const VK_DELETE = &H2E
Const VK_RBUTTON = &H2
Const VK_LMENU = &HA4
Const VK_RMENU = &HA5
Const VK_NUMLOCK = &H90
Const VK_TAB = &H9

Public osVersion As String

Private Const rcCOMCodeChannel1 As Long = &H1
Private Const rcCOMCodeChannel2 As Long = &H2
Private Const rcCOMCodeChannel3 As Long = &H4
Private Const rcCOMCodeChannel4 As Long = &H8

Public settingRCType As String
Public settingRCCOMNumber As Long
Private Const settingRCCOMCodeChannelDiscoLamp As Long = 1
Private Const settingRCHIDNumber As Long = 1
Private Const settingRCHIDChannelDiscoLamp As Long = 1
Public settingRCLPTAddress As Long

Public rcCOMCode As Long
Public rcHIDDI As Long
Public rcHIDHandle As Long

'added by Andi 22-01-2021
Public settingScreenResolution As String
'added by Andi 22-01-2021

'added by Andi 21-04-2021
Public settingMirrorBall As Long
'added by Andi 21-04-2021

Private Const usbAudioModelK3200D = "K-3200D"
Private Const usbAudioModelK3200P = "K-3200P"
Private Const usbAudioModelK6000D = "K-6000D"
Private Const usbAudioModelK6800D = "K-6800D"

Public settingAmpliCOM As Long
Private usbAudioModel As String


Public settingStatusAutoAbbreviation As Long

Private wsPasswordBytesTotal As Long
Private wsPasswordDataLength As Long
Private wsPasswordChallenge As Boolean

Private Sub Command1_Click()
    On Error Resume Next
    HilangkanDriveServer
    frmAdmin1.Show
End Sub

Private Sub Form_Load()
    
    On Error Resume Next

    If App.PrevInstance = True Then
      End
    End If
    
    'added by Andi 22-01-2021
    settingScreenResolution = getIniSetting("Display", "Resolution", "0")
    'added by Andi 22-01-2021
    
    'added by Andi 21-04-2021
    settingMirrorBall = getIniSetting("MirrorBall", "LampByMusic", "0")
    'added by Andi 21-04-2021
    
    osVersion = mdlWinVer.GetWindowsVersion()
    
    settingStatusAutoAbbreviation = getIniSetting("Status", "AutoAbbreviation", "1")
    

    settingRCType = getIniSetting("RelayControl", "Type", "LPT")
    settingRCLPTAddress = getIniSetting("RelayControl", "LPTAddress", "888")
    settingRCCOMNumber = getIniSetting("RelayControl", "COMNumber", "3")
    
    Select Case settingRCType
    
      Case "COM"
        
        usbRCCom.CommPort = settingRCCOMNumber
        usbRCCom.SThreshold = 1
        
        rcCOMCode = &HFF
        
      Case "HID"
        
        rcHIDDI = 0
        rcHIDHandle = 0
        
        If usb_relay_init() <> 0 Then
          LogText "usb_relay_init fail"
        Else
          rcHIDDI = usb_relay_device_enumerate()
          If rcHIDDI = 0 Then
            LogText "usb_relay_device_enumerate fail"
          Else
            rcHIDHandle = usb_relay_device_open(rcHIDDI + ((settingRCHIDNumber - 1) * 16))
            If rcHIDHandle = 0 Then
              LogText "usb_relay_device_open fail"
            End If
          End If
        End If
        
      Case "LPT"
      
    End Select

    turnDiscoLampOff
    
    
    Dim lokasi As String

    Me.Move (0 - Screen.Width)
    lokasi = App.Path
    Skin1.LoadSkin lokasi + "\skin\user.skn"
    Skin1.ApplySkinByName hWnd, "UserForm"
    
    LockClient
    
    frmBackground.Show
    
    setAudioEndPointVolumeMasterVolumeLevelPercent 0
    
    Dim PERINTAHDOS As String
    Dim ECHO As ICMP_ECHO_REPLY
    
    regCreate_Key_Value &H80000001, _
                        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
                        "NoDrives", "98304"
    
    'Read
    vpbServerUtama = getIniSetting("DataSource", "Host", "10.0.0.201")
    vpbServerBackup = getIniSetting("DataSource", "Backup", "10.0.0.202")
    
    vpbAmpliAuto = getIniSetting("Ampli", "Auto", "0")
    vpbAmpliMinus = getIniSetting("Ampli", "Min", "0")
      
    settingAmpliCOM = getIniSetting("Ampli", "COM", "0")
    If settingAmpliCOM > 0 Then
    
      usbAudioCom.CommPort = settingAmpliCOM
      usbAudioCom.InputMode = comInputModeBinary
      usbAudioCom.SThreshold = 1
    
      usbAudioModel = getIniSetting("Ampli", "Model", usbAudioModelK3200D)
      Select Case usbAudioModel
        Case usbAudioModelK3200D
          usbAudioCom.settings = "38400,n,8,1"
        Case usbAudioModelK3200P
          usbAudioCom.settings = "38400,n,8,1"
        Case usbAudioModelK6000D
          usbAudioCom.settings = "38400,n,8,1"
        Case usbAudioModelK6800D
          usbAudioCom.settings = "115200,n,8,1"
      End Select
    End If
    
    vpbRemoteStatus = getIniSetting("Status", "Remote", "1")
    
    
    vpbServerKeyWindows = ""
    vpbServerKeyMySQL = ""
    
    wsPasswordBytesTotal = 0
    wsPasswordDataLength = 0
    wsPasswordChallenge = False
    
    wsPassword.RemotePort = modProject.outletServerPort
    
    wsPassword.Connect vpbServerUtama
    
    
    If Err.Number <> 0 Then
      LogError Name, "Form_Load"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    LockClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Timer1.Enabled = False
    UnlockClient
End Sub

Private Sub Skin1_Click(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next
  If Source.GetName = "ShutButton" Then
      Command1_Click
  End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Sql = "SELECT apa, userroom from room where ROOMNAME = '" & txtRoom.text & "'"
    Set myrs = MyConn.Execute(Sql)
    
    If MyConn.Error.Number <> 0 Then
        Timer1.Enabled = False
        If AktifServerStatus = 1 Then
            konekServer2
        Else
            konekServer1
        End If
        Exit Sub
    End If
    
    If (myrs.Fields(0).value = "buka") Or (myrs.Fields(0).value = "welcome") Or (myrs.Fields(0).value = "tutup") Then
        Me.Hide
        Timer1.Enabled = False
        frmRoom.Show
        frmRoom.txtUser.text = myrs.Fields(1).value
        frmRoom.cbokategori.ListIndex = 0
        SizeCombo frmRoom, frmRoom.cbokategori
        frmRoom.txtSearch.SetFocus
      '  frmRoom.AktifServerStatus = AktifServerStatus
        sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
        LockRoom
    ElseIf myrs.Fields(0).value = "T" Then 'tutup
        HilangkanDriveServer
        End
        Exit Sub
    ElseIf myrs.Fields(0).value = "S" Then    'shutdown
        HilangkanDriveServer
        Call Shell("Shutdown /s /t 0")
        End
        Exit Sub
    ElseIf myrs.Fields(0).value = "R" Then 'restart
        HilangkanDriveServer
        Call Shell("Shutdown /r /t 0")
        End
        Exit Sub
    End If
    Set myrs = Nothing
End Sub

Sub konekServer1()
    
    On Error Resume Next

    Dim ECHO As ICMP_ECHO_REPLY
    Dim PERINTAHDOS As String
    
    Call Ping(vpbServerUtama, ECHO)
    If ECHO.status = 0 Then
        MyConn.OpenConnection vpbServerUtama, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
        If MyConn.Error.Number = 0 Then
            vpbServer = vpbServerUtama
            KonekServer
            AktifServerStatus = 1
        Else
            konekServer2
        End If
    Else
        konekServer2
    End If
    
    If Err.Number <> 0 Then
      LogError Name, "konekServer1"
    End If
End Sub

Sub konekServer2()
    
    On Error Resume Next

    Dim ECHO As ICMP_ECHO_REPLY
    Dim PERINTAHDOS As String
    
    Call Ping(vpbServerBackup, ECHO)
    If ECHO.status = 0 Then
        MyConn.OpenConnection vpbServerBackup, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
        If MyConn.Error.Number = 0 Then
            vpbServer = vpbServerBackup
            KonekServer
            AktifServerStatus = 2
        Else
            konekServer1
        End If
    Else
        konekServer1
    End If
    
    If Err.Number <> 0 Then
        LogError Name, "konekServer2"
    End If
End Sub


Private Sub wsPassword_DataArrival(ByVal bytesTotal As Long)
  
  On Error Resume Next
  
  Dim dataLength As Long
  Dim Data As String * 32767
  Dim buffer As String * 32767
  Dim challengeCode As String * 127
  Dim byteArray() As Byte
  Dim Key As String
  Dim a As Long
  Dim ComName As String * 255, cname As String
  Dim x As Long
  Dim Sql As String
  Dim myrs As MYSQL_RS
  Dim result As Long

  If wsPasswordBytesTotal = 0 Then
    If wsPasswordDataLength = 0 Then
      wsPassword.GetData wsPasswordDataLength, vbLong, 4
      If bytesTotal > 4 Then
        wsPasswordBytesTotal = bytesTotal - 4
      End If
    Else
      wsPasswordBytesTotal = bytesTotal
    End If
  Else
    wsPasswordBytesTotal = wsPasswordBytesTotal + bytesTotal
  End If
  
  Sleep 2 ' do not remove, strange

  If wsPasswordBytesTotal = wsPasswordDataLength Then
    
    Data = Space$(16384)
    
    wsPassword.GetData Data, vbString, wsPasswordDataLength
    
    If wsPasswordChallenge = False Then
  
      buffer = Space$(16384)
      
      dataLength = opensslAESDecrypt(Data, wsPasswordDataLength, buffer, 1073741822, 1073741821)
      If dataLength <> 4 Then
        LogText "opensslAESDecrypt error: " & dataLength
        Exit Sub
      End If
      
      challengeCode = Space$(127)
      
      If opensslSHA512Digest(buffer, dataLength, challengeCode) <> 64 Then
        LogText "opensslSHA512Digest error"
        Exit Sub
      End If
      
      dataLength = 64
      wsPassword.SendData dataLength
      
      ReDim byteArray(63)
      For a = 0 To 63
        byteArray(a) = Asc(Mid(challengeCode, a + 1, 1))
      Next
      wsPassword.SendData byteArray
      
      
      buffer = "getKeyWindows"
      dataLength = 13
      wsPassword.SendData dataLength
      
      ReDim byteArray(dataLength - 1)
      For a = 0 To dataLength - 1
        byteArray(a) = Asc(Mid(buffer, a + 1, 1))
      Next
      wsPassword.SendData byteArray
      
      wsPasswordChallenge = True
    
    Else
      
      buffer = Space$(16384)
  
      result = mbedtlsBlowfishDecrypt(Data, wsPasswordDataLength, buffer, 1073741823, 1073741824)
      If result <> 2048 Then
        LogText "mbedtlsBlowfishDecrypt fail: " & result & ", " & Err.Description
        Exit Sub
      End If
      
      Key = ""
      dataLength = (Asc(Mid(buffer, 1, 1)) * 1) + (Asc(Mid(buffer, 2, 1)) * 256) + (Asc(Mid(buffer, 3, 1)) * 65536) + (Asc(Mid(buffer, 4, 1)) * 16777216)
      For a = 1 To dataLength
        Key = Key & Mid(buffer, a + 4, 1)
      Next
    
    
      If vpbServerKeyWindows = "" Then
        
        vpbServerKeyWindows = Key
        
        buffer = "getKeyMySQL"
        dataLength = 11
        wsPassword.SendData dataLength
        
        ReDim byteArray(dataLength - 1)
        For a = 0 To dataLength - 1
          byteArray(a) = Asc(Mid(buffer, a + 1, 1))
        Next
        wsPassword.SendData byteArray
  
      Else
        
        vpbServerKeyMySQL = Key
      
      
        konekServer1
        
        x = GetComputerName(ComName, 255)
        cname = Trim(ComName)
        cname = Left(cname, Len(cname) - 1)
        txtRoom.text = cname
        vpbNamaKomputer = cname
        
        Sql = "SELECT room.IDROOM, roomprice.ROOMID FROM room INNER JOIN roomprice ON room.ROOMID = roomprice.ROOMID where room.ROOMNAME = '" & txtRoom.text & "';"
        Set myrs = MyConn.Execute(Sql)
        
        
        txtIdRoom.text = myrs.Fields(0).value
        txtRoomId.text = myrs.Fields(1).value
        
        Cinema.txtCompName.text = txtRoom.text
        Cinema.Show
        
        Timer1.Enabled = True
      
        wsPassword.Close
      End If
    End If
    
    wsPasswordBytesTotal = 0
    wsPasswordDataLength = 0
  End If
  
  If Err.Number <> 0 Then
    LogError Name, "wsPassword_DataArrival"
  End If
End Sub

Private Sub wsPassword_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  
  On Error Resume Next
  
  If Number = 10060 Then
    If wsPassword.RemoteHost = vpbServerUtama Then
      wsPassword.Connect vpbServerBackup
    Else
      wsPassword.Connect vpbServerUtama
    End If
  End If
  
  DoEvents
  
  If Err.Number <> 0 Then
    LogError Name, "wsPassword_Error"
  End If
End Sub


Private Sub setRCCOMCode(code As Long)
  
  On Error GoTo hell
  
  usbRCCom.PortOpen = True
  Sleep 2
  usbRCCom.Output = Chr(&H50)
  Sleep 2
  usbRCCom.Output = Chr(&H51)
  Sleep 2
  usbRCCom.Output = Chr(code)
  Sleep 2
  usbRCCom.PortOpen = False
  
  rcCOMCode = code
  
hell:

  If Err.Number <> 0 Then
    LogError Name, "setUSBRCCode"
  End If
  
End Sub

Public Sub turnDiscoLampOff()
  
  On Error GoTo hell
  
  Select Case settingRCType
    Case "COM"
      'setRCCOMCode rcCOMCode Or settingRCCOMCodeChannelDiscoLamp
      setRCCOMCode &HFF
    Case "HID"
      If rcHIDDI <> 0 And rcHIDHandle <> 0 Then
        'If usb_relay_device_close_one_relay_channel(rcHIDHandle, settingRCHIDChannelDiscoLamp) <> 0 Then
        If usb_relay_device_close_all_relay_channel(rcHIDHandle) <> 0 Then
          LogText "usb_relay_device_close_all_relay_channel fail"
        End If
      End If
    Case "LPT"
      Out settingRCLPTAddress, 1
  End Select
  
hell:

  If Err.Number <> 0 Then
    LogError Name, "turnDiscoLampOff"
  End If
  
End Sub
Public Sub turnDiscoLampOn()
  
  On Error GoTo hell
  
  Select Case settingRCType
    Case "COM"
      'setRCCOMCode rcCOMCode Xor settingRCCOMCodeChannelDiscoLamp
      setRCCOMCode &H0
    Case "HID"
      If rcHIDDI <> 0 And rcHIDHandle <> 0 Then
        'If usb_relay_device_open_one_relay_channel(rcHIDHandle, settingRCHIDChannelDiscoLamp) <> 0 Then
        If usb_relay_device_open_all_relay_channel(rcHIDHandle) <> 0 Then
          LogText "usb_relay_device_open_all_relay_channel fail"
        End If
      End If
    Case "LPT"
      Out settingRCLPTAddress, 0
  End Select
  
hell:

  If Err.Number <> 0 Then
    LogError Name, "turnDiscoLampOn"
  End If
End Sub


Friend Sub sendUSBAudioCommand(commandString As String)
  
  On Error GoTo hell
    
  Dim Data() As Byte
  Dim command As Byte
  Dim parameter1 As Byte
  Dim parameter2 As Byte

  Select Case usbAudioModel
    Case usbAudioModelK3200D
      Select Case UCase(commandString)
        Case "ECHODN"
          command = &H4
        Case "ECHOUP"
          command = &H3
        Case "MICDN"
          command = &H2
        Case "MICUP"
          command = &H1
        Case "PRESETDN"
          command = &H4
        Case "PRESETUP"
          command = &H3
        Case "PRESET0"
          command = &H22
        Case "PRESET1"
          command = &H21
        Case "PRESET2"
          command = &H22
        Case "PRESET3"
          command = &H23
        Case "PRESET4"
          command = &H24
        Case "PRESET5"
          command = &H25
        Case "PRESET6"
          command = &H26
        Case "PRESET7"
          command = &H27
        Case "PRESET8"
          command = &H28
      End Select
      ReDim Data(10)
      Data(0) = &H0
      Data(1) = &H44
      Data(2) = &H49
      Data(3) = &H47
      Data(4) = &H49
      Data(5) = &H31
      Data(6) = command
      Data(7) = &H47
      Data(8) = &H5A
      Data(9) = &H44
      Data(10) = &H4C
    Case usbAudioModelK3200P
      Select Case UCase(commandString)
        Case "ECHODN"
          command = &H4
        Case "ECHOUP"
          command = &H3
        Case "MICDN"
          command = &H2
        Case "MICUP"
          command = &H1
        Case "PRESETDN"
          command = &H4
        Case "PRESETUP"
          command = &H3
        Case "PRESET0"
          command = &H22
        Case "PRESET1"
          command = &H21
        Case "PRESET2"
          command = &H22
        Case "PRESET3"
          command = &H23
        Case "PRESET4"
          command = &H24
        Case "PRESET5"
          command = &H25
        Case "PRESET6"
          command = &H26
        Case "PRESET7"
          command = &H27
        Case "PRESET8"
          command = &H28
      End Select
      ReDim Data(10)
      Data(0) = &H0
      Data(1) = &H44
      Data(2) = &H49
      Data(3) = &H47
      Data(4) = &H49
      Data(5) = &H31
      Data(6) = command
      Data(7) = &H47
      Data(8) = &H5A
      Data(9) = &H44
      Data(10) = &H4C
    Case usbAudioModelK6000D
      Select Case UCase(commandString)
        Case "ECHODN"
          command = &H83
        Case "ECHOUP"
          command = &H82
        Case "MICDN"
          command = &H81
        Case "MICUP"
          command = &H80
        Case "PRESETDN"
          command = &H83
        Case "PRESETUP"
          command = &H82
        Case "PRESET0"
          command = &H2
        Case "PRESET1"
          command = &H1
        Case "PRESET2"
          command = &H2
        Case "PRESET3"
          command = &H3
        Case "PRESET4"
          command = &H4
        Case "PRESET5"
          command = &H5
        Case "PRESET6"
          command = &H6
        Case "PRESET7"
          command = &H7
        Case "PRESET8"
          command = &H8
      End Select
      ReDim Data(6)
      Data(0) = &H99
      Data(1) = &HAA
      Data(2) = &HBB
      Data(3) = command
      Data(4) = &HCC
      Data(5) = &HDD
      Data(6) = &HEE
    Case usbAudioModelK6800D
      Select Case UCase(commandString)
        Case "ECHODN"
          command = &H31
          parameter1 = &H2
          parameter2 = &H1
        Case "ECHOUP"
          command = &H31
          parameter1 = &H2
          parameter2 = &H0
        Case "MICDN"
          command = &H31
          parameter1 = &H0
          parameter2 = &H1
        Case "MICUP"
          command = &H31
          parameter1 = &H0
          parameter2 = &H0
        Case "PRESETDN"
          command = &H31
          parameter1 = &H2
          parameter2 = &H1
        Case "PRESETUP"
          command = &H31
          parameter1 = &H2
          parameter2 = &H0
        Case "PRESET0"
          command = &H33
          parameter1 = &H1
          parameter2 = &H1
        Case "PRESET1"
          command = &H33
          parameter1 = &H1
          parameter2 = &H0
        Case "PRESET2"
          command = &H33
          parameter1 = &H1
          parameter2 = &H1
        Case "PRESET3"
          command = &H33
          parameter1 = &H1
          parameter2 = &H2
        Case "PRESET4"
          command = &H33
          parameter1 = &H1
          parameter2 = &H3
        Case "PRESET5"
          command = &H33
          parameter1 = &H1
          parameter2 = &H4
        Case "PRESET6"
          command = &H33
          parameter1 = &H1
          parameter2 = &H5
        Case "PRESET7"
          command = &H33
          parameter1 = &H1
          parameter2 = &H6
        Case "PRESET8"
          command = &H33
          parameter1 = &H1
          parameter2 = &H7
      End Select
      ReDim Data(7)
      Data(0) = &H7B
      Data(1) = &H7D
      Data(2) = &H1
      Data(3) = command
      Data(4) = parameter1
      Data(5) = parameter2
      Data(6) = &H7D
      Data(7) = &H7B
  End Select
  
  usbAudioCom.PortOpen = True
  Sleep 1
  usbAudioCom.Output = Data
  Sleep 1
  usbAudioCom.PortOpen = False
  

hell:

  If Err.Number <> 0 Then
    LogError Name, "sendUSBAudioCommand"
  End If
  
End Sub
