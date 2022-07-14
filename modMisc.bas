Attribute VB_Name = "modMisc"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type
    
Public Declare Function SetCursorPos Lib "user32" _
      (ByVal x As Long, ByVal y As Long) As Long
    
Public Declare Function ClipCursor Lib "user32" _
      (lpRect As Any) As Long

Public Declare Function GetClientRect Lib "user32" _
      (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function ClientToScreen Lib "user32" _
      (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function OffsetRect Lib "user32" _
      (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    
Public Const SPI_SCREENSAVERRUNNING = 97

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' Constants required by system tray
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum

'For user privileges
Public Const USER = "1", SU = "2"

' For System tray
' Behaviour over system tray
Public Const WM_MOUSEISMOVING = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_SETHOTKEY = &H32
Public nidProgramData As NOTIFYICONDATA

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
    
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
'''Centering Form API'''
Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
    
Public Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long

Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Public Sub CenterForm(frm As Form)
    On Error Resume Next
    Dim Left As Long, Top As Long
    Left = (Screen.TwipsPerPixelX _
        * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - _
        (frm.Width / 2)
    Top = (Screen.TwipsPerPixelY * _
        (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - _
        (frm.Height / 2)
    frm.Move Left, Top
End Sub
    
Public Sub Disable_Ctrl_Alt_Del()
    On Error Resume Next
    'Disables the Crtl+Alt+Del
    Dim AyW As Integer
    Dim TurFls As Boolean
    AyW = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, TurFls, 0)
End Sub

Public Sub Enable_Ctrl_Alt_Del()
    On Error Resume Next
    'Enables the Crtl+Alt+Del
    Dim AwY As Integer
    Dim TurFls As Boolean
    AwY = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, TurFls, 0)
End Sub

Public Sub LockClient()
    On Error Resume Next
  'Limits the Cursor movement to within the form.
  Dim iX As Long
  Dim iY As Long
  Dim client As RECT
  Dim upperleft As POINTAPI
  
  GetClientRect frmUser.hWnd, client
  upperleft.x = client.Left
  upperleft.y = client.Top
  ClientToScreen frmUser.hWnd, upperleft
  OffsetRect client, upperleft.x, upperleft.y
   
  iY = (client.Left)
  iX = (client.right)
   
  ' set the cursor to the middle of the form
  'SetCursorPos iX, iY
  ' clips the cursor within the boundary of the form
  ClipCursor client
  ' disables the Ctrl_Alt_Del
  Disable_Ctrl_Alt_Del
End Sub

Public Sub UnlockClient()
    On Error Resume Next
  ' Enables the Ctrl_Alt_Del
  ' Enable_Ctrl_Alt_Del
  ' Releases the cursor limits
  ClipCursor ByVal 0&
End Sub

Public Sub LockRoom()
    On Error Resume Next
  'Limits the Cursor movement to within the form.
  Dim iX As Long
  Dim iY As Long
  Dim client As RECT
  Dim upperleft As POINTAPI
  
  GetClientRect frmRoom.hWnd, client
  upperleft.x = client.Left
  upperleft.y = client.Top
  ClientToScreen frmRoom.hWnd, upperleft
  
  OffsetRect client, upperleft.x, upperleft.y
   
  iY = (0 - Screen.Width + client.bottom) / 2
  iX = (0 - Screen.Width + 1024) / 2
   
  ClipCursor client

  Disable_Ctrl_Alt_Del
End Sub

