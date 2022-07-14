Attribute VB_Name = "Module1"
Option Explicit

' virtual key codes
' For more info see win32api.txt file supplied with Visual Basic
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F
Public Const VK_F1 = &H70
Public Const VK_F11 = &H7A

Private GDISABLE As Long

Sub AddInfo(info As String)
   Formdwlcore.List1.AddItem (info)
   Formdwlcore.List1.ListIndex = Formdwlcore.List1.ListCount - 1
End Sub


Sub DisableKeys(disable As Long)
  Dim TBar As Long
  Dim ret
  
  ' disables system trays
  ret = wlDisableItem(wlTaskTray, disable)
  ' disable Alt+Tab
  ret = wlDisableKey(0, VK_TAB, MOD_ALT, disable)
  'disable Ctrl+Esc
  ret = wlDisableKey(0, VK_ESCAPE, MOD_CONTROL, disable)
  ' MOD_ALL = also with Ctrl,Alt,Shift,Win Keys
  ret = wlDisableKey(0, VK_F1, MOD_ALL, disable)
  ' disable only F11
  ret = wlDisableKey(0, VK_F11, 0, disable)
  ' disables all WIN keys
  ret = wlDisableKey(0, 0, MOD_WIN, disable)
  ' only if dwlgina is installed
  If wlIsGinaInstalled = 1 Then
    ret = wlDisableKey(0, VK_DELETE, MOD_CONTROL Or MOD_ALT, disable)
  End If

  ' disables right mouse button only on taskbar window
  ' MOD_ALL = disables Key also with Ctrl, Alt, Shift and Win Keys
  TBar = FindWindowEx(0, 0, "Shell_TrayWnd", "")
  If TBar > 0 Then
    ret = wlDisableKey(TBar, VK_RBUTTON, MOD_ALL, disable)
  End If
End Sub


Sub MyKeyEvent(ByVal UserData As Long, ByVal wnd As Long, ByVal down As Long, ByVal vk As Long, ByVal mf As Long)
  Dim ret As Long
  
    
  If wlIsKeyDisabled(0, vk, mf) = 1 Then
    AddInfo ("Sorry, but this key is disabled")
  End If
 
  ' detect key Ctrl+Alt+Win+Shift+A (only an example)
  If (vk = Asc("A")) And (mf = MOD_WIN Or MOD_CONTROL Or MOD_ALT Or MOD_SHIFT) Then
    AddInfo ("'hand break' key pressed :-)")
  End If

 ' Ctrl+Shift+T to enable and disable keys
  If (vk = Asc("T")) And (mf = MOD_CONTROL Or MOD_SHIFT) Then
    If GDISABLE = 1 Then
      GDISABLE = 0
      AddInfo ("Keys enabled")
    Else
      GDISABLE = 1
      AddInfo ("Keys disabled")
    End If
    DisableKeys (GDISABLE)
  End If
End Sub


Sub dwlInit()
  Dim ret As Long
  
  AddInfo ("dWinlock core example")
  AddInfo ("(c) 2003 Kassl GmbH (http://www.kassl.de)")
  AddInfo ("")
  AddInfo ("Following keys are system wide disabled:")
  AddInfo ("Alt+Tab")
  AddInfo ("Ctrl+Esc")
  AddInfo ("Win keys")
  AddInfo ("F1 with all modifiers (Ctrl,Alt,Shift,Win)")
  AddInfo ("F11 without modifiers")
  AddInfo ("Right mouse button on task bar")
  AddInfo ("and also Ctrl+Alt+Del if dwgina is installed")
  AddInfo ("")
  AddInfo ("Press Shift+Ctrl+T to toggle (enable/disable)")
  AddInfo ("")
  ret = wlSetKeyCallback(AddressOf MyKeyEvent, 0, kDownEvents)
  DisableKeys (1)
  GDISABLE = 1
End Sub

Sub dwlExit()
  wlExit
End Sub


