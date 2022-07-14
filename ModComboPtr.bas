Attribute VB_Name = "ModComboPtr"
Option Explicit

Private Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, ByVal x As Long, ByVal Y As _
  Long, ByVal nWidth As Long, ByVal nHeight As Long, _
  ByVal bRepaint As Long) As Long
  
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" ( _
                 ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "User32.dll" (ByVal hCursor As Long) As Long

Private Declare Function apiCreateCaret Lib "user32" _
        Alias "CreateCaret" _
        (ByVal hWnd As Long, _
        ByVal hBitmap As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) _
        As Long

Private Declare Function apiShowCaret Lib "user32" _
        Alias "ShowCaret" _
        (ByVal hWnd As Long) _
        As Long
 Private Declare Function apiGetFocus Lib "user32" _
        Alias "GetFocus" _
         () As Long

Sub sMakeCaret(ctl As Control, _
                        intX As Integer, _
                        intY As Integer)
    On Error Resume Next
    Dim hWnd As Long
    hWnd = fhWnd(ctl)
    Call apiCreateCaret(hWnd, 0&, intX, intY)
    Call apiShowCaret(hWnd)
End Sub
     
Function fhWnd(ctl As Control) As Long
    On Error Resume Next
    ctl.SetFocus
    If Err Then
        fhWnd = 0
    Else
        fhWnd = apiGetFocus
    End If
    On Error GoTo 0
End Function

Public Function SetComboBoxHeight(objCB As ComboBox, _
  TheHeight As Single) As Boolean
End Function

' Resize a ComboBox's dropdown display area.
Public Sub SizeCombo(frm As Form, cbo As ComboBox)
    On Error Resume Next
Dim cbo_left As Integer
Dim cbo_top As Integer
Dim cbo_width As Integer
Dim cbo_height As Integer
Dim old_scale_mode As Integer

    ' Change the Scale Mode on the form to Pixels.
    old_scale_mode = frm.ScaleMode
    frm.ScaleMode = vbPixels

    ' Save the ComboBox's Left, Top, and Width values.
    cbo_left = cbo.Left
    cbo_top = cbo.Top
    cbo_width = cbo.Width

    ' Calculate the new height of the combo box.
    cbo_height = frm.ScaleHeight - cbo.Top - 5
    frm.ScaleMode = old_scale_mode

    ' Resize the combo box window.
    MoveWindow cbo.hWnd, cbo_left, cbo_top, _
        cbo_width, cbo_height, 1
End Sub
