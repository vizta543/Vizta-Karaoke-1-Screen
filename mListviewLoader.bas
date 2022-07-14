Attribute VB_Name = "mListviewLoader"

Option Explicit
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, lParam As Any) As Long
 
Private Const LVM_SETITEMCOUNT As Long = 4096 + 47
 
