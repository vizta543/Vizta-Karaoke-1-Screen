VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmDice2 
   BorderStyle     =   0  'None
   Caption         =   "frmDice2"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmDice2.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmDice2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private timerCount As Long

Private Sub Form_Load()
  
  On Error Resume Next

  Skin1.LoadSkin App.Path & "\skin\dadu.skn"
  Skin1.ApplySkinByName hWnd, "dadu"
  
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

  Move (0 - Screen.Width)
  
  timerCount = 0
  
  tmr.Enabled = True

End Sub

Private Sub tmr_Timer()

  On Error Resume Next
  
  timerCount = timerCount + 1
  
  If timerCount = 5 Then
  
    tmr.Enabled = False
    Unload Me
    Unload frmConfirmasi
  End If
End Sub
