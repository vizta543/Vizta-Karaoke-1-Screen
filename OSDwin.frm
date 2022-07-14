VERSION 5.00
Begin VB.Form OSDwin 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   690
   ClientLeft      =   -195
   ClientTop       =   -195
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer ScrollTimer 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   480
      Top             =   0
   End
   Begin VB.Label text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   150
   End
End
Attribute VB_Name = "OSDwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AnimateOSD As Boolean

Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    

Private Sub Form_Load()
Me.Top = 300
Me.Left = Screen.Width
Width = Screen.Width
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
   SWP_NOMOVE + SWP_NOSIZE
If AnimateOSD Then
Timer.Interval = 1500
Else
Timer.Interval = 4000
End If
Timer.Enabled = True
End Sub

Private Sub ScrollTimer_Timer()
If text.Left > -text.Width Then
text.Left = text.Left - 125
Else
text.Left = Screen.Width
End If
End Sub

Private Sub text_Click()
Unload Me
End Sub

Private Sub Timer_Timer()
Timer.Enabled = False
If AnimateOSD Then ScrollTimer.Enabled = True Else Unload Me
End Sub
