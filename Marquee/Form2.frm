VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H000090FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   FillColor       =   &H00000080&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form2"
   ScaleHeight     =   1065
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      FillColor       =   &H009D2206&
      ForeColor       =   &H0034FFFD&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   6000
         TabIndex        =   1
         Top             =   195
         Width           =   150
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Transparent Form
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_ALPHA = &H2&
'-----------------------------------

'Posisi
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
'----------------------------------

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Width = Screen.Width
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    '   SWP_NOMOVE + SWP_NOSIZE
    On Error Resume Next
    Dim bytOpacity As Byte       'Set the transparency level
    bytOpacity = 128
    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
End Sub
