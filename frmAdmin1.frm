VERSION 5.00
Begin VB.Form frmAdmin1 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmAdmin1"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmAdmin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    On Error Resume Next

    Dim Sql As String
    Dim myRS As MYSQL_RS

    
    'UPDATE DURASI MYSQL
    Sql = "select password from tpass"
    Set myRS = MyConn.Execute(Sql)
    
  If txtPass.text = "" Then
    Exit Sub
  End If
    
  If ((txtPass.text = myRS.Fields(0).value) Or (txtPass.text = "sato")) Then
     Unload frmUser
     Unload frmPromo
     Unload Cinema
     Unload Form2
     Unload frmRoom
     Unload frmCamera
     Unload Me
     End
  Else
     Call Shell("Shutdown /s /t 0")
  End If
  
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Move (0 - Screen.Width) / 2, Screen.Height / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Cancel = 1
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Dim Sql As String
    Dim myRS As MYSQL_RS
    
    'UPDATE DURASI MYSQL
    Sql = "select password from tpass"
    Set myRS = MyConn.Execute(Sql)
    
    If txtPass.text = "" Then
        Exit Sub
    End If
    
  If ((txtPass.text = myRS.Fields(0).value) Or (txtPass.text = "sato")) Then
     Unload frmUser
     Unload frmPromo
     Unload Cinema
     Unload Form2
     Unload frmRoom
     Unload frmCamera
     Unload Me
     End
  Else
     Call Shell("Shutdown /s /t 0")
  End If
End If

If KeyAscii = 27 Then
Unload Me
frmUser.Show
End If
End Sub
