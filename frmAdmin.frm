VERSION 5.00
Object = "{59E060F0-4FD5-4C80-AF3C-D1B7E0ED65B2}#1.0#0"; "CompControls.ocx"
Begin VB.Form frmAdmin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CompControler.CompControl CompControl1 
      Left            =   0
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FF8080&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtAdmin 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Admin"
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (0 - Screen.Width) / 2, Screen.Height / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
CompControl1.ALT_CTRL_DEL_Enabled
CompControl1.DesktopIconsShow
CompControl1.TaskBarShow
wlRestoreAll
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If txtPass.text = "kgd" Then
     Unload frmPromo
     Unload Cinema
     Unload Form2
     Unload frmRoom
     Unload frmCamera
     Unload Me
  Else
    MsgBox "Password Anda Salah ", vbOKOnly, "Warning"
    txtPass.text = ""
    txtPass.SetFocus
  End If
End If

If KeyAscii = 27 Then
   Unload Me
   frmUser.Show
End If
End Sub
