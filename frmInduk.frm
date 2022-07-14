VERSION 5.00
Begin VB.MDIForm frmInduk 
   Appearance      =   0  'Flat
   BackColor       =   &H000090FF&
   Caption         =   "MDIForm1"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Enabled         =   0   'False
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmInduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
'frmPromo.Show
'LockRoom
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Unload ActiveControl
Unload frmPromo
'Unload Me
'Unload ActiveForm
MyConn.CloseConnection
Set MyConn = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Unload ActiveForm
End Sub
