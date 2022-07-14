VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   1050
   ClientTop       =   1215
   ClientWidth     =   12000
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5280
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7320
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   3840
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Picture1 
      Height          =   7815
      Left            =   240
      ToolTipText     =   "Press Enter / Esc  to close image"
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer2.Interval = 10000
End Sub

Private Sub Timer2_Timer()
Dim Gambar As String
Dim rasio, rasio1, rasio2 As Double
Dim i As Integer
Dim SQL As String
Dim MyRs As MYSQL_RS
SQL = "SELECT LOGOPATH from logo"
Set MyRs = MyConn.Execute(SQL)

For i = 1 To MyRs.RecordCount
Timer1.Enabled = True
Gambar = "X:" & "\" & MyRs.Fields(0).Value
Form2.Image1.Picture = LoadPicture(Gambar)
Form2.Picture1.Picture = LoadPicture(Gambar)
Form2.BorderStyle = 0
Form2.WindowState = 2

If Form2.Picture1.Stretch = False Then
   Form2.Picture1.Stretch = True
Else: Form2.Picture1.Stretch = False
End If
   
If Form2.Picture1.Visible = True Then
Form2.Picture1.Visible = False
Else: Form2.Picture1.Visible = False
End If

Form2.Show

If Form2.Image1.Height > Form2.Height And Form2.Image1.Width <= Form2.Width Then
   rasio = Form2.Picture1.Height - Form2.Height
   rasio1 = rasio / Form2.Picture1.Height
   rasio2 = rasio1 * Form2.Picture1.Width
   Form2.Picture1.Width = Form2.Picture1.Width - rasio2
   Form2.Picture1.Height = Form2.Picture1.Height - (Form2.Picture1.Height - Form2.Height)
   Form2.Picture1.Stretch = True
   Form2.Picture1.Top = ((Form2.Height - Form2.Picture1.Height) / 2)
   Form2.Picture1.Left = ((Form2.Width - Form2.Picture1.Width) / 2)
   Form2.Show
   Form2.Picture1.Visible = True
  
End If

If Form2.Image1.Width > Form2.Width And Form2.Image1.Height <= Form2.Height Then
   rasio = Form2.Picture1.Width - Form2.Width
   rasio1 = rasio / Form2.Picture1.Width
   rasio2 = rasio1 * Form2.Picture1.Height
   Form2.Picture1.Width = Form2.Picture1.Width - (Form2.Picture1.Width - Form2.Width)
   Form2.Picture1.Height = Form2.Picture1.Height - rasio2
   Form2.Picture1.Stretch = True
   Form2.Picture1.Top = ((Form2.Height - Form2.Picture1.Height) / 2)
   Form2.Picture1.Left = ((Form2.Width - Form2.Picture1.Width) / 2)
   Form2.Show
   Form2.Picture1.Visible = True
  
End If

If Form2.Image1.Width > Form2.Width And Form2.Image1.Height > Form2.Height Then
 If Form2.Image1.Width <= Form2.Image1.Height Then
   rasio = Form2.Picture1.Height - Form2.Height
   rasio1 = rasio / Form2.Picture1.Height
   rasio2 = rasio1 * Form2.Picture1.Width
   Form2.Picture1.Width = Form2.Picture1.Width - rasio2
   Form2.Picture1.Height = Form2.Picture1.Height - (Form2.Picture1.Height - Form2.Height)
   Form2.Picture1.Top = ((Form2.Height - Form2.Picture1.Height) / 2)
   Form2.Picture1.Left = ((Form2.Width - Form2.Picture1.Width) / 2)
   Form2.Picture1.Stretch = True
   Form2.Show
   Form2.Picture1.Visible = True
   
 End If
 If Form2.Image1.Width > Form2.Image1.Height Then
   Form2.Picture1.Width = Form2.Picture1.Width - (Form2.Picture1.Width - Form2.Width)
   Form2.Picture1.Height = Form2.Picture1.Height - (Form2.Picture1.Height - Form2.Height)
   Form2.Picture1.Top = ((Form2.Height - Form2.Picture1.Height) / 2)
   Form2.Picture1.Left = ((Form2.Width - Form2.Picture1.Width) / 2)
   Form2.Picture1.Stretch = True
   Form2.Show
   Form2.Picture1.Visible = True
   
 End If
End If

If Form2.Image1.Width <= Form2.Width And Form2.Image1.Height <= Form2.Height Then
   Form2.Picture1.Width = Form2.Picture1.Width
   Form2.Picture1.Height = Form2.Picture1.Height
   Form2.Picture1.Stretch = False
   Form2.Picture1.Top = ((Form2.Height - Form2.Picture1.Height) / 2)
   Form2.Picture1.Left = ((Form2.Width - Form2.Picture1.Width) / 2)
   Form2.Show
   Form2.Picture1.Visible = True
End If
Timer2.Interval = 10000
Form2.Refresh
MyRs.MoveNext
Next i

End Sub
