VERSION 5.00
Begin VB.Form Cinema 
   BorderStyle     =   0  'None
   Caption         =   "Cinema"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtCompName 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   735
      Left            =   720
      ScaleHeight     =   675
      ScaleWidth      =   6675
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "VISUAL BASIC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "&Selesai"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataSource      =   "data"
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6240
      Top             =   2280
   End
   Begin VB.Label lblRecordAktif 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Record ="
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblKota 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
End
Attribute VB_Name = "Cinema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Connection
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim Kota As String
Dim Aktif As Integer

Private Sub Timer1_Timer()
  
  On Error Resume Next
  
  Form2.Picture5.Refresh
    
    If Form2.Label17.Left < -Form2.Label17.Width Then
      Form2.Label17.Left = Form2.Picture5.Width
      If katajalanx >= katajalantotal Then
        LoadKatajalan
      Else
        katajalanx = katajalanx + 1
      End If
      Form2.Label17 = katajalan(katajalanx - 1)
    End If
    
    Form2.Label17.Left = Form2.Label17.Left - 55
            
            
    If Err.Number <> 0 Then
      LogError Me.Name, "Timer1_Timer"
    End If

End Sub

Private Sub cmdSelesai_Click()
  On Error Resume Next
  End
End Sub

Public Sub Form_Load()

    On Error Resume Next
    
    Form2.Show
    
    Me.Width = 0
    Me.Height = 0
  
    LoadKatajalan
    
    Form2.Label17.Left = Form2.Picture5.Width
         
        
    If Err.Number <> 0 Then
      LogError Me.Name, "Form_Load"
    End If

End Sub

Sub LoadKatajalan()

    On Error Resume Next
    
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim i As Integer
    
    Cinema.Timer1.Enabled = False
    
    Sql = "select idtulisan, tulisan from osdtext where ROOM = '0' or ROOM = '" & txtCompName & "' order by ROOM desc;"
    Set myrs = MyConn.Execute(Sql)
    katajalantotal = myrs.recordCount
    ReDim katajalan(katajalantotal)
    myrs.MoveFirst
    i = 0
    Do Until myrs.EOF
        katajalan(i) = myrs.Fields(1).value
        myrs.MoveNext
        i = i + 1
    Loop
    
    Sql = "UPDATE room SET katajalan = 0 WHERE ROOMNAME = '" & txtCompName & "';"
    MyConn.Execute Sql
    
    katajalanx = 1
    Form2.Label17.Caption = katajalan(0)
    
    Cinema.Timer1.Enabled = True
        
        
    If Err.Number <> 0 Then
      LogError Me.Name, "LoadKatajalan"
    End If

End Sub

