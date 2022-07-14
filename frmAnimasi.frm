VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnimasi 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   6045
   ClientTop       =   4935
   ClientWidth     =   2910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   360
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393216
      BackColor       =   14215660
      FullWidth       =   41
      FullHeight      =   41
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sedang menghubungi server..."
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   750
      Width           =   2775
   End
End
Attribute VB_Name = "frmAnimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  Dim m_selOutputPlayerA As Integer
    
Private Sub Form_Load()
  'DoEvents
  'frmUtama.StatusBar1.Panels(1).text = "Sedang menghubungi server " & frmSettingDatabase.Server.text & "..."
  DoEvents
  If Dir(App.Path & "\Server.avi") <> "" Then
    Animation1.Visible = True
    DoEvents
    Animation1.Open (App.Path & "\Server.avi")
    DoEvents
    Animation1.play
    DoEvents
  End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

MyConn.SetOption MYSQL_OPT_COMPRESS
frmRoom.Show

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
Private Sub SoundCard()
    Dim nOutputs As Integer
    nOutputs = frmVideo.Amp3dj1.Get()
    Dim count As Integer
    For count = 0 To nOutputs - 1
        ComboPlayer.AddItem frmVideo.Amp3dj1.GetOutputDeviceDesc(count)
    Next count
    ComboPlayer.ListIndex = 0
End Sub
Private Sub BtnOK()
    'm_selOutputPlayerA = ComboPlayer.ListIndex
    'm_selOutputPlayerB = ComboPlayerB.ListIndex
    frmVideo.setOutput ComboPlayer.ListIndex
    'Unload Me
End Sub

