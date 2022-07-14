VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#7.12#0"; "3dabm7u.ocx"
Begin VB.Form FormSoundCard 
   BackColor       =   &H00523939&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sound card selection"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboPlayer 
      Height          =   315
      Left            =   308
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3645
   End
   Begin BTNENHLib4.BtnEnh BtnOK 
      Height          =   540
      Left            =   1380
      TabIndex        =   0
      Top             =   1680
      Width           =   1500
      _Version        =   458764
      _ExtentX        =   2646
      _ExtentY        =   952
      _StockProps     =   66
      Caption         =   "OK"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      Surface         =   9
      BackColorContainer=   5388601
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      UserData        =   0.1
      textCaption     =   "FormSoundCard.frx":0000
      textLT          =   "FormSoundCard.frx":0064
      textCT          =   "FormSoundCard.frx":007C
      textRT          =   "FormSoundCard.frx":0094
      textLM          =   "FormSoundCard.frx":00AC
      textRM          =   "FormSoundCard.frx":00C4
      textLB          =   "FormSoundCard.frx":00DC
      textCB          =   "FormSoundCard.frx":00F4
      textRB          =   "FormSoundCard.frx":010C
      colorBack       =   "FormSoundCard.frx":0124
      colorIntern     =   "FormSoundCard.frx":014A
      colorMO         =   "FormSoundCard.frx":0170
      colorFocus      =   "FormSoundCard.frx":0196
      colorDisabled   =   "FormSoundCard.frx":01BC
      colorPressed    =   "FormSoundCard.frx":01E2
   End
   Begin VB.Label Label1 
      BackColor       =   &H00523939&
      Caption         =   "Select the output sound card"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   308
      TabIndex        =   2
      Top             =   480
      Width           =   2580
   End
End
Attribute VB_Name = "FormSoundCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_selOutputPlayerA As Integer
Dim m_selOutputPlayerB As Integer
Private Sub BtnOK_Click()
    m_selOutputPlayerA = ComboPlayerA.ListIndex
    m_selOutputPlayerB = ComboPlayerB.ListIndex
    frmMain.setOutputs m_selOutputPlayerA, m_selOutputPlayerB
    Unload Me
End Sub

Private Sub Form_Load()
    Dim nOutputs As Integer
    nOutputs = frmMain.Amp3dj1.GetOutputDevicesCount()
    Dim count As Integer
    For count = 0 To nOutputs - 1
        ComboPlayerA.AddItem frmMain.Amp3dj1.GetOutputDeviceDesc(count)
        ComboPlayerB.AddItem frmMain.Amp3dj1.GetOutputDeviceDesc(count)
    Next count
    
    ComboPlayerA.ListIndex = m_selOutputPlayerA
    ComboPlayerB.ListIndex = m_selOutputPlayerB
    
    Check3Dabm7
End Sub

