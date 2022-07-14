VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{48B1464F-F357-11D7-B2E7-00001C56B9BE}#1.0#0"; "CoolButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{35FD49A8-724D-430C-A6E4-7FE74EE95312}#4.21#0"; "UniBox210.ocx"
Begin VB.Form frmRoom 
   BackColor       =   &H0000244C&
   BorderStyle     =   0  'None
   Caption         =   "frmRoom"
   ClientHeight    =   11520
   ClientLeft      =   -345
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrNonVocalML1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   4200
   End
   Begin VB.Timer tmrNonVocalMR1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1800
      Top             =   4200
   End
   Begin VB.Timer tmrVokal1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   4200
   End
   Begin VB.Timer tmrVokal2 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   4560
   End
   Begin VB.Timer tmrNonVocalML2 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   4560
   End
   Begin VB.Timer tmrNonVocalMR2 
      Enabled         =   0   'False
      Left            =   2160
      Top             =   4560
   End
   Begin VB.TextBox txtRemoteCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00022A6B&
      Height          =   285
      Left            =   13410
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Text            =   "123456"
      Top             =   720
      Width           =   1605
   End
   Begin VB.TextBox txtPlaying 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000244C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   11160
      Width           =   6975
   End
   Begin MSComctlLib.ListView lstTV 
      Height          =   1335
      Left            =   10080
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   9292
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtCategory 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000244C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   6240
      TabIndex        =   60
      Text            =   "Nations"
      Top             =   1260
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstMovie 
      Height          =   1335
      Left            =   8040
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   9292
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox picKeyTempo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1100
      Left            =   3960
      Picture         =   "frmRoom.frx":000C
      ScaleHeight     =   1095
      ScaleWidth      =   9555
      TabIndex        =   65
      Top             =   8040
      Visible         =   0   'False
      Width           =   9555
      Begin VB.Timer tmrPicKeyTempo 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   5040
         Top             =   120
      End
      Begin VB.Image KotakTambah 
         Height          =   390
         Index           =   0
         Left            =   5880
         Picture         =   "frmRoom.frx":D40C
         Top             =   480
         Width           =   465
      End
      Begin VB.Image KotakTambah 
         Height          =   390
         Index           =   1
         Left            =   6360
         Picture         =   "frmRoom.frx":FBC1
         Top             =   480
         Width           =   465
      End
      Begin VB.Image KotakTambah 
         Height          =   390
         Index           =   2
         Left            =   6795
         Picture         =   "frmRoom.frx":105C3
         Top             =   480
         Width           =   465
      End
      Begin VB.Image KotakTambah 
         Height          =   390
         Index           =   3
         Left            =   7200
         Picture         =   "frmRoom.frx":10FC5
         Top             =   480
         Width           =   465
      End
      Begin VB.Image KotakTambah 
         Height          =   390
         Index           =   4
         Left            =   7635
         Picture         =   "frmRoom.frx":119C7
         Top             =   480
         Width           =   465
      End
      Begin VB.Image KotakKurang 
         Height          =   405
         Index           =   0
         Left            =   4380
         Picture         =   "frmRoom.frx":1417C
         Top             =   480
         Width           =   480
      End
      Begin VB.Image KotakKurang 
         Height          =   405
         Index           =   1
         Left            =   3945
         Picture         =   "frmRoom.frx":16957
         Top             =   480
         Width           =   480
      End
      Begin VB.Image KotakKurang 
         Height          =   405
         Index           =   2
         Left            =   3510
         Picture         =   "frmRoom.frx":173B9
         Top             =   480
         Width           =   480
      End
      Begin VB.Image KotakKurang 
         Height          =   405
         Index           =   3
         Left            =   3090
         Picture         =   "frmRoom.frx":17E1B
         Top             =   480
         Width           =   480
      End
      Begin VB.Image KotakKurang 
         Height          =   405
         Index           =   4
         Left            =   2670
         Picture         =   "frmRoom.frx":1887D
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.TextBox txtVol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00955001&
      Height          =   645
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "100"
      Top             =   10200
      Width           =   1020
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006F1628&
      Height          =   720
      Left            =   4560
      TabIndex        =   2
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00580148&
      Height          =   270
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00838383&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00022993&
      Height          =   270
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   480
      Width           =   1575
   End
   Begin VB.Timer tmrRemote 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1080
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock WskRemote 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsTitle 
      Height          =   810
      Left            =   10800
      TabIndex        =   59
      Top             =   6480
      Width           =   1095
      _cx             =   1931
      _cy             =   1429
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Timer tmrMainVocal 
      Interval        =   1200
      Left            =   1560
      Top             =   1320
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsChatNewMessage 
      Height          =   765
      Left            =   60
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   1350
      _cx             =   2381
      _cy             =   1349
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsMovieCategory 
      Height          =   780
      Left            =   12000
      TabIndex        =   61
      Top             =   6480
      Visible         =   0   'False
      Width           =   2460
      _cx             =   4339
      _cy             =   1376
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin UniToolbox2.UniListView lstAll 
      Height          =   855
      Left            =   3720
      TabIndex        =   62
      Top             =   6120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1508
      Appearance      =   0
      BackColor       =   9292
      BorderStyle     =   0
      FlatScrollBar   =   -1  'True
      ForeColor       =   16777215
      FullRowSelect   =   -1  'True
      HideColumnHeaders=   -1  'True
      LabelEdit       =   1
      View            =   3
      PictureAlignment=   5
      Icons           =   "<None>"
      SmallIcons      =   "<None>"
      ColumnHeaderIcons=   "<None>"
      ColumnHeaders   =   "frmRoom.frx":192DF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   14.25
      FontWeight      =   700
      FontBold        =   -1  'True
      Settings        =   "UT2100SLqMMXcf88YWHV"
   End
   Begin VB.Timer tmrVolume 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1080
      Top             =   1800
   End
   Begin VB.TextBox txtSinopsis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H002B2B2B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   12000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   58
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtArtis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H002B2B2B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   12720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   57
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtChatAktif 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00580148&
      Height          =   345
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox lblRecording 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12960
      Picture         =   "frmRoom.frx":19337
      ScaleHeight     =   585
      ScaleWidth      =   2280
      TabIndex        =   41
      Top             =   2280
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Timer tmrNonVocalMR 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   2280
   End
   Begin VB.Timer tmrNonVocalML 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   2280
   End
   Begin VB.Timer tmrVokal 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2280
   End
   Begin VB.Timer tmrTemp 
      Interval        =   1000
      Left            =   1560
      Top             =   2760
   End
   Begin MSWinsockLib.Winsock chat 
      Index           =   0
      Left            =   600
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   120
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstChat 
      Height          =   1335
      Left            =   9480
      TabIndex        =   53
      Top             =   5880
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   7280168
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   52
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrAktif 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1320
   End
   Begin VB.PictureBox picbAbout 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   50
      Top             =   0
      Width           =   15
      Begin VB.PictureBox PicAbout 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   2415
         TabIndex        =   51
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.PictureBox picKategori 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6360
      Picture         =   "frmRoom.frx":1C98D
      ScaleHeight     =   405
      ScaleWidth      =   3255
      TabIndex        =   43
      Top             =   3960
      Width           =   3255
   End
   Begin VB.PictureBox PictureLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   10800
      ScaleHeight     =   435
      ScaleWidth      =   1095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
      Begin CoolButton.Button cmdSavePlayList 
         Height          =   375
         Left            =   840
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":2163C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   560
         Width           =   1935
      End
      Begin VB.TextBox txtLogin 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   420
         Left            =   2040
         TabIndex        =   21
         Top             =   100
         Width           =   1935
      End
      Begin CoolButton.Button cmdCancelLogin 
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":21658
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdLogin 
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":21674
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLogin 
         BackStyle       =   0  'Transparent
         Caption         =   "Member ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lstPlayUser 
      Height          =   375
      Left            =   7080
      TabIndex        =   44
      Top             =   6840
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.PictureBox btnRecStop 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12960
      Picture         =   "frmRoom.frx":21690
      ScaleHeight     =   585
      ScaleWidth      =   2400
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   1800
   End
   Begin VB.ComboBox cbodevice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboaudiodevice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbovideostand 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1560
      Top             =   2280
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtStartTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtRoomId 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtIdRoom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtNoOrderRoom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCompName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "frmRoom.frx":26105
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   600
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   1560
      Top             =   1800
   End
   Begin MSComctlLib.ListView lstPlaylist 
      Height          =   1095
      Left            =   5400
      TabIndex        =   0
      Top             =   6000
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   9292
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000244C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      ScaleHeight     =   375
      ScaleWidth      =   7215
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   980
      Width           =   7215
      Begin VB.TextBox txtNextSong 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000244C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "No Song"
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00580148&
      Height          =   525
      Left            =   5400
      MaxLength       =   22
      TabIndex        =   54
      Top             =   3240
      Visible         =   0   'False
      Width           =   5040
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsLogo 
      Height          =   1005
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   1515
      _cx             =   2672
      _cy             =   1773
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.ComboBox cbokategori 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00022993&
      Height          =   360
      ItemData        =   "frmRoom.frx":26339
      Left            =   480
      List            =   "frmRoom.frx":26340
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin MSComctlLib.ListView lstMidi 
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstMidiMusic 
      Height          =   375
      Left            =   7080
      TabIndex        =   45
      Top             =   5880
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   20000
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   11040
      Width           =   1095
      Begin CoolButton.Button cmdPlay 
         Height          =   735
         Left            =   360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26349
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdStop 
         Height          =   735
         Left            =   1320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26365
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdRequest 
         Height          =   735
         Left            =   2280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26381
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdPause 
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":2639D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdRepeat 
         Height          =   735
         Left            =   2280
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":263B9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdFast 
         Height          =   735
         Left            =   3240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":263D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdSlow 
         Height          =   735
         Left            =   3240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":263F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdPause 
         Height          =   735
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":2640D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   8640
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3840
      Width           =   15
      Begin CoolButton.Button cmdDelete 
         Height          =   735
         Left            =   3480
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26429
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdClear 
         Height          =   735
         Left            =   1920
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26445
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CoolButton.Button cmdLogout 
         Height          =   735
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRoom.frx":26461
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin CoolButton.Button cmdSong 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRoom.frx":2647D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmrLstAllLoad 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   600
      Top             =   1800
   End
   Begin MSWinsockLib.Winsock wsClientRemote 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   65533
   End
   Begin MSWinsockLib.Winsock wsServerRemote 
      Left            =   600
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   65534
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New Scripting.FileSystemObject

'TRANSPARAN FORM
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hWnd As Long, _
                    ByVal nIndex As Long) As Long

    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hWnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long

    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                    ByVal hWnd As Long, _
                    ByVal crKey As Long, _
                    ByVal bAlpha As Byte, _
                    ByVal dwFlags As Long) As Long

    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
'-----------------------------

'DISABLE SCROLLBAR
    Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, _
        ByVal wBar As Long, ByVal bShow As Long) As Long
    Private Const SB_HORZ = 0
    Private Const SB_VERT = 1
    Private Const SB_BOTH = 3
'-----------------------------

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'===Deklarasi Repaint Objek Form===
    Private Declare Function LockWindowUpdate Lib _
        "user32" (ByVal hWndLock As Long) As Long
'-----------------------------

    Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
    End Type

    Private Declare Function FindFirstFile Lib "kernel32" _
       Alias "FindFirstFileA" _
       (ByVal lpFileName As String, _
       lpFindFileData As WIN32_FIND_DATA) As Long

    Private Declare Function FindClose Lib "kernel32" _
       (ByVal hFindFile As Long) As Long
'--------------------------------

    Private Declare Function Inp Lib "inpout32.dll" _
    Alias "Inp32" (ByVal PortAddress As Integer) As Integer
    Private Declare Sub Out Lib "inpout32.dll" _
    Alias "Out32" (ByVal PortAddress As Integer, ByVal value As Integer)


    Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
      (ByVal hWnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Long, _
       lParam As Any) As Long

    Private Declare Function SetWindowLongA Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long


    Private Const LVM_SETITEMCOUNT As Long = 4096 + 47
    Private Const CB_SHOWDROPDOWN = &H14F
    Private Const GWL_WNDPROC = -4

Private Declare Function GetTickCount Lib "kernel32" () As Long


Dim rsAdo  As ADODB.Recordset

Dim login As Boolean
Dim PlaySong As Boolean
Dim l As Boolean
Dim Tstart As Date
Dim VcboOnclick As Boolean
Dim vplaylist As Boolean
Dim vKey As Integer
Public vTempo As Integer
Dim vWaktumulai As Date    'waktu login
Dim vtetapfokus As Integer  'fokus tetap di player waktu promo mulai
Dim vpbCameraStandar  ' PAL / NTSC
Dim vLstAllRow As Long
Dim vlstAllRecord As Integer 'Berapa record yg masuk ke lstAll
Public vpbJumlahNewPopuler As Integer
Public vMovieKategori As Integer
Public vTambahWaktu As Boolean  'mau tambah apa nggak
Public vHabisWaktu As Boolean 'sudah ditutup tinggal tunggu lagu terakhir
Public vpbDurasi As Integer
Private VnamaKom As String
Dim vcbodown As Boolean
Public caretLebar As Integer 'lebar dan tinggi caret
Public caretTinggi As Integer 'lebar dan tinggi caret
Public vtemp As Boolean
Public iSockets As Integer 'socket call
Public vMemberID As String
Public picfile As String
Public vrekamstate As Boolean 'rekam = true
Public vVideo As Integer
      '0=com 1=midi 2=fullscreen 3=chat 5=movie 6=freeze 7=tv 8=welcome 9=call
      '10=saran 11=about 12=help
Public vVideoAktif As Boolean
Public vpointer As Integer 'pointer ada di mana 0=list,1=title,2=artist, 3 playlist, 4=txtchat, 5=lstTV , 6=code
Public vpointerTemp As Integer 'pointer sementara
Dim HurufDrive As String
Dim vVocal As Integer 'vokalnya 0,1=ML 2,3=MR 4=ST

Public vExtensi As Boolean ' added by Andi 17-12-2020
Public vMPG As Boolean 'added by Andi 09-01-2021

'Dim vExtensi As Boolean ' added by Andi 17-12-2020

Public vVocalterus As Boolean
Dim vVolvocal As Integer 'repeating timer sampe berapa detik
Dim KeyboarAktif As Boolean 'false = tidak aktif / true = aktif
Dim lokasi As String
Public ScreenSaverAktif As Integer

'BUAT VOLUME MASTER
Dim rc As Long
Dim ok As Boolean

Dim VolMaster As Integer
Dim VolWave As Integer
Dim VolMic As Integer
Dim VolTV As Integer
Dim volcd As Integer
Dim volLine As Integer

Dim VolTemp As Long
Dim selIndex, selIndexPL As Integer 'buat simpen posisi index listview
Dim lErr As Long

Private ClientRemoteIp As Collection

Private promoSongCollection As Collection
Private promoSongCollectionCurrent As Long
Private promoSongPlay As Long
Private promoSongPlayCurrent As Long

Public promoImageCollection As Collection
Public promoImageCollectionCurrent As Long

Public promoAnimationCollection As Collection
Public promoAnimationCollectionCurrent As Long

Private lastInputTick As Long

Public Function FileExists(sSource As String) As Boolean

  On Error Resume Next

  Dim WFD As WIN32_FIND_DATA
  Dim hFile As Long

  hFile = FindFirstFile(sSource, WFD)
  FileExists = hFile <> INVALID_HANDLE_VALUE

  Call FindClose(hFile)


  If Err.Number <> 0 Then
    LogError Me.Name, "FileExists"
  End If
End Function

Public Sub pencet_btnplaylist()

    On Error Resume Next

    If lstPlaylist.Visible Then
        Exit Sub
    End If

    lstAll.Visible = False
    lstPlaylist.Visible = True
    frmTransparent.GantiSkin (5)
    DoEvents

    lstPlaylist.SetFocus
    ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
    vpointerTemp = vpointer
    vpointer = 3
    lErr = LockWindowUpdate(0)


    If Err.Number <> 0 Then
      LogError Me.Name, "pencet_btnplaylist"
    End If
End Sub

Public Sub btnRecStop_Click()

    On Error Resume Next

    vrekamstate = False
    btnRecStop.Visible = False

    frmCamera.Hide
    tmrRecord.Enabled = False
    lblRecording.Visible = False
    frmCamera.VideoCap1.CaptureMode = False

    Unload frmCamera


    If Err.Number <> 0 Then
      LogError Name, "btnRecStop_Click"
    End If
End Sub

Private Sub cboKategori_Click()

    On Error Resume Next

    If VcboOnclick Then
        Call SendMessage(cbokategori.hWnd, CB_SHOWDROPDOWN, False, ByVal 0)
        If vVideo = 0 Then
            CariKategori
        End If
        If vVideo = 1 Then
            midiKategori
        End If
        If vVideo = 5 Then
            moviecari
            txtSearch.SetFocus
        End If
    End If


    If Err.Number <> 0 Then
      LogError Name, "cboKategori_Click"
    End If
End Sub

Private Sub cbokategori_GotFocus()

    On Error Resume Next

    vcbodown = True 'sedang aktif / dalam posisi dropdown
    picKategori.Visible = False
    txtSearch.text = ""


    If Err.Number <> 0 Then
      LogError Name, "cbokategori_GotFocus"
    End If
End Sub

Private Sub cboKategori_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If KeyAscii = 13 Then
        On Error Resume Next
        If vVideo = 0 Then
            cmdSong_Click
        End If
        If vVideo = 1 Then
            midiKategori
        End If
        If vVideo = 5 Then
            moviecari
            txtSearch.SetFocus
        End If
    End If

    If Err.Number <> 0 Then
      LogError Name, "cboKategori_KeyPress"
    End If
End Sub

Private Sub cboKategori_LostFocus()

    On Error Resume Next

    VcboOnclick = True
    vcbodown = False 'sedang aktif / dalam posisi dropdown


    If Err.Number <> 0 Then
      LogError Name, "cboKategori_LostFocus"
    End If
End Sub

Public Sub cmdClear_Click()

  On Error Resume Next

  If login = False Then
    If lstPlaylist.ListItems.Count = 0 Then
      GoTo lblEnd
    Else
      lstPlaylist.ListItems.Clear
      txtNextSong.text = ""
          '--- HAPUS DATA PLAYLIST DI DATABASE ---'
          Dim idMember As String
              Dim Sql As String
              Dim myrs As MYSQL_RS
          If txtLogin = "" Then
              idMember = "00000"
              Sql = "DELETE FROM playlist " & _
                    "WHERE USERID = '" & idMember & _
                    "'  AND ROOM = '" & txtCompName.text & "';"
              Set myrs = MyConn.Execute(Sql)
          Else
              idMember = txtLogin.text
              Sql = "DELETE FROM playlist " & _
                    "WHERE USERID = '" & idMember & _
                    "';"
              Set myrs = MyConn.Execute(Sql)
          End If
    End If
  Else
    If lstPlayUser.ListItems.Count = 0 Then
      GoTo lblEnd
    Else
      lstPlayUser.ListItems.Clear
      txtNextSong.text = ""
    End If
  End If

  ClientRemotePlaylist


lblEnd:

  If Err.Number <> 0 Then
    LogError Me.Name, "cmdClear_Click"
  End If
End Sub

Public Sub cmdDelete_Click()

  On Error Resume Next

  If lstPlaylist.ListItems.Count > 0 Then

    ServerRemoteRemove lstPlaylist.selectedItem.index

    ClientRemotePlaylist

  End If


lblEnd:

  If Err.Number <> 0 Then
    LogError Me.Name, "cmdDelete_Click"
  End If

End Sub

Public Sub cmdDown_Click()
    On Error Resume Next
    Dim itmx As ListItem
    If lstPlaylist.selectedItem.index = lstPlaylist.ListItems.Count Then
        Set lstPlaylist.selectedItem = lstPlaylist.ListItems(lstPlaylist.ListItems.Count)
        Set lstPlaylist.DropHighlight = lstPlaylist.selectedItem
    Else
        Set itmx = lstPlaylist.ListItems.add(lstPlaylist.selectedItem.index + 2, , lstPlaylist.selectedItem.text)
            itmx.SubItems(1) = lstPlaylist.selectedItem.SubItems(1)
            itmx.SubItems(2) = lstPlaylist.selectedItem.SubItems(2)


        lstPlaylist.ListItems.Remove (lstPlaylist.selectedItem.index)

        savePlayList

        ClientRemotePlaylist

        Set lstPlaylist.selectedItem = lstPlaylist.ListItems(lstPlaylist.selectedItem.index + 1)
        Set lstPlaylist.DropHighlight = lstPlaylist.selectedItem

    End If
    txtNextSong.text = lstPlaylist.ListItems.Item(1)
End Sub

Public Sub cmdFast_Click()
    On Error Resume Next
    If PlaySong = False Then
        Exit Sub
    End If
    ScoreValid = False
    If frmVideo.pbDurasiAkhir < 10 Then
        frmVideo.pbDurasiAkhir = 0
    Else
        frmVideo.pbDurasiAkhir = frmVideo.pbDurasiAkhir - 10
    End If
    frmVideo.WindowsMediaPlayer1.Controls.currentPosition = frmVideo.pbDurasiAkhir
End Sub


Public Sub cmdLogout_Click()
    On Error Resume Next
    login = False
    cmdLogout.Visible = False
    LogOut
End Sub

'PAUSE
Public Sub cmdPause_Click(index As Integer)
    On Error Resume Next
    If PlaySong = False Then
        Exit Sub
    End If
    If index = 0 Then
       frmVideo.WindowsMediaPlayer1.Controls.pause
       cmdPause(0).Visible = False
       cmdPause(1).Visible = True
    End If
    'Play Again
    If index = 1 Then
       frmVideo.WindowsMediaPlayer1.Controls.play
       cmdPause(0).Visible = True
       cmdPause(1).Visible = False
    End If
End Sub

Public Sub cmdPlay_Click()

    On Error Resume Next

    Dim i As Integer

    If vpbBlackBox = 2 Then
      frmUser.turnDiscoLampOn
    End If

    If PlaySong = True Then
        waktuhabis
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        Unload frmPromo
        PlaySong = False
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
        Label4.Caption = "00:00"
        Label3.Caption = "00:00"
        LockRoom
             If lstAll.Visible = True Then
                PlayLstAll
             ElseIf lstPlayUser.Visible = True Then
                PlayLstUser
             ElseIf lstPlaylist.Visible = True Then
                If lstPlaylist.ListItems.Count = 0 Then
                     GoTo lblEnd
                Else
                    If vrekamstate = True Then
                        btnRecStop_Click
                    End If

                    PlayLstPlaylist
                    i = lstPlaylist.selectedItem.index
                        lErr = LockWindowUpdate(lstPlaylist.hWnd)
                        lstPlaylist.ListItems.Remove (i)
                        savePlayList
                    If i > 1 Or lstPlaylist.ListItems.Count > 0 Then
                        lstPlaylist.selectedItem.Selected = True
                    End If
                        ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
                        lErr = LockWindowUpdate(0)
                End If
             End If
    Else
            Unload frmPromo
            If lstAll.Visible = True Then
                PlayLstAll
            ElseIf lstPlayUser.Visible = True Then
                PlayLstUser
            ElseIf lstPlaylist.Visible = True Then
                If lstPlaylist.ListItems.Count = 0 Then
                    GoTo lblEnd
                Else
                    PlayLstPlaylist
                    i = lstPlaylist.selectedItem.index
                    lErr = LockWindowUpdate(lstPlaylist.hWnd)
                    lstPlaylist.ListItems.Remove (lstPlaylist.selectedItem.index)
                    savePlayList
                    If i > 1 Or lstPlaylist.ListItems.Count > 0 Then
                        lstPlaylist.selectedItem.Selected = True
                    End If
                    ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
                    lErr = LockWindowUpdate(0)
                End If
            End If
    End If

    ClientRemotePlaylist


lblEnd:

    If Err.Number <> 0 Then
      LogError Name, "cmdPlay_Click"
    End If
End Sub

Public Sub cmdRequest_Click()

    On Error Resume Next

    Dim LV As ListItem
    Dim myrs As MYSQL_RS
    Dim Sql As String
    Dim add As Boolean

    If (lstAll.Visible) And (lstAll.ListItems.Count) = 0 Then
        GoTo lblEnd
    End If

    If PlaySong = False Then
        cmdPlay_Click
    Else

        If lstPlaylist.ListItems.Count = 0 Then
            add = True
        Else
            If lstPlaylist.ListItems(lstPlaylist.ListItems.Count).SubItems(2) <> lstAll.ListItems(selIndex).SubItems(3) Then
                add = True
            Else
                add = False
            End If
        End If

        If add = True Then

            If lstAll.ListItems(selIndex).SubItems(7) = "1" Then

                Sql = "SELECT TITLE, SINGER FROM masters Where IDMUSIC = " & lstAll.ListItems(selIndex).SubItems(3) & " ;"
                Set myrs = MyConn.Execute(Sql)

                Set LV = lstPlaylist.ListItems.add(, , myrs.Fields(0).value)
                LV.SubItems(1) = myrs.Fields(1).value
            Else
                Set LV = lstPlaylist.ListItems.add(, , lstAll.ListItems(selIndex).SubItems(1))
                LV.SubItems(1) = lstAll.ListItems(selIndex).SubItems(2)
            End If

            LV.SubItems(2) = lstAll.ListItems(selIndex).SubItems(3)
            LV.SubItems(3) = lstAll.ListItems(selIndex).SubItems(4)
            LV.SubItems(4) = lstAll.ListItems(selIndex).SubItems(5)
            LV.SubItems(5) = lstAll.ListItems(selIndex).SubItems(6)

            DoEvents


            savePlayList
        End If
        txtNextSong.text = lstPlaylist.ListItems.Item(1)

        ClientRemotePlaylist
    End If

lblEnd:

    If Err.Number <> 0 Then
      LogError Me.Name, "PlayLagu"
    End If

End Sub


Public Sub cmdSlow_Click()
    On Error Resume Next
    If PlaySong = False Then
        Exit Sub
    End If

    ScoreValid = False
    If frmVideo.pbDurasi >= frmVideo.pbDurasiAkhir + 10 Then
        cmdStop_Click
    Else
        frmVideo.pbDurasiAkhir = frmVideo.pbDurasiAkhir + 10
        frmVideo.WindowsMediaPlayer1.Controls.currentPosition = frmVideo.pbDurasiAkhir
    End If
End Sub

Public Sub cmdSong_Click()
    On Error Resume Next
    If vpointer = 3 Then
        vpointer = vpointerTemp
    End If

    If vpbHits = True Then
        frmTransparent.GantiSkin (2)
    ElseIf vpbNew = True Then
        frmTransparent.GantiSkin (3)
    ElseIf vpbPopuler = True Then
        frmTransparent.GantiSkin (4)
    Else
        frmTransparent.GantiSkin (0)
    End If

    lstPlaylist.Visible = False
    lstAll.Visible = True

    If Not lstAll.ListItems.Count > 0 Then
        CariTitle
    End If

    txtCategory.text = cbokategori.text

    If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
        If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
            lstAll.SetFocus
            lstAll.ListItems(selIndex).Selected = True
            Set lstAll.DropHighlight = lstAll.selectedItem
            selIndex = lstAll.selectedItem.index
            lstAll.selectedItem.Selected = False
        End If
        sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    End If
End Sub

Public Sub CariKategori()
    On Error Resume Next
    If vpointer = 3 Then
        vpointer = vpointerTemp
    End If
    Dim lokasi As String
    If vpbHits = True Then
        lokasi = App.Path + "\picture\normalscreen\"
        'lstAll.Picture = LoadPicture(lokasi + "hit.jpg")
    ElseIf vpbNew = True Then
        lokasi = App.Path + "\picture\normalscreen\"
    '    lstAll.Picture = LoadPicture(lokasi + "new.jpg")
    ElseIf vpbPopuler = True Then
        lokasi = App.Path + "\picture\normalscreen\"
    '    lstAll.Picture = LoadPicture(lokasi + "popular.jpg")
    Else
        lokasi = App.Path + "\picture\normalscreen\"
    '    lstAll.Picture = LoadPicture(lokasi + "songlist.jpg")
    End If

    lstPlaylist.Visible = False
    lstAll.Visible = True

        CariTitle

    txtCategory.text = cbokategori.text

    If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
        If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
            selIndex = 1
            lstAll.ListItems(1).Selected = True
            Set lstAll.DropHighlight = lstAll.selectedItem
            lstAll.selectedItem.Selected = False
        End If
        sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    End If
    DoEvents
End Sub

Public Sub cmdStop_Click()

    On Error Resume Next

    If promoSongPlay > 0 And promoSongCollection.Count > 0 Then
      promoSongPlayCurrent = promoSongPlayCurrent + 1
      If promoSongPlayCurrent = promoSongPlay Then
        ServerRemotePlay promoSongCollection(promoSongCollectionCurrent)
        If promoSongCollectionCurrent = promoSongCollection.Count Then
          promoSongCollectionCurrent = 1
        Else
          promoSongCollectionCurrent = promoSongCollectionCurrent + 1
        End If
        promoSongPlayCurrent = 0
        Exit Sub
      End If
    End If


    If Not (vVideoAktif) Then
        Unload frmVideo
        Exit Sub
    End If

    vKey = 0
    vTempo = 0

    tmrVokal.Enabled = False
    If vpbBlackBox = 2 Then
        frmUser.turnDiscoLampOff
    End If
    If lstPlaylist.ListItems.Count = 0 Then
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        frmPromo.Show
        frmTransparent.Show
        frmRoom.Show
        PlaySong = False
        txtPlaying.text = ""
        If vVideo = 7 Then
            Form2.Show
            frmCamera.Show
        End If
    Else
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        Unload frmPromo
        PlayLst
        If promoImageCollection.Count > 0 Then
          frmPromoImage.Show
        End If
        frmTransparent.Show
        frmRoom.Show
        If vpbBlackBox = 2 Then
            frmUser.turnDiscoLampOn
        End If
        txtPlaying.text = lstPlaylist.ListItems.Item(1) + " - " + lstPlaylist.ListItems.Item(1).SubItems(1)
        If lstPlaylist.ListItems.Count > 0 Then
            lErr = LockWindowUpdate(lstPlaylist.hWnd)
            lstPlaylist.ListItems.Remove (1)
            savePlayList
            ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
            lErr = LockWindowUpdate(0)
        End If
    End If
    If lstPlaylist.ListItems.Count = 0 Then
        txtNextSong.text = "NO SONG"
    Else
        txtNextSong.text = lstPlaylist.ListItems.Item(1)
    End If

       If vrekamstate = True Then
            btnRecStop_Click
       End If

    '===FOCUS POINTER
       If vpbFrmCountry = True Then
            frmCountry.lstCountry.SetFocus
        ElseIf vpbFrmCategory = True Then
            frmCategory.Text1.SetFocus
        ElseIf vpbFrmConfirmasi = True Then
            frmConfirmasi.Text2.SetFocus
        ElseIf vpbFrmCall = True Then
            frmCall.txtPesan.SetFocus
        ElseIf vpbfrmTambahJam = True Then
            frmTambahJam.Text1.SetFocus
        ElseIf vpbfrmSaran = True Then
            frmSaran.txtPesan.SetFocus
        ElseIf vpbfrmNew = True Then
            frmNew.lstTop.SetFocus
        ElseIf vpbfrmPopuler = True Then
            frmPopuler.lstTop.SetFocus
        ElseIf vpbfrmAbout = True Then
            frmabout.flsLogo.SetFocus
        ElseIf vpbfrmHelp = True Then
            frmHelp.flsLogo.SetFocus
        ElseIf vpbfrmVocal = True Then
            frmVocal.Text1.SetFocus
        Else
            If (frmRoom.vVideo = 0) Then
                Form2.Show
                frmTransparent.Show
                If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                ElseIf frmRoom.vpointer = 3 Then
                    frmRoom.lstPlaylist.Visible = True
                    frmRoom.lstAll.Visible = False
                    frmRoom.lstPlaylist.SetFocus
                End If
            ElseIf vVideo = 5 Then
                Form2.Show
                frmTransparent.Show
                If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                ElseIf frmRoom.vpointer = 3 Then
                    frmRoom.lstPlaylist.Visible = True
                    frmRoom.lstAll.Visible = False
                    frmRoom.lstPlaylist.SetFocus
                End If
            ElseIf vVideo = 3 Then
                sMakeCaret frmRoom.txtChat, frmRoom.caretLebar, frmRoom.caretTinggi
            ElseIf vVideo = 7 Then  'TV
                lstTV.SetFocus
            End If
        End If

        LockRoom
        vtetapfokus = 0

    waktuhabis

    ClientRemotePlaylist


    If Err.Number <> 0 Then
      LogError Name, "cmdStop_Click"
    End If
End Sub

Public Sub cmdUp_Click()

    On Error Resume Next

    Dim selectedIndex As Long
    Dim title As String
    Dim singer As String
    Dim Path As String
    Dim analog As String
    Dim vol As String
    Dim idMusic As String
    Dim li As ListItem

    selectedIndex = lstPlaylist.selectedItem.index


    If selectedIndex = 1 Then

        Set lstPlaylist.DropHighlight = lstPlaylist.selectedItem
    Else

        title = lstPlaylist.ListItems(selectedIndex).text
        singer = lstPlaylist.ListItems(selectedIndex).ListSubItems(1).text
        idMusic = lstPlaylist.ListItems(selectedIndex).ListSubItems(2).text
        Path = lstPlaylist.ListItems(selectedIndex).ListSubItems(3).text
        analog = lstPlaylist.ListItems(selectedIndex).ListSubItems(4).text
        vol = lstPlaylist.ListItems(selectedIndex).ListSubItems(5).text

        lstPlaylist.ListItems.Remove selectedIndex

        Set li = lstPlaylist.ListItems.add(selectedIndex - 1, , title)
        li.SubItems(1) = singer
        li.SubItems(2) = idMusic
        li.SubItems(3) = Path
        li.SubItems(4) = analog
        li.SubItems(5) = vol

        Set lstPlaylist.selectedItem = lstPlaylist.ListItems(selectedIndex - 1)
        Set lstPlaylist.DropHighlight = lstPlaylist.ListItems(selectedIndex - 1)

        DoEvents


        savePlayList

        ClientRemotePlaylist
    End If

    txtNextSong.text = lstPlaylist.ListItems.Item(1)


    If Err.Number <> 0 Then
      LogError Name, "cmdUp_Click"
    End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    ScreenSaverAktif = 0
    If UkuranVideo <> 1 Then
        If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode = 8) Or (KeyCode >= 48 And KeyCode <= 57) Or _
           (KeyCode >= 37 And KeyCode <= 40) Then
            ScreenSaverAktif = 0
            Maksimal
        End If
    End If

    If Err.Number <> 0 Then
      LogError Name, "Form_KeyDown"
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
        sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    End If

    If Err.Number <> 0 Then
      LogError Name, "Form_KeyPress"
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Dim ComName As String * 255, cname As String
    Dim x As Long
    Dim ch As ColumnHeader
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim idMusic As String
    Dim promoDirectory As String
    Dim promoImagePath As String
    Dim promoAnimationPath As String

    lastInputTick = 0
    
    vExtensi = False 'added by Andi 21-12-2020
    vMPG = False 'added by Andi 09-01-2021

    ScoreSetup = True
    x = GetComputerName(ComName, 255)
    VnamaKom = Trim(ComName)
    VnamaKom = Left(VnamaKom, Len(VnamaKom) - 1)
    vVocalterus = False
    vpbMember = ""

    AktifUkuran

    vpointer = 1
    vpointerTemp = 1
    vpbBolehMainVocal = True
    ScreenSaverAktif = 0

    vpbBlackBox = 2
    vpbMute = False
    vKey = 0
    vTempo = 0

    txtVol.text = 60
    vVideo = 0
    vVideoAktif = False
    Skin1.Tag = 0

    'Winsock Order
    iSockets = 0

    vMemberID = ""

    vHabisWaktu = False
    vTambahWaktu = True  'waktu masih ada

    PictureLogin.Height = 0

    VcboOnclick = True
    Tstart = Time

    'tampil
    lstAll.Visible = True
    txtDate.text = Date
    txtStartTime.text = Time
    Dim sqlm As String
    Dim Im As Integer
    'Dim chm As ColumnHeader

    txtCompName.text = VnamaKom

    'Aktifkan transparent color
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
    
     'edited by Andi 22-01-2021
    If frmUser.settingScreenResolution = "S-FULLHD" Then
        Me.Width = 1920 * Screen.TwipsPerPixelX
        Me.Height = 1080 * Screen.TwipsPerPixelY
        flsTitle.Movie = App.Path + "\picture\anim\titlev2"
    ElseIf frmUser.settingScreenResolution = "S-SD" Then
        Me.Width = 1024 * Screen.TwipsPerPixelX
        Me.Height = 768 * Screen.TwipsPerPixelY
        flsTitle.Movie = App.Path + "\picture\anim\title"
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        Me.Width = 1280 * Screen.TwipsPerPixelX
        Me.Height = 720 * Screen.TwipsPerPixelY
        flsTitle.Movie = App.Path + "\picture\anim\titlev2"
    Else
        MsgBox "Screen Resolution is not supported, Contact Administrator!!!", _
        vbExclamation + vbOKOnly, "Alert!!"
        Sleep 1000
        End
        Exit Sub
    End If
    'added by Andi 22-01-2021

    lokasi = App.Path
'    Skin1.LoadSkin lokasi + "\skin\main.skn"
'    Skin1.ApplySkinByName hWnd, "MainFormMask"

    flsLogo.Movie = lokasi + "\picture\normalscreen\mainmin"
    PlaySong = False
    login = False
    Me.Move (0 - Screen.Width), 0 'Move Form to Monitor
    hot ' hotkey remote
    '-------------------------------------------------------------------------
    lstAll.Enabled = True
    'lstAll.Picture = LoadPicture(lokasi + "\picture\normalscreen\songlist.jpg")

    lErr = LockWindowUpdate(0)

    tampil 'structure ListView
    midilist

    screennormal

    '----playlist midi
    'midilist
    lstMidi.Visible = False
    lstMidiMusic.Visible = False

    cmdPause(0).Enabled = False
    cmdPause(1).Enabled = False
    cmdStop.Enabled = False
    cmdSlow.Enabled = False
    cmdFast.Enabled = False
    cmdRepeat.Enabled = False

    Dim sqlt As String
    Dim MyRst As MYSQL_RS
    sqlt = "SELECT TYPE, TYPENAME FROM kategori ORDER BY TYPE" 'Get Kategori
    Set MyRst = MyConn.Execute(sqlt)

    cbokategori.Clear
    cbokategori.AddItem "ALL"
    While Not MyRst.EOF
        cbokategori.AddItem MyRst.Fields(1).value
        cbokategori.ItemData(cbokategori.NewIndex) = MyRst.Fields(0).value
        MyRst.MoveNext
    Wend


    sqlt = "SELECT volmaster, volwave, volmic, voltv, volcd, video FROM room WHERE ROOMNAME ='" & txtCompName.text & "';"
    Set MyRst = MyConn.Execute(sqlt)
    VolMaster = MyRst.Fields(0).value
    VolWave = MyRst.Fields(1).value
    VolMic = MyRst.Fields(2).value
    volLine = MyRst.Fields(2).value
    VolTV = MyRst.Fields(3).value
    volcd = MyRst.Fields(4).value

    vpbCameraStandar = MyRst.Fields(5).value

    setAudioEndPointVolumeMasterVolumeLevelPercent VolMaster

    txtVol.text = VolMaster

    loadplaylist

    Set MyRst = Nothing

    DoEvents


    updateStructure

    Set ClientRemoteIp = New Collection

    wsServerRemote.Bind


    Set MyRst = MyConn.Execute("select remoteCode from room where ROOMNAME = '" & txtCompName.text & "'")
    txtRemoteCode.text = MyRst.Fields("remoteCode").value
    MyRst.CloseRecordset
    Set MyRst = Nothing
    DoEvents


    Set promoSongCollection = New Collection 'create object

    Set MyRst = MyConn.Execute("select IDMUSIC from promo order by playOrder")

    While MyRst.EOF = False
      promoSongCollection.add MyRst.Fields("IDMUSIC").value
      MyRst.MoveNext
    Wend
    MyRst.CloseRecordset
    Set MyRst = Nothing
    DoEvents

    If promoSongCollection.Count > 1 Then
      For a = 1 To promoSongCollection.Count
        For b = 1 To promoSongCollection.Count
          Randomize
          c = Round((promoSongCollection.Count - 1) * Rnd) + 1
          idMusic = promoSongCollection.Item(c)
          promoSongCollection.Remove c
          promoSongCollection.add idMusic, , 1
        Next
      Next
      DoEvents
    End If

    promoSongCollectionCurrent = 1


    Set MyRst = MyConn.Execute("select promoSongPlay from room where ROOMNAME = '" & txtCompName.text & "'")
    promoSongPlay = MyRst.Fields("promoSongPlay").value
    MyRst.CloseRecordset
    Set MyRst = Nothing
    DoEvents

    If promoSongPlay > 0 Then
      promoSongPlay = promoSongPlay + 1
    End If

    promoSongPlayCurrent = 0


    konekServer1


    Set MyRst = MyConn.Execute("select promoDirectory from room where ROOMNAME = '" & txtCompName.text & "'")
    promoDirectory = MyRst.Fields("promoDirectory").value
    MyRst.CloseRecordset
    Set MyRst = Nothing
    DoEvents


    Set promoImageCollection = New Collection
    a = 1
    Do
      promoImagePath = "\promo-pc\" & promoDirectory & "\promo_" & a & ".jpg"
      If FileExists("\\" & vpbServerUtama & promoImagePath) = False And FileExists("\\" & vpbServerBackup & promoImagePath) = False Then
        Exit Do
      End If
      promoImageCollection.add promoImagePath
      a = a + 1
      DoEvents
    Loop While True

    If promoImageCollection.Count > 1 Then
      For a = 1 To promoImageCollection.Count
        For b = 1 To promoImageCollection.Count
          Randomize
          c = Round((promoImageCollection.Count - 1) * Rnd) + 1
          promoImagePath = promoImageCollection.Item(c)
          promoImageCollection.Remove c
          promoImageCollection.add promoImagePath, , 1
        Next
      Next
      DoEvents
    End If

    promoImageCollectionCurrent = 1


    Set promoAnimationCollection = New Collection
    a = 1
    Do
      promoAnimationPath = "\promo-pc\" & promoDirectory & "\promo_" & a & ".swf"
      If FileExists("\\" & vpbServerUtama & promoAnimationPath) = False And FileExists("\\" & vpbServerBackup & promoAnimationPath) = False Then
        Exit Do
      End If
      promoAnimationCollection.add promoAnimationPath
      a = a + 1
      DoEvents
    Loop While True

    If promoAnimationCollection.Count > 1 Then
      For a = 1 To promoAnimationCollection.Count
        For b = 1 To promoAnimationCollection.Count
          Randomize
          c = Round((promoAnimationCollection.Count - 1) * Rnd) + 1
          promoAnimationPath = promoAnimationCollection.Item(c)
          promoAnimationCollection.Remove c
          promoAnimationCollection.add promoAnimationPath, , 1
        Next
      Next
      DoEvents
    End If

    promoAnimationCollectionCurrent = 1


'-----------------------------------------------------------------------------------
    LockRoom

    vpointer = 1
    frmPromo.Show
    Form2.Height = 0
'    frmCamera.Visible = False

    DoEvents

    frmTransparent.Show
    frmTransparent.LoadSkin (0)
    frmLoading.Show
    frmLoading.Text1.SetFocus

    SetCursorPos Screen.Width, 0

    DoEvents


    If Err.Number <> 0 Then
      LogError Name, "Form_Load"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If vVideo = 0 Then
        HotKeyDeactivate Me.hWnd
        SetWindowLongA Me.hWnd, GWL_WNDPROC, oldProc
    End If

    'LogOut
    PlaySong = False
    l = False
    login = False
    Timer2.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    frmUser.turnDiscoLampOff
End Sub

Private Sub lstAll_DblClick()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub lstAll_GotFocus()
On Error Resume Next
    If lstAll.ListItems.Count > 0 Then
        lstAll.selectedItem.Selected = True
        Set lstAll.DropHighlight = Nothing
        vtemp = False
    End If
End Sub

Private Sub lstAll_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
        If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
            sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
        End If
    End If
End Sub

Private Sub lstAll_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = 0
    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
End Sub

Private Sub lstAll_LostFocus()
On Error Resume Next
    If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
            Set lstAll.DropHighlight = lstAll.selectedItem
            selIndex = lstAll.selectedItem.index
            lstAll.selectedItem.Selected = False
    End If
End Sub

Private Sub lstChat_LostFocus()
    On Error Resume Next
    lstMidi.selectedItem.Selected = True
End Sub

Private Sub lstMidi_GotFocus()
    On Error Resume Next
    lstMidi.selectedItem.Selected = True
    Set lstMidi.DropHighlight = Nothing
End Sub

Private Sub lstMidi_LostFocus()
    On Error Resume Next
    If lstMidi.ListItems.Count = 0 Then
        Exit Sub
    Else
        lstMidi.selectedItem.Selected = False
        Set lstMidi.DropHighlight = lstMidi.selectedItem
    End If
End Sub

Private Sub lstMidiMusic_GotFocus()
'    If lstMidiMusic.ListItems.Count > 0 Then
'        lstMidiMusic.selectedItem.Selected = True
'        Set lstMidiMusic.DropHighlight = Nothing
'    End If
End Sub

Private Sub lstMidiMusic_LostFocus()
'    If lstMidiMusic.ListItems.Count = 0 Then
'        Exit Sub
'    Else
'        lstMidiMusic.selectedItem.Selected = False
'        Set lstMidiMusic.DropHighlight = lstMidiMusic.selectedItem
'    End If
End Sub

Private Sub lstMovie_Click()
    On Error Resume Next
    If lstMovie.ListItems.Count <= 0 Then
        Exit Sub
    End If

    Dim sqlnih As String
    Dim myrsnih As MYSQL_RS
    sqlnih = "SELECT artis, sinopsis  FROM film Where ID = '" & lstMovie.selectedItem.SubItems(2) & "';"
    Set myrsnih = MyConn.Execute(sqlnih)
    txtArtis.text = myrsnih.Fields(0).value
    txtSinopsis.text = myrsnih.Fields(1).value

    Set myrsnih = Nothing
End Sub

Private Sub lstMovie_DblClick()
    On Error Resume Next
    If lstMovie.ListItems.Count > 0 Then
        PlayMovie
    End If
End Sub

Private Sub lstMovie_GotFocus()
    On Error Resume Next
    If lstMovie.ListItems.Count > 0 Then
        lstMovie.selectedItem.Selected = True
        Set lstMovie.DropHighlight = Nothing
    End If
End Sub

Private Sub lstMovie_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = vbKeyDown Then
            lstMovie_Click
        End If
End Sub

Private Sub lstMovie_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If lstMovie.ListItems.Count > 0 Then
        If KeyAscii = 13 Then
            PlayMovie
        End If
    End If
    sMakeCaret txtSearch, caretLebar, caretTinggi
    KeyAscii = 0

    If Err.Number <> 0 Then
      LogError Name, "lstMovie_KeyPress"
    End If

End Sub

Private Sub lstMovie_LostFocus()
    On Error Resume Next
    If lstMovie.ListItems.Count > 0 Then
        lstMovie.selectedItem.Selected = False
        Set lstMovie.DropHighlight = lstMovie.selectedItem
    End If
End Sub

Private Sub lstPlaylist_DblClick()
    On Error Resume Next
    If lstPlaylist.ListItems.Count = 0 Then
        Exit Sub
    Else
        cmdPlay_Click
    End If
End Sub

Private Sub lstPlaylist_GotFocus()
    On Error Resume Next
    vplaylist = False
    lstAll.Visible = False

    If lstPlaylist.ListItems.Count > 0 Then
        lstPlaylist.selectedItem.Selected = True
        Set lstPlaylist.DropHighlight = Nothing
    End If
End Sub

Private Sub lstPlaylist_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If lstPlaylist.ListItems.Count > 0 Then
        Set lstPlaylist.DropHighlight = Nothing
    End If
End Sub

Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ScreenSaverAktif = 0
    If UkuranVideo <> 1 Then
        If (KeyCode >= 37 And KeyCode <= 40) Then
            ScreenSaverAktif = 0
            Maksimal
        End If
    End If

    lErr = LockWindowUpdate(lstPlaylist.hWnd)
    ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
End Sub

Private Sub lstPlaylist_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If KeyAscii = 13 Then
        If lstPlaylist.ListItems.Count = 0 Then
            Exit Sub
        Else
            cmdPlay_Click
        End If
    End If

    If UkuranVideo <> 1 Then
        ScreenSaverAktif = 0
        Maksimal
    End If

    If Err.Number <> 0 Then
      LogError Name, "lstPlaylist_KeyPress"
    End If

End Sub

Private Sub lstPlaylist_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
    lErr = LockWindowUpdate(0)
End Sub

Private Sub lstPlaylist_LostFocus()
    On Error Resume Next
    If lstPlaylist.ListItems.Count > 0 Then
        lstPlaylist.selectedItem.Selected = False
        Set lstPlaylist.DropHighlight = lstPlaylist.selectedItem
    End If
End Sub

Public Sub lstTV_Click()
    On Error Resume Next

    setAudioEndPointVolumeMasterVolumeLevelPercent 0

   If lstTV.selectedItem.SubItems(2) = frmCamera.VideoCap1.channel Then
        '---------PAUSE COMP------------
        If PlaySong = True Then
            frmVideo.WindowsMediaPlayer1.Controls.pause
        End If
        Minimal
        ScreenSaverAktif = 99
    Else
        frmCamera.VideoCap1.channel = lstTV.selectedItem.SubItems(2)
    End If

    VolTemp = 0
    tmrVolume.Enabled = True
End Sub

Private Sub lstTV_GotFocus()
    On Error Resume Next
    vpointer = 5
    If lstTV.ListItems.Count > 0 Then
        lstTV.selectedItem.Selected = True
        Set lstTV.DropHighlight = Nothing
    End If
End Sub

Private Sub lstTV_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    ScreenSaverAktif = 0
    If UkuranVideo <> 1 Then
        If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode = 8) Or (KeyCode >= 48 And KeyCode <= 57) Or _
           (KeyCode >= 37 And KeyCode <= 40) Then
            ScreenSaverAktif = 0
            Maksimal
        End If
    End If

    lErr = LockWindowUpdate(lstTV.hWnd)
    ShowScrollBar lstTV.hWnd, SB_VERT, False
End Sub

Private Sub lstTV_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If KeyAscii = 13 Then
        If UkuranVideo = 1 Then
            frmRoom.lstTV_Click
        Else
            frmRoom.ScreenSaverAktif = 0
            frmRoom.Maksimal
        End If
    End If

    If Err.Number <> 0 Then
      LogError Name, "lstPlaylist_KeyPress"
    End If

End Sub

Private Sub lstTV_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ShowScrollBar lstTV.hWnd, SB_VERT, False
    lErr = LockWindowUpdate(0)
End Sub

Private Sub lstTV_LostFocus()
    On Error Resume Next
    If lstTV.ListItems.Count > 0 Then
        lstTV.selectedItem.Selected = False
        Set lstTV.DropHighlight = lstTV.selectedItem
    End If
End Sub

Private Sub picComboRight1_Click()
    On Error Resume Next
    cbokategori.SetFocus
    Call SendMessage(cbokategori.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
End Sub

Private Sub picComboRight2_Click()
    On Error Resume Next
    cbokategori.SetFocus
    Call SendMessage(cbokategori.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
End Sub

Private Sub picKategori_Click()
    On Error Resume Next
    cbokategori.SetFocus
    Call SendMessage(cbokategori.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
End Sub

Private Sub Skin1_DblClick(ByVal Source As ACTIVESKINLibCtl.ISkinObject)
    On Error Resume Next

    If Source.GetName = "btnshutdown" Then
        Call Shell("Shutdown /s /t 0")
    ElseIf Source.GetName = "btnclose" Then
        cmdStop_Click
        Unload frmVideo
        frmUser.Show
        frmUser.Timer1.Enabled = False
        Unload Me
        frmCamera.WindowState = 0
        frmCamera.VideoCap1.Width = 7530
        frmCamera.VideoCap1.Height = 5610
        frmCamera.VideoCap1.VideoStandard = vpbCameraStandar
        frmPromo.Show
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    On Error Resume Next
If iSockets > 0 Then
    Socket(iSockets).Close
    Unload Socket(iSockets)
    iSockets = iSockets - 1
End If
End Sub

Public Sub LogOut()
    On Error Resume Next
     If PlaySong = True Then
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        txtUser.text = ""
        Unload frmRoom
        frmUser.Show
     Else
        txtUser.text = ""
        Unload frmRoom
        frmUser.Show
     End If

    frmUser.turnDiscoLampOff
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    PictureLogin.Visible = True
    PictureLogin.Height = PictureLogin.Height + 50
    If PictureLogin.Height >= 1760 Then
        Timer2.Enabled = False
        If cmdLogin.Visible Then
            txtLogin.SetFocus
        End If
    End If
End Sub

Private Sub Timer4_Timer()
    On Error Resume Next
    PictureLogin.Height = PictureLogin.Height - 75
    If PictureLogin.Height <= 75 Then
        Timer4.Enabled = False
        PictureLogin.Height = 0
    End If
End Sub

Private Sub Timer5_Timer()
    On Error Resume Next
    Picture3.Refresh
    txtNextSong.Move txtNextSong.Left - 30
    If txtNextSong.Left < -txtNextSong.Width Then
       txtNextSong.Left = Picture3.Width
    End If
End Sub

Private Sub tmrAktif_Timer()

    On Error Resume Next

    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim ECHO As ICMP_ECHO_REPLY
    Dim remoteCode As String
    Dim Count As Long

    ScreenSaverAktif = ScreenSaverAktif + 1

    If ScreenSaverAktif > 100 Then ScreenSaverAktif = 100
    If (vpbFrmCountry Or vpbFrmCategory Or vpbFrmConfirmasi Or vpbfrmTambahJam Or _
        vpbfrmSaran Or vpbfrmAbout Or vpbfrmHelp Or vpbfrmVocal Or vpbfrmWelcome) Then
        ScreenSaverAktif = 0
    End If
    If UkuranVideo <> 2 Then
        If ScreenSaverAktif = 10 Then
            If Not (vVideo = 3) Or (vVideo = 1) Then
                Minimal
            End If
        End If
    End If

    Call Ping(vpbServer, ECHO)
    If ECHO.status <> 0 Then
        tmrAktif.Enabled = False
        Sleep 1000
        Call Ping(vpbServer, ECHO)
        If ECHO.status <> 0 Then
            Sleep 3000
            Call Ping(vpbServer, ECHO)
            If ECHO.status <> 0 Then
                DoEvents
                tmrAktif.Enabled = False
                If AktifServerStatus = 1 Then
                    konekServer2
                Else
                    konekServer1
                End If
                tmrAktif.Enabled = True
                GoTo lblEnd
            End If
        End If
        tmrAktif.Enabled = True
        GoTo lblEnd
    End If

    Sql = "SELECT APA, IDROOM, STATUS, WKTSTART, USERROOM, DURASI, CHAT, PLAYLIST, KATAJALAN, KIRIMLAGU, TAMBAHWAKTU from room where ROOMNAME ='" & txtCompName & "'"
    Set myrs = MyConn.Execute(Sql)
    If MyConn.Error.Number <> 0 Then
        tmrAktif.Enabled = False
        If AktifServerStatus = 1 Then
            konekServer2
        Else
            konekServer1
        End If
        GoTo lblEnd
    End If

    txtCompName.Tag = myrs.Fields(1).value

If Err.Number <> 0 Then
  LogError Name, "tmrAktif_Timer 1"
End If
    If ((myrs.Fields(0).value = "tutup") And ((vVideo = 5) Or (vVideo = 7) Or (vVideoAktif = False) Or (vHabisWaktu = True))) Or (myrs.Fields(0).value = "hajar") Then
        vpbRoomStatus = 0
        Set myrs = Nothing
        If vpbfrmWelcome = True Then
            frmWelcome.Show
            GoTo lblEnd
        End If
        cmdClear_Click
        prcTutupAktifForm
        If Not (vVideo = 0) Then
            PlayerKomputer
            DoEvents
        End If
        promoSongPlayCurrent = -1
        cmdStop_Click

        frmWelcome.Show
        frmRoom.vVideo = 8

        DoEvents
        frmUser.turnDiscoLampOff

        GoTo lblEnd
    ElseIf myrs.Fields(0).value = "T" Then 'tutup aplikasi
        Set myrs = Nothing
        HilangkanDriveServer
        End
        GoTo lblEnd
    ElseIf myrs.Fields(0).value = "S" Then    'shutdown
        Set myrs = Nothing
        HilangkanDriveServer
        Call Shell("Shutdown /s /t 0")
        End
        GoTo lblEnd
    ElseIf myrs.Fields(0).value = "R" Then 'restart
        Set myrs = Nothing
        HilangkanDriveServer
        Call Shell("Shutdown /r /t 0")
        End
        GoTo lblEnd
    ElseIf myrs.Fields(0).value = "welcome" Then 'welcome screen
        Set myrs = Nothing
        vpbRoomStatus = 1
        If vpbfrmWelcome = True Then
            frmWelcome.Show
            GoTo lblEnd
        End If
        cmdClear_Click
        prcTutupAktifForm
        If Not (vVideo = 0) Then
            PlayerKomputer
            DoEvents
        End If
        promoSongPlayCurrent = -1
        cmdStop_Click

        frmWelcome.Show
        frmRoom.vVideo = 8


        Do While True

            Randomize
            remoteCode = Mid(Round(Rnd, 6), 3)
            While Left(remoteCode, 1) = "0"
              remoteCode = Mid(remoteCode, 2)
            Wend
            remoteCode = right("654321" & remoteCode, 6)

            Set myrs = MyConn.Execute("select count(*) from room where remoteCode = '" & remoteCode & "'")
            Count = myrs.Fields(0).value
            myrs.CloseRecordset
            Set myrs = Nothing
            If Count = 0 Then
                Sql = "update room set remoteCode='" & remoteCode & "' where ROOMNAME='" & txtCompName.text & "'"
                MyConn.Execute Sql
                If MyConnBackup.State = MY_CONN_OPEN Then
                    MyConnBackup.Execute Sql
                End If
                txtRemoteCode.text = remoteCode
                Exit Do
            End If
        Loop

        Set ClientRemoteIp = Nothing
        Set ClientRemoteIp = New Collection

        GoTo lblEnd
    ElseIf (myrs.Fields(0).value = "buka") And (vpbfrmWelcome = True) Then  'Buka Welcome

        Set myrs = Nothing
        frmWelcome.BukaKaraoke

        GoTo lblEnd
    End If
If Err.Number <> 0 Then
  LogError Name, "tmrAktif_Timer 2"
End If

     '------WAKTU------'
     Dim vJam As Integer
     Dim vWaktuHabis As Date

     If myrs.Fields(2).value = "chekin" Then
         vWaktumulai = myrs.Fields(3).value
         vpbDurasi = myrs.Fields(5).value

         txtUser.text = myrs.Fields(4).value
         txtTime.text = TimeSelisih(Now)

         vWaktuHabis = vWaktumulai

         vJam = hour(vWaktumulai) + myrs.Fields(5).value
         If vJam > 23 Then
             vWaktuHabis = CDate(vWaktuHabis) + Int(vJam / 24)
             vJam = vJam - Int(vJam / 24) * 24
         End If
         vWaktuHabis = CDate(DateSerial(year(vWaktuHabis), month(vWaktuHabis), day(vWaktuHabis)) & " " & Str$(vJam) & ":" & minute(vWaktuHabis) & ":" & second(vWaktuHabis))
If Err.Number <> 0 Then
  LogError Name, "tmrAktif_Timer 3"
End If

         If ((hour(vWaktuHabis - Now) * CLng(3600)) + (minute(vWaktuHabis - Now) * CLng(60)) + second(vWaktuHabis - Now) = 600) And (vTambahWaktu = True) And (vWaktuHabis > Now) Then
             txtTime.ForeColor = &HFF&
         Else
             If ((hour(vWaktuHabis - Now) * CLng(60)) + minute(vWaktuHabis - Now) <= 9) And (vTambahWaktu = False) Then
                 If (txtTime.ForeColor = &H22A6B) Then
                     txtTime.ForeColor = &HFF&
                     txtUser.ForeColor = &HFF&
                 Else
                     txtTime.ForeColor = &H22A6B
                     txtUser.ForeColor = &H22A6B
                 End If
             Else
                 txtTime.ForeColor = &H22A6B
                 txtUser.ForeColor = &H22A6B
             End If

         End If

         If myrs.Fields(6).value = 1 Then
             UpdateChat
         End If

         'UPDATE PLAYLIST
         If myrs.Fields(7).value = 1 Then
             lstPlaylist.ListItems.Clear
             loadplaylist
         End If

         'UPDATE KATAJALAN
         If myrs.Fields(8).value = 1 Then
             Cinema.LoadKatajalan
         End If

         'UPDATE TERIMA LAGU
         If myrs.Fields(9).value = 1 Then
             TerimaLagu
         End If

     Else
         txtUser.text = modProject.brandName
         txtTime.text = ""
     End If
     Set myrs = Nothing

lblEnd:

    If Err.Number <> 0 Then
      LogError Me.Name, "tmrAktif_Timer"
    End If

End Sub

Public Sub tmrLstAllLoad_Timer()
    On Error Resume Next
    tmrLstAllLoad.Enabled = False
        If AktifServerStatus = 1 Then
            konekServer1
        Else
           konekServer2
        End If
'    Unload frmLoading
    tmrAktif_Timer
'    Unload frmCamera
    tmrAktif.Enabled = True
    sMakeCaret txtSearch, caretLebar, caretTinggi
End Sub

Private Sub tmrMainVocal_Timer()
    On Error Resume Next
    Unload frmVocal
    tmrMainVocal.Enabled = False
End Sub

Private Sub tmrNonVocalML_Timer()
  On Error Resume Next
  If (frmVideo.WindowsMediaPlayer1.playState = 3) Then
     tmrNonVocalML.Enabled = False
     If vVocalterus Then
        setvocal
     Else
        frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 2 'Music Only
     End If
  End If
End Sub

Private Sub tmrNonVocalMR_Timer()
  On Error Resume Next
  If (frmVideo.WindowsMediaPlayer1.playState = 3) Then
     tmrNonVocalMR.Enabled = False
    If vVocalterus Then
        setvocal
    Else
        frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 3 'Vocal Only
    End If
  End If
End Sub

Private Sub tmrPicKeyTempo_Timer()
    On Error Resume Next
    If UkuranVideo = 1 Then
        Maksimal
    ElseIf UkuranVideo = 2 Then
        Minimal
    End If
    tmrPicKeyTempo.Enabled = False
End Sub

Private Sub tmrRecord_Timer()

    On Error Resume Next

    Dim lspc, lbps, lnofc, ltnoc As Long

    If lblRecording.Visible Then
        lblRecording.Visible = False
    Else
        lblRecording.Visible = True
    End If

    Dim result As Long
    Dim MyFree As Currency
    result = GetDiskFreeSpace(HurufDrive & ":\", lspc, lbps, lnofc, ltnoc)
    If result = 0 Then btnRecStop_Click
    MyFree = (CCur(lspc) * CCur(lbps) * CCur(lnofc)) / CCur(1000000000)

    If MyFree <= 1 Then
        btnRecStop_Click
    End If


    If Err.Number <> 0 Then
      LogError Name, "tmrRecord_Timer"
    End If

End Sub

Private Sub tmrRemote_Timer()
    On Error Resume Next
    resetAmpli
    tmrRemote.Enabled = False
End Sub

Private Sub tmrVokal_Timer()
  On Error Resume Next
  If (frmVideo.WindowsMediaPlayer1.playState = 3) Then
     tmrVokal.Enabled = False
     If vVocalterus Then
        setvocal
     Else
        frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 1 'Stereo
     End If
  End If
End Sub

Private Sub tmrVolume_Timer()
    On Error Resume Next

    setAudioEndPointVolumeMasterVolumeLevelPercent VolTemp

    If VolTemp >= Val(txtVol.text) Then
        tmrVolume.Enabled = False
    End If
    VolTemp = VolTemp + 5
End Sub

Private Sub txtArtis_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    KeyAscii = 0
    If vpointer = 1 Then
        sMakeCaret txtSearch, caretLebar, caretTinggi
    End If

    If Err.Number <> 0 Then
      LogError Name, "txtArtis_KeyPress"
    End If

End Sub

Private Sub txtChat_GotFocus()
    On Error Resume Next
    vpointer = 4
    sMakeCaret txtChat, 7, 60
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        chatSend
    End If

    If Err.Number <> 0 Then
      LogError Name, "txtChat_KeyPress"
    End If

End Sub

Public Sub txtSearch_Change()
    On Error Resume Next
    If lstAll.Visible Then
        CariTitle
    End If

    If vVideo = 5 Then  'Movie
        moviecari
    End If
End Sub

Private Sub txtSearch_GotFocus()
    On Error Resume Next
    If vpointer = 3 Then
        lokasi = App.Path + "\picture\normalscreen\"
        flsLogo.Movie = lokasi + "\logo"
    End If
    lstPlaylist.Visible = False
    DoEvents

    If vVideo = 0 Then
        lstAll.Visible = True
        If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
            Set lstAll.DropHighlight = lstAll.selectedItem
            selIndex = lstAll.selectedItem.index
            lstAll.selectedItem.Selected = False
        End If
        sMakeCaret txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    End If
    If vVideo = 5 Then
      If lstMovie.Visible = True Then
        If lstMovie.ListItems.Count > 0 Then
            lstMovie.selectedItem.Selected = False
            Set lstMovie.DropHighlight = lstMovie.selectedItem
        End If
        sMakeCaret txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
      End If
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    sMakeCaret txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    If vVideo = 0 Then
        '---------com--------------
        If lstAll.Visible Then
            If KeyCode = vbKeyDown Then
                KeyCode = 0
                If lstAll.ListItems.Count = 0 Then
                    Exit Sub
                End If
                If (selIndex < 9) Then
                    If selIndex >= lstAll.ListItems.Count Then
                        selIndex = lstAll.ListItems.Count
                    Else
                        selIndex = selIndex + 1
                    End If
                    lstAll.ListItems(selIndex).Selected = True
                Else
                    TambahLstAll
                End If
                'LEWATKAN KALAU HURUF CINA
                If lstAll.ListItems(selIndex).SubItems(7) = "1" Then
                    If selIndex < 9 Then
                        If selIndex >= lstAll.ListItems.Count Then
                            selIndex = lstAll.ListItems.Count
                        Else
                            selIndex = selIndex + 1
                        End If
                        lstAll.ListItems(selIndex).Selected = True
                    Else
                        TambahLstAll
                    End If
                End If
            End If
            If KeyCode = vbKeyUp Then
                KeyCode = 0
                If lstAll.ListItems.Count = 0 Then
                    Exit Sub
                End If
                If selIndex > 1 Then
                    selIndex = selIndex - 1
                    lstAll.ListItems(selIndex).Selected = True
                Else
                    KurangLstAll
                End If
                If lstAll.ListItems(selIndex).SubItems(7) = "1" Then
                    If selIndex > 1 Then
                        selIndex = selIndex - 1
                        lstAll.ListItems(selIndex).Selected = True
                    Else
                        KurangLstAll
                    End If
                End If
            End If
            '--------------------------------------------
            If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
                    Set lstAll.DropHighlight = lstAll.selectedItem
                    selIndex = lstAll.selectedItem.index
                    lstAll.selectedItem.Selected = False
            End If

        End If
    End If

    If vVideo = 5 Then
        '--------- Movie ---------
            If KeyCode = vbKeyDown Then
                KeyCode = 0
                If lstMovie.ListItems.Count = 0 Then
                    Exit Sub
                End If
                If lstMovie.selectedItem.index < 9 Then
                    terus_Listview lstMovie, " ", 2
                Else
                    TambahLstMovie lstMovie
                End If
            End If
            If KeyCode = vbKeyUp Then
                KeyCode = 0
                If lstMovie.ListItems.Count = 0 Then
                    Exit Sub
                End If
                If lstMovie.selectedItem.index > 1 Then
                        naik_Listview lstMovie, " ", 2
                Else
                    KurangLstMovie lstMovie
               End If
           End If
            If lstMovie.ListItems.Count > 0 Then
                lstMovie.selectedItem.Selected = False
                Set lstMovie.DropHighlight = lstMovie.selectedItem
                lstMovie_Click
            End If
    End If

    If Err.Number <> 0 Then
      LogError Name, "txtSearch_KeyDown"
    End If
End Sub

Public Sub txtSearch_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    If KeyAscii = 13 Then
        If UkuranVideo <> 1 Then
            ScreenSaverAktif = 0
            Maksimal
            Exit Sub
        End If
    End If
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    ScreenSaverAktif = 0
    If vVideo = 0 Then
        '---------com-------------
        If KeyAscii = 13 Then
            KeyAscii = 0
            PlayLagu
        End If
        Exit Sub
    End If
    If vVideo = 1 Then
        If KeyAscii = 13 Then
            txtSearch.text = ""
        End If
        Exit Sub
    End If
    If vVideo = 5 Then
        If KeyAscii = 13 Then
            If lstMovie.ListItems.Count = 0 Then
                Exit Sub
            End If
            PlayMovie
            ScreenSaverAktif = 0
        End If
        Exit Sub
    End If

  If Err.Number <> 0 Then
    LogError Name, "txtSearch_KeyPress"
  End If

End Sub

Function Search_Listview(LV As ListView, SearchText As String, Column As Long, Optional SelectRow As Boolean = True, Optional MakeVisible As Boolean = True) As Long
    On Error Resume Next
Dim SearchLength As Long
Dim CurrentRow As Long
Dim result As Long

    SearchLength = Len(SearchText)
    If SearchLength = 0 Then
        Search_Listview = -1
        Exit Function
    End If

    result = -1
    SearchText = UCase$(SearchText)
    If Column = 1 Then
        For CurrentRow = 1 To LV.ListItems.Count
            If UCase$(Left$(LV.ListItems(CurrentRow).text, SearchLength)) = SearchText Then
                result = LV.ListItems(CurrentRow).index
                Exit For
            End If
        Next CurrentRow
    Else
        For CurrentRow = 1 To LV.ListItems.Count
            If UCase$(Left$(LV.ListItems(CurrentRow).ListSubItems(Column - 1).text, SearchLength)) = SearchText Then
                result = LV.ListItems(CurrentRow).index
                Exit For
            End If
        Next CurrentRow
    End If

    If result > -1 Then
        If SelectRow Then LV.ListItems(result).Selected = True
        If MakeVisible Then LV.ListItems(result).EnsureVisible
    End If
    Search_Listview = result
End Function

Function terus_Listview(LV As ListView, SearchText As String, Column As Long, Optional SelectRow As Boolean = True, Optional MakeVisible As Boolean = True) As Long
    On Error Resume Next
Dim SearchLength As Long
Dim CurrentRow As Long
Dim result As Long


    result = LV.selectedItem.index + 1
    If result > LV.ListItems.Count Then
    result = LV.ListItems.Count
    End If

    If result > -1 Then
        If SelectRow Then LV.ListItems(result).Selected = True
        If MakeVisible Then LV.ListItems(result).EnsureVisible
    End If
    'Search_Listview = Result
End Function

Function naik_Listview(LV As ListView, SearchText As String, Column As Long, Optional SelectRow As Boolean = True, Optional MakeVisible As Boolean = True) As Long
    On Error Resume Next
Dim SearchLength As Long
Dim CurrentRow As Long
Dim result As Long

    result = LV.selectedItem.index - 1
    If result < 1 Then
        result = 1
    End If

    If result > -1 Then
        If SelectRow Then LV.ListItems(result).Selected = True
        If MakeVisible Then LV.ListItems(result).EnsureVisible
    End If
    'Search_Listview = Result
End Function

Public Sub hot()
    On Error Resume Next
    If vpbRemoteStatus = 1 Then
        '****************************** REMOTE LAMA ***********************************
        oldProc = SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
        HotKeyActivate Me.hWnd, 0, vbKeyF1                  '1
        HotKeyActivate Me.hWnd, 0, vbKeyF2                  '2
        HotKeyActivate Me.hWnd, 0, vbKeyF3                  '3
        HotKeyActivate Me.hWnd, 0, vbKeyF4                  '4
        HotKeyActivate Me.hWnd, 0, vbKeyF5                  '5
        HotKeyActivate Me.hWnd, 0, vbKeyF6                  '6
        HotKeyActivate Me.hWnd, 0, vbKeyF7                  '7
        HotKeyActivate Me.hWnd, 0, vbKeyF8                  '8
        HotKeyActivate Me.hWnd, 0, vbKeyF9                  '9
        HotKeyActivate Me.hWnd, 0, vbKeyF10                 '10
        HotKeyActivate Me.hWnd, 0, vbKeyF11                 '11

        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyA         '12
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyB         '13
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyC         '14
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyD         '15
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyE         '16
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyF         '17
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyG         '18
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyH         '19
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyI         '20
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyJ         '21
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyK         '22
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyL         '23
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyM         '24
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyN         '25
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyO         '26
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyP         '27
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyQ         '28
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyR         '29
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyS         '30
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyT         '31
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyU         '32
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyV         '33
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyW         '34
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyX         '35
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyY         '36
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyZ         '37

        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyA             '38
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyB             '39
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyC             '40
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyD             '41
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyE             '42
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyF             '43
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyG             '44
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyH             '45
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyI             '46
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyJ             '47
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyK             '48
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyL             '49
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyM             '50
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyN             '51
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyO             '52
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyP             '53
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyQ             '54
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyR             '55

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '56
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '57

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '58
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '59

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '60
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '61
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '62
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '63

        HotKeyActivate Me.hWnd, 0, vbKeyReturn              '64
        HotKeyActivate Me.hWnd, 0, vbKeyTab                 '65
        HotKeyActivate Me.hWnd, 0, vbKeyPageUp              '66
        HotKeyActivate Me.hWnd, 0, vbKeyPageDown            '67
        HotKeyActivate Me.hWnd, 0, vbKeyHome                '68
        HotKeyActivate Me.hWnd, 0, vbKeyEnd                 '69
        HotKeyActivate Me.hWnd, 0, vbKeyInsert              '70
        HotKeyActivate Me.hWnd, 0, vbKeyDelete              '71
        HotKeyActivate Me.hWnd, 0, vbKeyDelete              '72
        '**************************** REMOTE LAMA END**********************************
    Else
        oldProc = SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
        HotKeyActivate Me.hWnd, 0, vbKeyF1                  '1
        HotKeyActivate Me.hWnd, 0, vbKeyF2                  '2
        HotKeyActivate Me.hWnd, 0, vbKeyF3                  '3
        HotKeyActivate Me.hWnd, 0, vbKeyF4                  '4
        HotKeyActivate Me.hWnd, 0, vbKeyF5                  '5
        HotKeyActivate Me.hWnd, 0, vbKeyF7                  '6
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '7
        HotKeyActivate Me.hWnd, 0, vbKeyF6                  '8
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '9
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyQ         '10
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyP             '11

        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyA         '12
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyB         '13
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyC         '14
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyD         '15
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyE         '16
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyF         '17
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyG         '18
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyH         '19
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyM         '20
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyL         '21
        HotKeyActivate Me.hWnd, 0, vbKeyF8                  '22
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyV         '23
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyM         '24
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '25
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyK         '26
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '27
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyQ         '28
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '29
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyN         '30
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyO         '31
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '32
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyW         '33
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyX         '34
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyY         '35
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyY         '36
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyZ         '37

        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyA             '38
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyB             '39
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyC             '40
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyD             '41
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyE             '42
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyF             '43
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyG             '44
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyH             '45
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyI             '46
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyJ             '47
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyK             '48
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '49
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyL             '50
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyM             '51
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyN             '52
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyQ             '53
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyR             '54
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyR         '55

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '56
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '57

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '58
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '59

        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '60
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '61
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '62
        HotKeyActivate Me.hWnd, 0, vbKeySeparator           '63

        HotKeyActivate Me.hWnd, 0, vbKeyReturn              '64
        HotKeyActivate Me.hWnd, 0, vbKeyF11                 '65
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyS         '66
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyT         '67
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyU         '68
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyP         '69
        HotKeyActivate Me.hWnd, 0, vbKeyF9                  '70
        HotKeyActivate Me.hWnd, 0, vbKeyF10                 '71
        HotKeyActivate Me.hWnd, 0, vbKeyDelete              '72
        HotKeyActivate Me.hWnd, MOD_ALT, vbKeyO             '73
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyI         '74
        HotKeyActivate Me.hWnd, MOD_CONTROL, vbKeyJ         '75
    End If
End Sub

Private Sub roomId()
    On Error Resume Next
Dim Sql As String
Dim myrs As MYSQL_RS

Sql = "SELECT room.IDROOM, roomprice.ROOMID FROM room INNER JOIN roomprice ON room.ROOMID = roomprice.ROOMID where room.ROOMNAME = '" & txtCompName.text & "';"
Set myrs = MyConn.Execute(Sql)


txtIdRoom.text = myrs.Fields(0).value
txtRoomId.text = myrs.Fields(1).value
End Sub

Private Sub TampilAll()
    On Error Resume Next
Dim sqlm As String
Dim i As Double

Dim chm As ColumnHeader
Dim LV As ListItem
Dim MyRsm As MYSQL_RS

    lstAll.Enabled = True

    sqlm = "SELECT TITLE, SINGER, IDMUSIC FROM masters WHERE FLAG = 'Y' order by TITLE ASC, SINGER ASC;"
    Set MyRsm = MyConn.Execute(sqlm)

    For i = 1 To 10
        MyRsm.MoveNext
    Next i
    'i = 0
    lErr = LockWindowUpdate(lstAll.hWnd)
    Do Until MyRsm.EOF
        'Isi list data
        Set LV = lstAll.ListItems.add(, , (MyRsm.Fields(0).value))
        LV.SubItems(1) = MyRsm.Fields(1).value
        LV.SubItems(2) = MyRsm.Fields(2).value
        MyRsm.MoveNext
     '   If i Mod 500 = 0 Then
     '       DoEvents
     '       lErr = LockWindowUpdate(lstAll.hWnd)
     '       lErr = LockWindowUpdate(0)
     '   End If
      '  i = i + 1
    Loop
    lstAll.Enabled = True
    lErr = LockWindowUpdate(0)
        If AktifServerStatus = 1 Then
            konekServer1
        Else
           konekServer2
        End If
'    Unload frmLoading

    tmrAktif.Enabled = True
End Sub

Private Sub tampil()
    On Error Resume Next

    Dim chmU As UniToolbox2.ColumnHeader
    Dim chm As ColumnHeader
    Dim chp As ColumnHeader

    lstAll.Visible = True
    lstPlaylist.Visible = False

    lstAll.ColumnHeaders.Clear
    lstAll.ListItems.Clear
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)
    Set chmU = lstAll.ColumnHeaders.add(, , , 6100)
    Set chmU = lstAll.ColumnHeaders.add(, , , 5000)
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)
    Set chmU = lstAll.ColumnHeaders.add(, , , 0)

    lstPlaylist.ColumnHeaders.Clear
    lstPlaylist.ListItems.Clear
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 5900)
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 5145)
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 0)
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 0)
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 0)
    Set chp = lstPlaylist.ColumnHeaders.add(, , , 0)

    lstMovie.ColumnHeaders.Clear
    Set chm = lstMovie.ColumnHeaders.add(, , , 10)
    Set chm = lstMovie.ColumnHeaders.add(, , , 11250, lvwColumnCenter)
    Set chm = lstMovie.ColumnHeaders.add(, , , 10)

    lstTV.ColumnHeaders.Clear
    Set chm = lstTV.ColumnHeaders.add(, , , 10)
    Set chm = lstTV.ColumnHeaders.add(, , , 11250, lvwColumnCenter)
    Set chm = lstTV.ColumnHeaders.add(, , , 10)

    lstChat.ColumnHeaders.Clear
    Set chm = lstChat.ColumnHeaders.add(, , , 2000)
    Set chm = lstChat.ColumnHeaders.add(, , , 9000)
End Sub

Private Sub PlayLstAll()

    On Error Resume Next

    Dim PathLagu As String
    Dim PathLaguUtama As String
    Dim PathLaguBackup As String
    Dim ECHO As ICMP_ECHO_REPLY

    vTempo = 0
    vKey = 0
    Dim sqlnih As String
    Dim myrsnih As MYSQL_RS

    ScoreValid = True

    PathLagu = Replace$(lstAll.ListItems(selIndex).SubItems(4), "/", "\")
    If AktifServerStatus = 1 Then
        PathLaguUtama = "\\" & vpbServerUtama & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerBackup & "\Data\" & PathLagu
    Else
        PathLaguUtama = "\\" & vpbServerBackup & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerUtama & "\Data\" & PathLagu
    End If

    If FileExists(PathLaguUtama) Then
       frmVideo.Show
       frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
    Else
        Call Ping(vpbServerBackup, ECHO)
        If ECHO.status = 0 Then
            If FileExists(PathLaguBackup) Then
                frmVideo.Show
                frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
            Else
                If lstPlaylist.ListItems.Count = 0 Then
                    frmPromo.Show
                    frmTransparent.Show
                    frmRoom.Show
                    PlaySong = False
                Else
                        tmrVokal.Enabled = False
                        tmrNonVocalML.Enabled = False
                        tmrNonVocalMR.Enabled = False
                        frmVideo.WindowsMediaPlayer1.URL = ""
                        frmVideo.WindowsMediaPlayer1.Controls.stop
                        Unload frmVideo
                        Unload frmPromo
                        PlayLst
                        If vpbBlackBox = 2 Then
                            frmUser.turnDiscoLampOn
                        End If
                        If lstPlaylist.ListItems.Count > 0 Then
                            lstPlaylist.ListItems.Remove (1)
                            savePlayList
                            ClientRemotePlaylist
                        End If
                End If
                GoTo Populer
            End If
        Else
            If lstPlaylist.ListItems.Count = 0 Then
                frmPromo.Show
                frmTransparent.Show
                frmRoom.Show
                PlaySong = False
            Else
                    tmrVokal.Enabled = False
                    tmrNonVocalML.Enabled = False
                    tmrNonVocalMR.Enabled = False
                    frmVideo.WindowsMediaPlayer1.URL = ""
                    frmVideo.WindowsMediaPlayer1.Controls.stop
                    Unload frmVideo
                    Unload frmPromo
                    PlayLst
                    If vpbBlackBox = 2 Then
                        frmUser.turnDiscoLampOn
                    End If
                    If lstPlaylist.ListItems.Count > 0 Then
                        lstPlaylist.ListItems.Remove (1)
                        savePlayList
                        ClientRemotePlaylist
                    End If
            End If
            GoTo Populer
        End If
    End If

    frmVideo.WindowsMediaPlayer1.Controls.play
    PlaySong = True

    frmVideo.WindowsMediaPlayer1.settings.volume = 0
    'ANALOG---------------------------------------------------------------------------
    tmrVokal.Enabled = False
    
'    'added by Andi 14-12-2020
'    Dim extension As String
'    extension = fso.GetExtensionName(PathLagu)
'    If extension = "mp4" Or extension = "mpg" Then
'      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = True
'          tmrNonVocalMR.Enabled = False
'          vVocal = 2
'          vExtensi = True ' added by Andi 17-12-2020
'      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = True
'          vVocal = 3
'          vExtensi = True 'added by Andi 17-12-2020
'      End If
'    Else
'      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = True
'          tmrNonVocalMR.Enabled = False
'          vVocal = 2
'          vExtensi = False 'added by Andi 17-12-2020
'      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = True
'          vVocal = 3
'          vExtensi = False 'added by Andi 17-12-2020
'      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "ST" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = False
'          vVocal = 1
'          vExtensi = False 'added by Andi 17-12-2020
'      End If
'    End If
'    'added by Andi 14-12-2020
    
      'added by Andi 09-01-2021
    tmrVokal1.Enabled = False
    tmrVokal2.Enabled = False
    'added by Andi 09-01-2021
    
    'added by Andi 14-12-2020
    Dim extension As String
    extension = fso.GetExtensionName(PathLagu)
    If extension = "mp4" Then
      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
          tmrVokal1.Enabled = False
          tmrNonVocalML1.Enabled = True
          tmrNonVocalMR1.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
          tmrVokal1.Enabled = False
          tmrNonVocalML1.Enabled = False
          tmrNonVocalMR1.Enabled = True
          vVocal = 3
          vExtensi = True 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      End If
    ElseIf extension = "mpg" Then
      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
          tmrVokal2.Enabled = False
          tmrNonVocalML2.Enabled = True
          tmrNonVocalMR2.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
          vMPG = True 'added by Andi 09-01-2021
      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
          tmrVokal2.Enabled = False
          tmrNonVocalML2.Enabled = False
          tmrNonVocalMR2.Enabled = True
          vVocal = 3
          vExtensi = True 'added by Andi 17-12-2020
          vMPG = True 'added by Andi 09-01-2021
      End If
    Else
      'Original Code
      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "ST" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = False
          vVocal = 1
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      End If
    'Original Code
    End If
    'added by Andi 14-12-2020
    
    vVolvocal = 0

    frmVideo.WindowsMediaPlayer1.settings.volume = lstAll.ListItems(selIndex).SubItems(6)

    txtPlaying.text = lstAll.ListItems(selIndex).SubItems(1) + " - " + lstAll.ListItems(selIndex).SubItems(2)

'    Set myrsnih = Nothing

Populer:
    '===POPULER SONG===
    Dim x As Double
    sqlnih = "SELECT IDMUSIC, POPULER FROM masters WHERE IDMUSIC = " & lstAll.ListItems(selIndex).SubItems(3)
    Set myrsnih = MyConn.Execute(sqlnih)
    x = Val(myrsnih.Fields(1).value)
    Set myrsnih = Nothing
    sqlnih = "UPDATE masters SET POPULER= " & Str(x + 1) & " Where IDMUSIC = " & lstAll.ListItems(selIndex).SubItems(3)
    MyConn.Execute sqlnih


    If Err.Number <> 0 Then
      LogError Name, "PlayLstAll"
    End If
End Sub

Sub PlayMovie()
    On Error Resume Next
Dim PathLagu As String
Dim PathLaguUtama As String
Dim PathLaguBackup As String
Dim ECHO As ICMP_ECHO_REPLY

    vTempo = 0
    vKey = 0
    frmUser.turnDiscoLampOff

    Dim sqlnih As String
    Dim myrsnih As MYSQL_RS
    sqlnih = "SELECT PATH, VOL FROM film Where ID = '" & lstMovie.selectedItem.SubItems(2) & "';"
    Set myrsnih = MyConn.Execute(sqlnih)

    PathLagu = Replace$(myrsnih.Fields(0).value, "/", "\")

    If AktifServerStatus = 1 Then
        Call Ping(vpbServerBackup, ECHO)
        If ECHO.status = 0 Then
            PathLaguUtama = "\\" & vpbServerBackup & "\Movie\" & PathLagu
        Else
            PathLaguUtama = "\\" & vpbServerUtama & "\Movie\" & PathLagu
        End If
            PathLaguBackup = "\\" & vpbServerUtama & "\Movie\" & PathLagu
    Else
        Call Ping(vpbServerUtama, ECHO)
        If ECHO.status = 0 Then
            PathLaguUtama = "\\" & vpbServerUtama & "\Movie\" & PathLagu
        Else
            PathLaguUtama = "\\" & vpbServerBackup & "\Movie\" & PathLagu
        End If
            PathLaguBackup = "\\" & vpbServerBackup & "\Movie\" & PathLagu
    End If

    If FileExists(PathLaguUtama) Then
        ScoreValid = False
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        Unload frmPromo
        PlaySong = False
        frmVideo.Show
        frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
        frmVideo.WindowsMediaPlayer1.Controls.play
        PlaySong = True
        frmVideo.WindowsMediaPlayer1.settings.volume = myrsnih.Fields(1).value
    ElseIf FileExists(PathLaguBackup) Then
        ScoreValid = False
        tmrVokal.Enabled = False
        tmrNonVocalML.Enabled = False
        tmrNonVocalMR.Enabled = False
        frmVideo.WindowsMediaPlayer1.URL = ""
        frmVideo.WindowsMediaPlayer1.Controls.stop
        Unload frmVideo
        Unload frmPromo
        PlaySong = False
        frmVideo.Show
        frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
        frmVideo.WindowsMediaPlayer1.Controls.play
        PlaySong = True
        frmVideo.WindowsMediaPlayer1.settings.volume = myrsnih.Fields(1).value
   ' Else
   '     frmPromo.Show
   '     frmRoom.Show
   '     PlaySong = False
    End If

    txtPlaying.text = lstMovie.selectedItem.SubItems(1)

    vtetapfokus = 0
    Set myrsnih = Nothing

    frmTransparent.Show
    If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
        sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
    End If
End Sub

Private Sub PlayLstPlaylist()
On Error Resume Next
Dim PathLagu As String
Dim PathLaguUtama As String
Dim PathLaguBackup As String
Dim ECHO As ICMP_ECHO_REPLY

Dim MyRsP As MYSQL_RS
Dim SQLP As String

    vTempo = 0
    vKey = 0

    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    cmdFast.Enabled = True
    cmdSlow.Enabled = True
    cmdRepeat.Enabled = True
    cmdPause(0).Visible = True
    cmdPause(1).Visible = False
    If lstPlaylist.ListItems.Count = 0 Then
        Exit Sub
    End If

    '----HAPUS PLAYLIST DARI DATABASE----'
    Dim idMember As String
    Dim Sqld As String
    Dim myrsD As MYSQL_RS
    If txtLogin = "" Then
        idMember = "00000"
        Sqld = "DELETE FROM playlist " & _
              "WHERE USERID = '" & idMember & _
              "' AND IDMUSIC = " & lstPlaylist.selectedItem.SubItems(2) & _
              "  AND ROOM = '" & txtCompName.text & "';"
        Set myrsD = MyConn.Execute(Sqld)
        MyConn.Execute Sqld
    Else
        idMember = txtLogin.text
        Sqld = "DELETE FROM playlist " & _
              "WHERE USERID = '" & idMember & _
              "' AND IDMUSIC = " & lstPlaylist.selectedItem.SubItems(2) & _
              " ;"
        Set myrsD = MyConn.Execute(Sqld)
        MyConn.Execute Sqld
    End If

    ScoreValid = True

    PathLagu = Replace$(lstPlaylist.selectedItem.SubItems(3), "/", "\")
    If AktifServerStatus = 1 Then
        PathLaguUtama = "\\" & vpbServerUtama & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerBackup & "\Data\" & PathLagu
    Else
        PathLaguUtama = "\\" & vpbServerBackup & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerUtama & "\Data\" & PathLagu
    End If

    If FileExists(PathLaguUtama) Then
       frmVideo.Show
       frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
    Else
        Call Ping(vpbServerBackup, ECHO)
        If ECHO.status = 0 Then
            If FileExists(PathLaguBackup) Then
               frmVideo.Show
               frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
            Else
                If lstPlaylist.ListItems.Count = 0 Then
                    frmPromo.Show
                    frmTransparent.Show
                    frmRoom.Show
                    PlaySong = False
                Else
                        tmrVokal.Enabled = False
                        tmrNonVocalML.Enabled = False
                        tmrNonVocalMR.Enabled = False
                        frmVideo.WindowsMediaPlayer1.URL = ""
                        frmVideo.WindowsMediaPlayer1.Controls.stop
                        Unload frmVideo
                        Unload frmPromo
                        PlayLst
                        If vpbBlackBox = 2 Then
                            frmUser.turnDiscoLampOn
                        End If
                        If lstPlaylist.ListItems.Count > 0 Then
                            lstPlaylist.ListItems.Remove (1)
                            savePlayList
                            ClientRemotePlaylist
                        End If
                End If
                GoTo Populer
                Exit Sub
            End If
        Else
            If lstPlaylist.ListItems.Count = 0 Then
                frmPromo.Show
                frmTransparent.Show
                frmRoom.Show
                PlaySong = False
            Else
                    tmrVokal.Enabled = False
                    tmrNonVocalML.Enabled = False
                    tmrNonVocalMR.Enabled = False
                    frmVideo.WindowsMediaPlayer1.URL = ""
                    frmVideo.WindowsMediaPlayer1.Controls.stop
                    Unload frmVideo
                    Unload frmPromo
                    PlayLst
                    If vpbBlackBox = 2 Then
                        frmUser.turnDiscoLampOn
                    End If
                    If lstPlaylist.ListItems.Count > 0 Then
                        lErr = LockWindowUpdate(lstPlaylist.hWnd)
                        lstPlaylist.ListItems.Remove (1)
                        savePlayList
                        ClientRemotePlaylist
                        ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
                        lErr = LockWindowUpdate(0)
                    End If
            End If
            GoTo Populer
            Exit Sub
        End If
    End If

    frmVideo.WindowsMediaPlayer1.Controls.play
    PlaySong = True

    frmVideo.WindowsMediaPlayer1.settings.volume = 0

    'Analog ------------------------------------------------------------
'    'added by Andi 14-12-2020
'    Dim extension As String
'    extension = fso.GetExtensionName(PathLagu)
'    If extension = "mp4" Or extension = "mpg" Then
'      If lstPlaylist.selectedItem.SubItems(4) = "ML" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = True
'          tmrNonVocalMR.Enabled = False
'          vVocal = 2
'          vExtensi = True ' added by Andi 17-12-2020
'      ElseIf lstPlaylist.selectedItem.SubItems(4) = "MR" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = True
'          vVocal = 3
'          vExtensi = True ' added by Andi 17-12-2020
'      End If
'    Else
'      If lstPlaylist.selectedItem.SubItems(4) = "ML" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = True
'          tmrNonVocalMR.Enabled = False
'          vVocal = 2
'          vExtensi = False 'added by Andi 17-12-2020
'      ElseIf lstPlaylist.selectedItem.SubItems(4) = "MR" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = True
'          vVocal = 3
'          vExtensi = False 'added by Andi 17-12-2020
'      ElseIf lstPlaylist.selectedItem.SubItems(4) = "ST" Then
'          tmrVokal.Enabled = False
'          tmrNonVocalML.Enabled = False
'          tmrNonVocalMR.Enabled = False
'          vVocal = 1
'          vExtensi = False 'added by Andi 17-12-2020
'        End If
'      End If
'     'added by Andi 14-12-2020

        'added by Andi 14-12-2020
    Dim extension As String
    extension = fso.GetExtensionName(PathLagu)
    If extension = "mp4" Then
      If lstPlaylist.selectedItem.SubItems(4) = "ML" Then
          tmrVokal1.Enabled = False
          tmrNonVocalML1.Enabled = True
          tmrNonVocalMR1.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstPlaylist.selectedItem.SubItems(4) = "MR" Then
          tmrVokal1.Enabled = False
          tmrNonVocalML1.Enabled = False
          tmrNonVocalMR1.Enabled = True
          vVocal = 3
          vExtensi = True ' added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      End If
    ElseIf extension = "mpg" Then
      If lstAll.ListItems(selIndex).SubItems(5) = "ML" Then
          tmrVokal2.Enabled = False
          tmrNonVocalML2.Enabled = True
          tmrNonVocalMR2.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
          vMPG = True 'added by Andi 09-01-2021
      ElseIf lstAll.ListItems(selIndex).SubItems(5) = "MR" Then
          tmrVokal2.Enabled = False
          tmrNonVocalML2.Enabled = False
          tmrNonVocalMR2.Enabled = True
          vVocal = 3
          vExtensi = True 'added by Andi 17-12-2020
          vMPG = True 'added by Andi 09-01-2021
      End If
    Else
      'Original Code
      If lstPlaylist.selectedItem.SubItems(4) = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstPlaylist.selectedItem.SubItems(4) = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
      ElseIf lstPlaylist.selectedItem.SubItems(4) = "ST" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = False
          vVocal = 1
          vExtensi = False 'added by Andi 17-12-2020
          vMPG = False 'added by Andi 09-01-2021
        End If
      'Original Code
      End If
     'added by Andi 14-12-2020
     
    vVolvocal = 0

    frmVideo.WindowsMediaPlayer1.settings.volume = lstPlaylist.selectedItem.SubItems(5)

    txtPlaying.text = lstPlaylist.selectedItem.text + " - " + lstPlaylist.selectedItem.SubItems(1)

'    Set MyRsP = Nothing

Populer:
    '===POPULER SONG===
    Dim x As Double
    SQLP = "SELECT IDMUSIC, POPULER FROM masters WHERE IDMUSIC = " & lstPlaylist.selectedItem.SubItems(2)
    Set MyRsP = MyConn.Execute(SQLP)
    x = Val(MyRsP.Fields(1).value)
    Set MyRsP = Nothing
    SQLP = "UPDATE masters SET POPULER= " & Str(x + 1) & " Where IDMUSIC = " & lstPlaylist.selectedItem.SubItems(2)
    MyConn.Execute SQLP
End Sub

Private Sub PlayLst()

    On Error Resume Next

    Dim PathLagu As String
    Dim PathLaguUtama As String
    Dim PathLaguBackup As String
    Dim ECHO As ICMP_ECHO_REPLY
    Dim x As Double

    Dim MyRsP As MYSQL_RS
    Dim SQLP As String

    vTempo = 0
    vKey = 0

PlayUlang:

    If lstPlaylist.ListItems.Count = 0 Then
        GoTo lblEnd
    End If


    ScoreValid = True

    PathLagu = Replace$(lstPlaylist.ListItems.Item(1).SubItems(3), "/", "\")
    If AktifServerStatus = 1 Then
        PathLaguUtama = "\\" & vpbServerUtama & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerBackup & "\Data\" & PathLagu
    Else
        PathLaguUtama = "\\" & vpbServerBackup & "\Data\" & PathLagu
        PathLaguBackup = "\\" & vpbServerUtama & "\Data\" & PathLagu
    End If

    If FileExists(PathLaguUtama) Then
       frmVideo.Show
       frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
    Else
        Call Ping(vpbServerBackup, ECHO)
        If ECHO.status = 0 Then
            If FileExists(PathLaguBackup) Then
               frmVideo.Show
               frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
            Else
                '===POPULER SONG===
                SQLP = "SELECT IDMUSIC, POPULER FROM masters WHERE IDMUSIC = '" & lstPlaylist.ListItems.Item(1).SubItems(2) & "';"
                Set MyRsP = MyConn.Execute(SQLP)
                x = Val(MyRsP.Fields(1).value)
                Set MyRsP = Nothing
                SQLP = "UPDATE masters SET POPULER= " & Str(x + 1) & " Where IDMUSIC = '" & lstPlaylist.ListItems.Item(1).SubItems(2) & "';"
                MyConn.Execute SQLP
                If lstPlaylist.ListItems.Count > 0 Then
                        lErr = LockWindowUpdate(lstPlaylist.hWnd)
                        lstPlaylist.ListItems.Remove (1)
                        savePlayList
                        ClientRemotePlaylist
                        ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
                        lErr = LockWindowUpdate(0)
                End If
                If lstPlaylist.ListItems.Count = 0 Then
                    txtNextSong.text = "NO SONG"
                Else
                    txtNextSong.text = lstPlaylist.ListItems.Item(1)
                End If
                If lstPlaylist.ListItems.Count = 0 Then
                    Unload frmVideo
                    frmPromo.Show
                    frmTransparent.Show
                    frmRoom.Show
                End If
                If vVideo = 7 Then
                    Form2.Show
                    frmCamera.Show
                End If
                PlaySong = False
                GoTo PlayUlang
            End If
        Else
            '===POPULER SONG===
            SQLP = "SELECT IDMUSIC, POPULER FROM masters WHERE IDMUSIC = '" & lstPlaylist.ListItems.Item(1).SubItems(2) & "';"
            Set MyRsP = MyConn.Execute(SQLP)
            x = Val(MyRsP.Fields(1).value)
            Set MyRsP = Nothing
            SQLP = "UPDATE masters SET POPULER= " & Str(x + 1) & " Where IDMUSIC = '" & lstPlaylist.ListItems.Item(1).SubItems(2) & "';"
            MyConn.Execute SQLP
            If lstPlaylist.ListItems.Count > 0 Then
                lErr = LockWindowUpdate(lstPlaylist.hWnd)
                lstPlaylist.ListItems.Remove (1)
                savePlayList
                ClientRemotePlaylist
                ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
                lErr = LockWindowUpdate(0)
            End If
            If lstPlaylist.ListItems.Count = 0 Then
                txtNextSong.text = "NO SONG"
            Else
                txtNextSong.text = lstPlaylist.ListItems.Item(1)
            End If
            If lstPlaylist.ListItems.Count = 0 Then
                Unload frmVideo
                frmPromo.Show
                frmTransparent.Show
                frmRoom.Show
            End If
            If vVideo = 7 Then
                Form2.Show
                frmCamera.Show
            End If
            PlaySong = False
            GoTo PlayUlang
        End If
    End If
    If UkuranVideo <> 1 Then
        Form2.Show
    End If
    frmVideo.WindowsMediaPlayer1.Controls.play
    PlaySong = True

    frmVideo.WindowsMediaPlayer1.settings.volume = 0
    'Analog ------------------------------------------------------------
    
    'added by Andi 14-12-2020
    Dim extension As String
    extension = fso.GetExtensionName(PathLagu)
    If extension = "mp4" Or extension = "mpg" Then
      If lstPlaylist.ListItems.Item(1).SubItems(4) = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
      ElseIf lstPlaylist.ListItems.Item(1).SubItems(4) = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = True ' added by Andi 17-12-2020
      End If
    Else
      If lstPlaylist.ListItems.Item(1).SubItems(4) = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = False 'added by Andi 17-12-2020
      ElseIf lstPlaylist.ListItems.Item(1).SubItems(4) = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = False 'added by Andi 17-12-2020
      ElseIf lstPlaylist.ListItems.Item(1).SubItems(4) = "ST" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = False
          vVocal = 1
          vExtensi = False 'added by Andi 17-12-2020
      End If
    End If
    'added by Andi 14-12-2020
    
    vVolvocal = 0
    frmVideo.WindowsMediaPlayer1.settings.volume = lstPlaylist.ListItems.Item(1).SubItems(5)

    Set MyRsP = Nothing

    '===POPULER SONG===

    SQLP = "SELECT IDMUSIC, POPULER FROM masters WHERE IDMUSIC = " & lstPlaylist.ListItems.Item(1).SubItems(2)
    Set MyRsP = MyConn.Execute(SQLP)
    x = Val(MyRsP.Fields(1).value)
    Set MyRsP = Nothing
    SQLP = "UPDATE masters SET POPULER= " & Str(x + 1) & " Where IDMUSIC = " & lstPlaylist.ListItems.Item(1).SubItems(2)
    MyConn.Execute SQLP


lblEnd:

    If Err.Number <> 0 Then
      LogError Name, "PlayLst"
    End If
End Sub

Private Sub PlayLstUser()
On Error Resume Next

    vTempo = 0
    vKey = 0

    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    cmdFast.Enabled = True
    cmdSlow.Enabled = True
    cmdRepeat.Enabled = True
    cmdPause(0).Visible = True
    cmdPause(1).Visible = False

    If lstPlaylist.ListItems.Count = 0 Then
        Exit Sub
    End If

    Dim SQLP As String
    Dim MyRsP As MYSQL_RS
        SQLP = "SELECT PATH, ANALOG, VOL FROM masters Where IDMUSIC = '" & lstPlayUser.ListItems.Item(1).SubItems(2) & "';"
        Set MyRsP = MyConn.Execute(SQLP)

    ScoreValid = True
    frmVideo.Show
    If FileExists("P:\" & MyRsP.Fields(0).value) Then
        frmVideo.WindowsMediaPlayer1.URL = "P:\" & MyRsP.Fields(0).value
    Else
        frmVideo.WindowsMediaPlayer1.URL = "Q:\" & MyRsP.Fields(0).value
    End If
    frmVideo.WindowsMediaPlayer1.Controls.play
    PlaySong = True

    frmVideo.WindowsMediaPlayer1.settings.volume = 0

    'Analog ------------------------------------------------------------
    
    'added by Andi 14-12-2020
    Dim extension, PathLagu As String
    PathLagu = MyRsP.Fields(0).value
    extension = fso.GetExtensionName(PathLagu)
    If extension = "mp4" Or "mpg" Then
      If MyRsP.Fields(1).value = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = True ' added by Andi 17-12-2020
      ElseIf MyRsP.Fields(1).value = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = True ' added by Andi 17-12-2020
      End If
    Else
      If MyRsP.Fields(1).value = "ML" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = True
          tmrNonVocalMR.Enabled = False
          vVocal = 2
          vExtensi = False 'added by Andi 17-12-2020
      ElseIf MyRsP.Fields(1).value = "MR" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = True
          vVocal = 3
          vExtensi = False 'added by Andi 17-12-2020
      ElseIf MyRsP.Fields(1).value = "ST" Then
          tmrVokal.Enabled = False
          tmrNonVocalML.Enabled = False
          tmrNonVocalMR.Enabled = False
          vVocal = 1
          vExtensi = False 'added by Andi 17-12-2020
      End If
    End If
    'added by Andi 14-12-2020
    
    vVolvocal = 0
    frmVideo.WindowsMediaPlayer1.settings.volume = MyRsP.Fields(2).value

    Set MyRsP = Nothing
End Sub

Public Sub VideocapCD()

    On Error Resume Next

    Dim VideoFormatIndex As Integer
    Dim result As Integer
    Dim i As Integer

    vrekamstate = True
    btnRecStop.Visible = True
    flsChatNewMessage.Visible = False
    frmCamera.Show
    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.TVMute = True
    frmCamera.Visible = False

    frmCamera.VideoCap1.VideoStandard = vpbCameraStandar 'cbovideostand.ListIndex
    ControlCap ("Video Composite")
    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB24 (320x240)")
    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If
    frmCamera.VideoCap1.CaptureVideo = True
    frmCamera.VideoCap1.CaptureAudio = True
    frmCamera.VideoCap1.ShowPreview = True
    frmCamera.VideoCap1.UseVideoCompressor = False
    frmCamera.VideoCap1.UseAudioCompressor = False

    HurufDrive = "C"
    For i = 0 To 4
        If UCase(Left$(Trim$(Drive1.List(i)), 1)) = "D" Then
            HurufDrive = "D"
        End If
    Next i

    If Not (FileExists(HurufDrive & ":\" & modProject.brandName)) Then
        MkDir HurufDrive & ":\" & modProject.brandName
    End If

    i = 1
    Do Until Not (FileExists(HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "CD" & i & ".AVI"))
        i = i + 1
    Loop
    picfile = HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "CD" & i & ".AVI"

    frmCamera.VideoCap1.CaptureMode = True
    frmCamera.VideoCap1.CaptureFileName = picfile

    result = frmCamera.VideoCap1.Start
    frmCamera.Refresh
    DoEvents

'    frmCamera.Show
'    frmCamera.Visible = False

    If result = -1 Then
        'MsgBox "Capture Failure,Video, Audio Compressor not correct or capture file opening"
        Exit Sub
    End If

    If result = -2 Then
        'MsgBox "Capture file not found"
        Exit Sub
    End If

    tmrRecord.Enabled = True
End Sub

Public Sub VideocapVCD()
    On Error Resume Next

    Dim VideoFormatIndex As Integer
    Dim result As Integer

    vrekamstate = True

    frmCamera.Show
    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.TVMute = True
    frmCamera.Width = 7680
    frmCamera.Left = Screen.Width - frmCamera.Width
    frmCamera.Top = 1050
    frmCamera.Height = 5700
    frmCamera.VideoCap1.Top = 0
    frmCamera.VideoCap1.Left = 75
    frmCamera.VideoCap1.Height = 5610
    frmCamera.VideoCap1.Width = 7530

    btnRecStop.Visible = True
    flsChatNewMessage.Visible = False

    frmCamera.VideoCap1.VideoStandard = vpbCameraStandar

    ControlCap ("Video Composite")

    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB24 (320x240)")

    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If

    frmCamera.VideoCap1.CaptureVideo = True
    frmCamera.VideoCap1.CaptureAudio = True

    frmCamera.VideoCap1.ShowPreview = True
    frmCamera.VideoCap1.UseVideoCompressor = False
    frmCamera.VideoCap1.UseAudioCompressor = False

    Dim i As Integer
    'Dim HurufDrive As String

    HurufDrive = "C"
    For i = 0 To 4
        If UCase(Left$(Trim$(Drive1.List(i)), 1)) = "D" Then
            HurufDrive = "D"
        End If
    Next i

    If Not (FileExists(HurufDrive & ":\" & modProject.brandName)) Then
        MkDir HurufDrive & ":\" & modProject.brandName
    End If

    i = 1
    Do Until Not (FileExists(HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "VCD" & i & ".AVI"))
        i = i + 1
    Loop
    picfile = HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "VCD" & i & ".AVI"

    frmCamera.VideoCap1.CaptureMode = True
    frmCamera.VideoCap1.CaptureFileName = picfile

    result = frmCamera.VideoCap1.Start
    frmCamera.Show

    If result = -1 Then
        'MsgBox "Capture Failure,Video, Audio Compressor not correct or capture file opening"
        Exit Sub
    End If

    If result = -2 Then
        'MsgBox "Capture file not found"
        Exit Sub
    End If

    tmrRecord.Enabled = True
End Sub

Public Sub VideocapHP()
    On Error Resume Next

    Dim VideoFormatIndex As Integer
    Dim result As Integer

    vrekamstate = True
    frmCamera.Show
    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.TVMute = True
    frmCamera.Width = 7680
    frmCamera.Left = Screen.Width - frmCamera.Width
    frmCamera.Top = 1050
    frmCamera.Height = 5700
    frmCamera.VideoCap1.Top = 0
    frmCamera.VideoCap1.Left = 75
    frmCamera.VideoCap1.Height = 5610
    frmCamera.VideoCap1.Width = 7530

    btnRecStop.Visible = True
    flsChatNewMessage.Visible = False

    frmCamera.VideoCap1.VideoStandard = vpbCameraStandar

    ControlCap ("Video Composite")

    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB24 (320x240)")

    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If

    frmCamera.VideoCap1.CaptureVideo = True
    frmCamera.VideoCap1.CaptureAudio = True

    frmCamera.VideoCap1.ShowPreview = True
    frmCamera.VideoCap1.UseVideoCompressor = False
    frmCamera.VideoCap1.UseAudioCompressor = False

    Dim i As Integer

    HurufDrive = "C"
    For i = 0 To 4
        If UCase(Left$(Trim$(Drive1.List(i)), 1)) = "D" Then
            HurufDrive = "D"
        End If
    Next i

    If Not (FileExists(HurufDrive & ":\" & modProject.brandName)) Then
        MkDir HurufDrive & ":\" & modProject.brandName
    End If

    i = 1
    Do Until Not (FileExists(HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "HP" & i & ".AVI"))
        i = i + 1
    Loop
    picfile = HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "HP" & i & ".AVI"

    frmCamera.VideoCap1.CaptureMode = True
    frmCamera.VideoCap1.CaptureFileName = picfile

    result = frmCamera.VideoCap1.Start

    If result = -1 Then
        'MsgBox "Capture Failure,Video, Audio Compressor not correct or capture file opening"
        Exit Sub
    End If

    If result = -2 Then
        'MsgBox "Capture file not found"
        Exit Sub
    End If

    tmrRecord.Enabled = True
End Sub

Public Sub VideocapDVD()
    On Error Resume Next

    Dim VideoFormatIndex As Integer
    Dim result As Integer

    vrekamstate = True

    frmCamera.Show
    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.TVMute = True

    frmCamera.Width = 7680
    frmCamera.Left = Screen.Width - frmCamera.Width
    frmCamera.Top = 1050
    frmCamera.Height = 5700
    frmCamera.VideoCap1.Top = 0
    frmCamera.VideoCap1.Left = 75
    frmCamera.VideoCap1.Height = 5610
    frmCamera.VideoCap1.Width = 7530


    btnRecStop.Visible = True
    flsChatNewMessage.Visible = False

    frmCamera.VideoCap1.VideoStandard = vpbCameraStandar
    ControlCap ("Video Composite")
    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB24 (640x480)")

    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If

    frmCamera.VideoCap1.CaptureVideo = True
    frmCamera.VideoCap1.CaptureAudio = True

    frmCamera.VideoCap1.ShowPreview = True
    frmCamera.VideoCap1.UseVideoCompressor = False
    frmCamera.VideoCap1.UseAudioCompressor = False

    Dim i As Integer

    HurufDrive = "C"
    For i = 0 To 4
        If UCase(Left$(Trim$(Drive1.List(i)), 1)) = "D" Then
            HurufDrive = "D"
        End If
    Next i

    If Not (FileExists(HurufDrive & ":\" & modProject.brandName)) Then
        MkDir HurufDrive & ":\" & modProject.brandName
    End If

    i = 1
    Do Until Not (FileExists(HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "DVD" & i & ".AVI"))
        i = i + 1
    Loop
    picfile = HurufDrive & ":\" & modProject.brandName & "\" & modProject.brandName & "DVD" & i & ".AVI"

    frmCamera.VideoCap1.CaptureMode = True
    frmCamera.VideoCap1.CaptureFileName = picfile

    result = frmCamera.VideoCap1.Start
    If result = -1 Then
        'MsgBox "Capture Failure,Video, Audio Compressor not correct or capture file opening"
        Exit Sub
    End If

    If result = -2 Then
        'MsgBox "Capture file not found"
        Exit Sub
    End If

    tmrRecord.Enabled = True
End Sub

Private Sub VideocapPicture()
    On Error Resume Next
If Not (FileExists("D:\" & modProject.brandName)) Then
    MkDir "D:\" & modProject.brandName
End If

flsChatNewMessage.Visible = False

Dim i As Integer
i = 1
Do Until Not (FileExists("D:\" & modProject.brandName & "\" & modProject.brandName & "PIC" & i & ".jpg"))
    i = i + 1
Loop
    picfile = "D:\" & modProject.brandName & "\" & modProject.brandName & "PIC" & i & ".jpg"

   If frmCamera.VideoCap1.SnapShotJPEG(picfile, 100) Then
     frmFotoConfirm.Show
     frmFotoConfirm.Image1.Picture = LoadPicture(picfile)
     frmFotoConfirm.Image1.Stretch = True
   End If
End Sub

Sub ControlCap(vasal As String)
    On Error Resume Next

    Dim strVideoInput As String
    Dim strVideoFormat As String
    Dim strAudioDevice As String
    Dim strVideoCompressor As String
    Dim strAudioCompressor As String
    Dim strDevice As String
    Dim deviceIndex As Integer
    Dim videoinputindex As Integer
    Dim VideoFormatIndex As Integer
    Dim AudioIndex As Integer

    strDevice = cbodevice.List(cbodevice.ListIndex)
    deviceIndex = frmCamera.VideoCap1.Devices.FindDevice(strDevice)
    If deviceIndex <> -1 Then
        frmCamera.VideoCap1.Device = deviceIndex
    End If

    'strVideoInput = cboVideoInput.List(cboVideoInput.ListIndex)
    videoinputindex = frmCamera.VideoCap1.VideoInputs.FindVideoInput(vasal)

    If videoinputindex <> -1 Then
        frmCamera.VideoCap1.VideoInput = videoinputindex
    End If

    'Video Format
    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB24 (320x240)")
    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If
    'Audio Device
    strAudioDevice = cboaudiodevice.List(cboaudiodevice.ListIndex)
    AudioIndex = frmCamera.VideoCap1.AudioDevices.FindDevice(cboaudiodevice.List(cboaudiodevice.ListIndex))
    If AudioIndex <> -1 Then
        frmCamera.VideoCap1.AudioDevice = 0
    End If

'Video Compressor

'strVideoCompressor = "MSScreen Encoder DM0"
 '   VideoCompressorIndex = frmCamera.VideoCap1.VideoCompressors.FindVideoCompressor("MSScreen Encoder DM0")


   ' If VideoCompressorIndex <> -1 Then
   '     frmCamera.VideoCap1.VideoCompressor = VideoCompressorIndex
   ' End If


    'strAudioCompressor = "WM Speech Encoder DM0"
    'AudioCompressorIndex = frmCamera.VideoCap1.AudioCompressors.FindAudioCompressor("WM Speech Encoder DM0")

    'If AudioCompressorIndex <> -1 Then
    '    frmCamera.VideoCap1.AudioCompressor = AudioCompressorIndex
    'End If
End Sub

' check unused
'Public Sub pesan()
'    On Error Resume Next
'Dim Data As String
'
'        If frmRoom.Socket(iSockets).State = sckConnected Then
'            Data = "R13"
'            frmRoom.Socket(iSockets).SendData Data
'        End If
'        frmRoom.Socket(0).Close
'End Sub

Sub prcPanggil()
    On Error Resume Next
    Socket(iSockets).Close

    Dim Sql As String
    Dim myrs As MYSQL_RS
    Sql = "select panggil from room where roomname= '" & txtCompName & "';"
    Set myrs = MyConn.Execute(Sql)

    If myrs.FieldCount > 0 Then
        Socket(iSockets).RemoteHost = myrs.Fields(0).value
        Socket(iSockets).RemotePort = "10123"
        Socket(iSockets).Connect
    End If

    frmConfirmasi.vpengirim = 5
    frmConfirmasi.Text1 = ""
    frmConfirmasi.Text2 = "CALL WAITER ?"
    frmConfirmasi.text3 = ""
    frmConfirmasi.Show
End Sub

Public Sub MidiAktif()

    On Error Resume Next

    Dim VideoFormatIndex As Integer

    '-------aktifkan camera--------
    frmCamera.Show
    frmCamera.VideoCap1.Start
    frmCamera.WindowState = 2
    frmCamera.VideoCap1.Left = 0
    frmCamera.VideoCap1.Width = 15360 'frmCamera.Width
    frmCamera.VideoCap1.Height = 11520 'frmCamera.Height

    frmCamera.VideoCap1.Visible = True

    frmCamera.VideoCap1.VideoStandard = getIniSetting("Midi", "VideoStandard", "2")

    ControlCap ("S-Video")

    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat(getIniSetting("Midi", "VideoFormat", "RGB24 (640x480)"))
    If VideoFormatIndex <> -1 Then
        frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If

  '  frmCamera.VideoCap1.AudioInputPin = 5
    frmCamera.VideoCap1.CaptureMode = False
    frmCamera.VideoCap1.Start
End Sub

Sub midilist()
    On Error Resume Next

    Dim chm As ColumnHeader
    Dim LV As ListItem

    If lstMidi.ListItems.Count > 0 Then
        Exit Sub
    End If

    lstMidi.Visible = True
    lstMidi.ListItems.Clear
    lstMidi.ColumnHeaders.Clear
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Sql = "SELECT IDMUSIC,TITLE, ARTIST FROM mastermidi ORDER BY IDMUSIC ;"
    Set myrs = MyConn.Execute(Sql)

    'Memberi judul list data
    Set chm = lstMidi.ColumnHeaders.add(, , , 1500)
    Set chm = lstMidi.ColumnHeaders.add(, , , 6000)
    Set chm = lstMidi.ColumnHeaders.add(, , , 6000)
    Set chm = lstMidi.ColumnHeaders.add(, , , 5)

    'Do Until myrs.EOF
        'Isi list data
        Set LV = lstMidi.ListItems.add(, , (myrs.Fields(0).value))
        LV.SubItems(1) = myrs.Fields(1).value
        LV.SubItems(2) = myrs.Fields(2).value
        myrs.MoveNext
    'Loop
    Set myrs = Nothing
End Sub

Sub movielist()
    On Error Resume Next

    Dim LV As ListItem

    lstMovie.Visible = True
    lstMovie.ListItems.Clear
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Sql = "SELECT ID, TITLE FROM FILM ORDER BY TITLE  LIMIT 0, 9 ;"
    Set myrs = MyConn.Execute(Sql)
    vLstAllRow = myrs.recordCount

    Do Until myrs.EOF
        'Isi list data
        Set LV = lstMovie.ListItems.add(, , (myrs.Fields(0).value))
        LV.SubItems(1) = myrs.Fields(1).value
        LV.SubItems(2) = myrs.Fields(0).value
        myrs.MoveNext
    Loop

    Set myrs = Nothing
End Sub

Public Sub moviecari()
    On Error Resume Next

    Dim cariapa As String
    Dim sqlm As String
    Dim myrs As MYSQL_RS
    Dim textbersih As String
    Dim LV As ListItem

    Select Case vpointer
    Case 2
        cariapa = "ARTIS" & " LIKE '%"
    Case Else
        cariapa = "TITLE LIKE '"
    End Select
    textbersih = mysql_escape_string(txtSearch.text)
    sqlm = "SELECT ID, TITLE FROM FILM WHERE " & cariapa & textbersih & "%' "
              If vMovieKategori = 1 Then sqlm = sqlm & " AND drama  = 1 "
              If vMovieKategori = 2 Then sqlm = sqlm & " AND komedi = 1 "
              If vMovieKategori = 3 Then sqlm = sqlm & " AND action = 1 "
              If vMovieKategori = 4 Then sqlm = sqlm & " AND horor  = 1 "
              If vMovieKategori = 5 Then sqlm = sqlm & " AND kartun = 1 "
    If cbokategori.ListIndex > 0 Then
        sqlm = sqlm & " AND negara = " & cbokategori.ItemData(cbokategori.ListIndex)
    End If
    sqlm = sqlm & " ORDER BY TITLE LIMIT 0,9 ;"
    Set myrs = MyConn.Execute(sqlm)
    vLstAllRow = myrs.recordCount
        lstMovie.ListItems.Clear
        Do Until myrs.EOF
            'Isi list data
            Set LV = lstMovie.ListItems.add(, , (myrs.Fields(0).value))
            LV.SubItems(1) = myrs.Fields(1).value
            LV.SubItems(2) = myrs.Fields(0).value
            myrs.MoveNext
        Loop
    Set myrs = Nothing

    If lstMovie.ListItems.Count > 0 Then
        lstMovie.selectedItem.Selected = False
        Set lstMovie.DropHighlight = lstMovie.selectedItem
        lstMovie_Click
    End If
End Sub

Sub midiKategori()
    On Error Resume Next
    Dim chm As ColumnHeader
    Dim LV As ListItem

    If cbokategori.ListIndex = 0 Then
        lstMidiMusic.Visible = False
        lstMidi.Visible = True
    Else
        lstMidiMusic.ListItems.Clear
        lstMidiMusic.ColumnHeaders.Clear

        If cbokategori.ListIndex > 3 Then
            lstMidi.Visible = False
            lstMidiMusic.Visible = True
            txtSearch.SetFocus
            Exit Sub
        End If
        Dim Sql As String
        Dim myrs As MYSQL_RS
        Sql = "SELECT IDMUSIC,TITLE, ARTIST FROM mastermidi Where TYPE = " & cbokategori.ItemData(cbokategori.ListIndex) & " ORDER BY IDMUSIC ;"
        Set myrs = MyConn.Execute(Sql)

        'Memberi judul list data
        Set chm = lstMidiMusic.ColumnHeaders.add(, , , 1400)
        Set chm = lstMidiMusic.ColumnHeaders.add(, , , 4500)
        Set chm = lstMidiMusic.ColumnHeaders.add(, , , 3900)
        Set chm = lstMidiMusic.ColumnHeaders.add(, , , 5)

        Do Until myrs.EOF
            'Isi list data
            Set LV = lstMidiMusic.ListItems.add(, , (myrs.Fields(0).value))
            LV.SubItems(1) = myrs.Fields(1).value
            LV.SubItems(2) = myrs.Fields(2).value
            myrs.MoveNext
        Loop
        Set myrs = Nothing

        lstMidi.Visible = False
        lstMidiMusic.Visible = True
    End If
End Sub

Sub setvolumenaik()
    On Error Resume Next

    If Val(txtVol.text) + 5 >= 100 Then
        txtVol.text = "100"
        setAudioEndPointVolumeMasterVolumeLevelPercent 100
    Else
        setAudioEndPointVolumeMasterVolumeLevelPercent CLng(txtVol.text) + 5
        txtVol.text = Val(txtVol.text) + 5
    End If

    ClientRemoteVolume
End Sub

Sub setvolumeturun()
    On Error Resume Next
    If Val(txtVol.text) - 5 <= 0 Then
        txtVol.text = "0"
        setAudioEndPointVolumeMasterVolumeLevelPercent 0
    Else
        setAudioEndPointVolumeMasterVolumeLevelPercent CLng(txtVol.text) - 5
        txtVol.text = Val(txtVol.text) - 5
    End If

    ClientRemoteVolume
End Sub


Sub tutup()
    On Error Resume Next
    'tutup semua untuk perpindahan applikasi
        '---frameCombo
        lstChat.Visible = False
        lstAll.Visible = False
        lstPlaylist.Visible = False
        lstMidi.Visible = False
        lstTV.Visible = False
        lstMovie.Visible = False

        picKategori.Visible = False
        cbokategori.Visible = False
        Picture3.Visible = False
        txtVol.Visible = False
        txtChat.Visible = False
        txtChat.text = ""

        '---MOVIE
        txtArtis.Visible = False
        txtSinopsis.Visible = False
End Sub

Sub chatSend()
    On Error Resume Next

    Dim myrs As MYSQL_RS
    Dim Sql As String
    Dim LV As ListItem

    If txtChat.text = "" Then
        Exit Sub
    End If
    Set LV = lstChat.ListItems.add(, , txtUser.text)
        Dim Data As String
        Data = txtChat.text
        Sql = "INSERT INTO chat (RoomAsal, RoomTujuan, Pesan) VALUES (" & Str$(txtCompName.Tag) & _
              ", " & Str$(txtChatAktif.Tag) & ", '" & mysql_escape_string(Data) & "')"
        Set myrs = MyConn.Execute(Sql)
        Set myrs = Nothing
    LV.SubItems(1) = Data

    Sql = "UPDATE room SET CHAT= 1 WHERE IDROOM= " & Str$(txtChatAktif.Tag) & ";"
    Set myrs = MyConn.Execute(Sql)
    Set myrs = Nothing

    txtChat = ""
    chatakhir_Listview lstChat, " ", 1
    If lstChat.ListItems.Count > 0 Then
        lstChat.selectedItem.Selected = False
    End If
End Sub

Public Sub chatbuka()
    On Error Resume Next
    If vVideo = 3 Then
        frmCountry.Show
        frmCountry.lstCountry.SetFocus
        Exit Sub
    End If

    If vVideo = 0 Then
        Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
        frmTransparent.Hide
    End If

    txtUser.SetFocus
    vVideo = 3
    cbokategori.ListIndex = 0

    'Buka Transparent

    '--------SKIN COY----------
    lokasi = App.Path
    Skin1.LoadSkin lokasi + "\skin\sknchat.skn"
    Skin1.ApplySkinByName hWnd, "sknchat"
    flsLogo.Movie = lokasi + "\picture\chat\logo.swf"
    flsTitle.Movie = lokasi + "\picture\anim\message"

    txtSearch.Visible = False
    txtCategory.Visible = False
    tutup

    txtPlaying.Visible = False
    picKeyTempo.Visible = False
    tmrPicKeyTempo.Enabled = False
    flsTitle.Visible = False
    txtSearch.Visible = False
    flsMovieCategory.Visible = False
    flsChatNewMessage.Visible = False
    txtRemoteCode.Visible = False

    '-----TAMPILKAN CHAT
    flsLogo.Top = 0
    flsLogo.Width = 15360
    flsLogo.Height = 11520
    flsLogo.Left = 0
    flsLogo.Visible = True

    txtChatAktif.Top = 2400
    txtChatAktif.Width = 3495
    txtChatAktif.Left = 6240
    txtChatAktif.Height = 345
    txtChatAktif.Visible = True

    lstChat.Left = 3360
    lstChat.Height = 5175
    lstChat.Top = 3000
    lstChat.Width = 11385

    txtChat.Font.Size = 22
    txtChat.Top = 10320
    txtChat.Left = 7560
    txtChat.Height = 525
    txtChat.Width = 5040

    txtUser.Top = 520
    txtUser.Left = 13440
    txtUser.Height = 270
    txtUser.Width = 1575
    txtUser.BackColor = &HFFFFFF
    txtUser.ForeColor = &H580148
    txtUser.Visible = True

    txtTime.Top = 870
    txtTime.Left = 13440
    txtTime.Height = 270
    txtTime.Width = 1575
    txtTime.BackColor = &HFFFFFF
    txtTime.ForeColor = &H580148
    txtTime.Visible = True

    txtVol.Top = 10200
    txtVol.Left = 14040
    txtVol.Height = 645
    txtVol.Width = 1020
    txtVol.Font.Size = 27
    txtVol.BackColor = &HFFFFFF
    txtVol.ForeColor = &H580148
    txtVol.Visible = True

    lstChat.Visible = True
    txtChat.Visible = True

    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim LV As ListItem

    Sql = " SELECT chat.pesan, ra.userroom as asal, rt.userroom as tujuan FROM chat inner join room ra on chat.RoomAsal = ra.idroom " & _
          " inner join room rt on chat.roomtujuan = rt.idroom " & _
          " Where  chat.roomtujuan = " & Str$(txtCompName.Tag) & _
          " order by chat.id desc limit 0,10"
    Set myrs = MyConn.Execute(Sql)
    myrs.MoveLast
    Do Until myrs.BOF
        If (myrs.Fields(1).value = txtUser.text) Then
            Set LV = lstChat.ListItems.add(, , (myrs.Fields(1).value))
            LV.SubItems(1) = myrs.Fields(0).value
        Else
            Set LV = lstChat.ListItems.add(, , (myrs.Fields(1).value))
            LV.SubItems(1) = myrs.Fields(0).value
        End If
        myrs.MovePrevious
    Loop

    myrs.CloseRecordset
    Sql = "delete from chat where roomtujuan = '" & Str$(txtCompName.Tag) & "';"
    Set myrs = MyConn.Execute(Sql)
    Set myrs = Nothing
    If lstChat.ListItems.Count > 0 Then
        lstChat.selectedItem.Selected = False
    End If
    txtChat.SetFocus

    chatbuka
End Sub

Public Sub hotnon()
    On Error Resume Next
    HotKeyDeactivate Me.hWnd
    SetWindowLongA Me.hWnd, GWL_WNDPROC, oldProc
End Sub

Function TimeSelisih(Tjalan As Date) As String
    On Error Resume Next
    Dim selisihwaktu As Date
    selisihwaktu = Tjalan - vWaktumulai
    TimeSelisih = right("00" & hour(selisihwaktu), 2) & ":" & right("00" & minute(Now - vWaktumulai), 2) & ":" & right("00" & second(Now - vWaktumulai), 2)
End Function

Sub loadplaylist()

    On Error Resume Next

    Dim LV As ListItem
    Dim Sql As String
    Dim myrs As MYSQL_RS

    If txtLogin = "" Then
        Sql = "select mcd.title, mcd.singer, pl.idmusic, mcd.PATH, mcd.ANALOG, mcd.VOL from playlist pl inner join masters mcd " & _
              "on pl.idmusic = mcd.idmusic " & _
              "WHERE ROOM = '" & txtCompName.text & "';"
    Else
        Sql = "select mcd.title, mcd.singer, pl.idmusic, mcd.PATH, mcd.ANALOG, mcd.VOL from playlist pl inner join masters mcd " & _
              "on pl.idmusic = mcd.idmusic " & _
              "WHERE pl.userid = '" & txtLogin & "';"
    End If

    Set myrs = MyConn.Execute(Sql)

    DoEvents


    lErr = LockWindowUpdate(lstPlaylist.hWnd)
    Do Until myrs.EOF
        Set LV = frmRoom.lstPlaylist.ListItems.add(, , myrs.Fields(0).value)
        LV.SubItems(1) = myrs.Fields(1).value
        LV.SubItems(2) = myrs.Fields(2).value
        LV.SubItems(3) = myrs.Fields(3).value
        LV.SubItems(4) = myrs.Fields(4).value
        LV.SubItems(5) = myrs.Fields(5).value
        myrs.MoveNext
    Loop
    lErr = LockWindowUpdate(0)

    Sql = "UPDATE room SET PLAYLIST = 0 WHERE IDROOM= '" & txtCompName.Tag & "';"
    Set myrs = MyConn.Execute(Sql)


    If Err.Number <> 0 Then
      LogError Name, "loadplaylist"
    End If
End Sub

Function chatakhir_Listview(LV As ListView, SearchText As String, Column As Long, Optional SelectRow As Boolean = True, Optional MakeVisible As Boolean = True) As Long
    On Error Resume Next
Dim SearchLength As Long
Dim CurrentRow As Long
Dim result As Long

    result = LV.ListItems.Count   'LV.SelectedItem.Index + 1
    If result > LV.ListItems.Count Then
    result = LV.ListItems.Count
    End If

    If result > -1 Then
        If SelectRow Then LV.ListItems(result).Selected = True
        If MakeVisible Then LV.ListItems(result).EnsureVisible
    End If
End Function

Public Sub setvocal()
  'tentukan vocal/nonvocal 1=st 2=ml 3=mr
  'current audio language index 1=stereo<music+vocal>,2=karaoke<music only>,3=mono<vocal only>
  On Error Resume Next
  If (vExtensi) Then
    If (vVocal = 2) Then
      If vVideoAktif Then
        frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 2
      End If
        vVocal = 21
        ScoreValid = False
    ElseIf (vVocal = 21) Then
      If vVideoAktif Then
        frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 1
      End If
      vVocal = 2
    End If
  Else
        If (vVocal = 2) Then
        If vVideoAktif Then
                frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 3
        End If
        vVocal = 21
        ScoreValid = False
    ElseIf (vVocal = 21) Then
        If vVideoAktif Then
                frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 2
        End If
        vVocal = 2
    ElseIf (vVocal = 3) Then
        If vVideoAktif Then
            frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 2
        End If
        vVocal = 31
        ScoreValid = False
    ElseIf (vVocal = 31) Then
        If vVideoAktif Then
            frmVideo.WindowsMediaPlayer1.Controls.currentAudioLanguageIndex = 3
        End If
        vVocal = 3
    End If
  End If


'Err1:
End Sub

Public Sub screennormal()
    On Error Resume Next
        caretLebar = 7
        caretTinggi = 50

'        frmTransparent.Show
'        frmRoom.Show
        frmRoom.Height = 11520
        flsLogo.Visible = True

        lokasi = App.Path + "\picture\normalscreen\"
'        picKategori.Picture = LoadPicture(lokasi + "kategori.jpg")

        flsTitle.Movie = App.Path + "\picture\anim\title"
        flsLogo.Movie = App.Path + "\picture\normalscreen\logo"
        flsLogo.Top = 0
        flsLogo.Width = 1875
        flsLogo.Height = 1240
        flsLogo.Left = 0

        flsMovieCategory.Visible = False

        Picture3.Visible = False
        cbokategori.Visible = True
        picKategori.Visible = False
        txtCategory.Visible = True

        flsTitle.Visible = True
        flsTitle.Top = 90
        flsTitle.Width = 3015
        flsTitle.Height = 1050
        flsTitle.Left = 1920

        txtRemoteCode.Visible = modProject.frmRoomRemoteCode

        Picture3.Top = 980
        Picture3.Left = 4920
        Picture3.Height = 375
        Picture3.Width = 7215

        picKategori.Top = 410
        picKategori.Left = 5880
        picKategori.Height = 400
        picKategori.Width = 3255

        cbokategori.Left = 4680
        cbokategori.Width = 3855
        cbokategori.Top = Screen.Height + 8160
        cbokategori.Font.Size = 17
        cbokategori.BackColor = &HFFFFFF
        cbokategori.ForeColor = &H22A6B

        txtCategory.Top = 1260
        txtCategory.Left = 6240
        txtCategory.Width = 2895
        txtCategory.Font.Size = 16
        txtCategory.BackColor = &H244C&
        txtCategory.ForeColor = &HFFFFFF

        lstAll.Font.Size = 20
        lstAll.Top = 1920
        lstAll.Left = 2280
        lstAll.Height = 4695
        lstAll.Width = 11625

        lstPlaylist.Top = lstAll.Top
        lstPlaylist.Left = lstAll.Left
        lstPlaylist.Height = lstAll.Height
        lstPlaylist.Width = lstAll.Width - 255
        lstPlaylist.Font.Bold = lstAll.Font.Bold
        lstPlaylist.Font.Name = lstAll.Font.Name
        lstPlaylist.Font.Size = lstAll.Font.Size

        txtVol.Top = 6120
        txtVol.Left = 14160
        txtVol.Height = 645
        txtVol.Width = 900
        txtVol.Font.Size = 27
        txtVol.BackColor = &HFFFFFF
        txtVol.ForeColor = &H955001

        txtUser.Top = 520
        txtUser.Left = 6960
        txtUser.Height = 0
        txtUser.Width = 0
        txtUser.BackColor = &HFFFFFF
        txtUser.ForeColor = &H22A6B

        txtTime.Top = 520
        txtTime.Left = 6960
        txtTime.Height = 0
        txtTime.Width = 0
        txtTime.BackColor = &HFFFFFF
        txtTime.ForeColor = &H22A6B

        txtUser.Visible = True
        txtTime.Visible = True

        lblRecording.Top = 360
        lblRecording.Left = 12960
        lblRecording.Height = 585
        lblRecording.Width = 2280

        btnRecStop.Top = lblRecording.Top
        btnRecStop.Left = lblRecording.Left
        btnRecStop.Height = lblRecording.Height
        btnRecStop.Width = lblRecording.Width

        txtSearch.Font.Size = 28
        txtSearch.Top = 380
        txtSearch.Left = 5040
        txtSearch.Height = 680
        txtSearch.Width = 6615
        txtSearch.BackColor = &HFFFFFF
        txtSearch.ForeColor = &H955001

        lstAll.Visible = True
End Sub

Sub konekServer1()

    On Error Resume Next

    Dim ECHO As ICMP_ECHO_REPLY
    Dim remoteName As String
    Dim USER As String

    vpbServer = vpbServerUtama

    Call Ping(vpbServerUtama, ECHO)
    If ECHO.status = 0 Then
        MyConn.OpenConnection vpbServerUtama, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
        DoEvents
        If MyConn.Error.Number = 0 Then

            Dim myrs As MyVbQL.MYSQL_RS
            Set myrs = MyConn.Execute("select now();")
            Dim dateNow As Date
            dateNow = myrs.Fields(0).value
            setTime year(dateNow), month(dateNow), day(dateNow), hour(dateNow), minute(dateNow), second(dateNow)
            myrs.CloseRecordset
            DoEvents

            KonekServer

            remoteName = "\\" & vpbServerUtama
            USER = "Administrator"
            addNetworkConnection remoteName, Len(remoteName), USER, Len(USER), vpbServerKeyWindows, Len(vpbServerKeyWindows)
            DoEvents

            If vpbServerBackup <> vpbServerUtama Then
              Call Ping(vpbServerBackup, ECHO)
              If ECHO.status = 0 Then
                  remoteName = "\\" & vpbServerBackup
                  addNetworkConnection remoteName, Len(remoteName), USER, Len(USER), vpbServerKeyWindows, Len(vpbServerKeyWindows)
                  DoEvents
                  MyConnBackup.OpenConnection vpbServerBackup, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
                  DoEvents
              End If
            End If

            AktifServerStatus = 1
            tmrAktif.Enabled = True
        Else
            konekServer2
        End If
    Else
        konekServer2
    End If
End Sub

Sub konekServer2()

    On Error Resume Next

    Dim remoteName As String
    Dim USER As String
    Dim ECHO As ICMP_ECHO_REPLY

    vpbServer = vpbServerBackup

    Call Ping(vpbServerBackup, ECHO)
    If ECHO.status = 0 Then
        MyConn.OpenConnection vpbServerBackup, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
        DoEvents
        If MyConn.Error.Number = 0 Then

            Dim myrs As MyVbQL.MYSQL_RS
            Set myrs = MyConn.Execute("select now();")
            Dim dateNow As Date
            dateNow = myrs.Fields(0).value
            setTime year(dateNow), month(dateNow), day(dateNow), hour(dateNow), minute(dateNow), second(dateNow)
            myrs.CloseRecordset
            DoEvents

            KonekServer

            remoteName = "\\" & vpbServerBackup
            USER = "Administrator"
            addNetworkConnection remoteName, Len(remoteName), USER, Len(USER), vpbServerKeyWindows, Len(vpbServerKeyWindows)
            DoEvents

            If vpbServerUtama <> vpbServerBackup Then
              Call Ping(vpbServerUtama, ECHO)
              If ECHO.status = 0 Then
                  remoteName = "\\" & vpbServerUtama
                  addNetworkConnection remoteName, Len(remoteName), USER, Len(USER), vpbServerKeyWindows, Len(vpbServerKeyWindows)
                  DoEvents
                  MyConnBackup.OpenConnection vpbServerUtama, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
                  DoEvents
              End If
            End If

            AktifServerStatus = 2
            tmrAktif.Enabled = True
        Else
            konekServer1
        End If
    Else
        konekServer1
    End If
End Sub


Sub waktuhabis()
    On Error Resume Next
        Dim Sql As String
        Dim myrs As MYSQL_RS
        Sql = "SELECT APA from room where ROOMNAME ='" & txtCompName & "'"
        Set myrs = MyConn.Execute(Sql)

        If (myrs.Fields(0).value = "tutup") Then
            frmRoom.vHabisWaktu = True
        End If
End Sub

Public Sub moviestart()
    On Error Resume Next

    Dim vMuteTemp As Boolean

    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    If vpbProsesEksekusi Then
        Exit Sub
    End If
    vpbProsesEksekusi = True


    If vpbfrmCamera Then
        If (frmCamera.Width = Screen.Width) Or (vVideo = 1) Then
            setAudioEndPointVolumeMasterVolumeLevelPercent 0
            vMuteTemp = True
        End If
    End If

    picKeyTempo.Visible = False
    tmrPicKeyTempo.Enabled = False
    txtPlaying.Visible = True
    txtSearch.text = ""
    lstMovie.Visible = True
    If lstMovie.ListItems.Count = 0 Then
        movielist
    End If
    flsLogo.Visible = False

    txtChatAktif.Visible = False
    flsChatNewMessage.Visible = False
    txtSearch.Visible = True
    txtSearch.SetFocus

    tutup
    vVideo = 5
    '--------SKIN COY----------
    Skin1.RemoveSkin
    'Aktifkan transparent color
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY

    frmTransparent.LoadSkin (5)
    frmTransparent.Show

    flsTitle.Movie = App.Path + "\picture\anim\titlemovie"
    flsMovieCategory.Movie = App.Path + "\picture\anim\all"
    flsLogo.Movie = App.Path + "\picture\movie\logo.swf"

    Form2.Height = 0

    Picture3.Visible = False
    picKategori.Visible = False
    cbokategori.Visible = True
    txtCategory.Visible = True
    cbokategori.Visible = False

    flsLogo.Top = 0
    flsLogo.Width = 1875
    flsLogo.Height = 1240
    flsLogo.Left = 0
    flsLogo.Visible = True

    flsTitle.Top = 90
    flsTitle.Width = 3015
    flsTitle.Height = 1050
    flsTitle.Left = 1920
    flsTitle.Visible = True

    txtRemoteCode.Visible = False

    flsMovieCategory.Top = 210
    flsMovieCategory.Width = 2460
    flsMovieCategory.Height = 780
    flsMovieCategory.Left = 12600
    flsMovieCategory.Visible = True

    txtArtis.Left = 11760
    txtArtis.Top = 1320
    txtArtis.Height = 3135
    txtArtis.Width = 3015

    txtSinopsis.Left = 11760
    txtSinopsis.Top = 5040
    txtSinopsis.Height = 3015
    txtSinopsis.Width = 3015

    picKategori.Left = 4680
    picKategori.Top = 600
    picKategori.Height = 300
    picKategori.Width = 3375

    cbokategori.Font.Size = 16
    cbokategori.Left = 4680
    cbokategori.Width = 3855
    cbokategori.Top = Screen.Height + 8160
    cbokategori.BackColor = &HFFFFFF
    cbokategori.ForeColor = &H0&

    txtCategory.Top = 1260
    txtCategory.Left = 6240
    txtCategory.Width = 2895
    txtCategory.Font.Size = 16
    txtCategory.BackColor = &H244C&
    txtCategory.ForeColor = &HFFFFFF

    txtSearch.BackColor = &HFFFFFF
    txtSearch.ForeColor = &H0&
    txtSearch.Font.Size = 28
    txtSearch.Top = 380
    txtSearch.Left = 5040
    txtSearch.Height = 680
    txtSearch.Width = 6615
    txtSearch.BackColor = &HFFFFFF

    PictureLogin.Height = 0

    txtUser.Top = 3240
    txtUser.Left = 7560
    txtUser.Height = 0
    txtUser.Width = 0
    txtUser.BackColor = &HFFFFFF
    txtUser.ForeColor = &H800000

    txtTime.Top = 3240
    txtTime.Left = 7560
    txtTime.Height = 0
    txtTime.Width = 0
    txtTime.BackColor = &HFFFFFF
    txtTime.ForeColor = &H800000

    Dim sqlt As String
    Dim MyRst As MYSQL_RS
    sqlt = "SELECT id, negara FROM filmnegara"
    Set MyRst = MyConn.Execute(sqlt)
    cbokategori.Clear
    cbokategori.AddItem "ALL"
    While Not MyRst.EOF
        cbokategori.AddItem MyRst.Fields(1).value
        cbokategori.ItemData(cbokategori.NewIndex) = MyRst.Fields(0).value
        MyRst.MoveNext
    Wend
    cbokategori.ListIndex = 0
    txtCategory.text = cbokategori.text

    lstTV.Visible = False
    lstMidi.Visible = False
    lstMidiMusic.Visible = False
    lstMovie.Visible = True

    Unload frmCamera

    If vVideoAktif Then
        'vVideo = 99
        frmVideo.Show
        frmVideo.WindowsMediaPlayer1.Controls.play
        'vVideo = 5
    Else
        frmPromo.Show
        frmPromo.Refresh
    End If
    frmTransparent.Show

    If vMuteTemp = True Then
        VolTemp = 0
        tmrVolume.Enabled = True
    End If

    lstMovie.Font.Size = 20
    lstMovie.Top = 1920
    lstMovie.Left = 2280
    lstMovie.Height = 4695
    lstMovie.Width = 11625

    txtVol.Top = 6120
    txtVol.Left = 14160
    txtVol.Height = 645
    txtVol.Width = 900
    txtVol.Font.Size = 27
    txtVol.BackColor = &HFFFFFF
    txtVol.ForeColor = &H6F1628
    txtVol.Visible = True

    frmRoom.Height = 11520

    lstMovie.SetFocus
    txtSearch.SetFocus
    vpointer = 1
    UkuranVideo = 1

    vpbProsesEksekusi = False
End Sub

Private Sub txtSinopsis_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    KeyAscii = 0
    If vpointer = 1 Then
        sMakeCaret txtSearch, caretLebar, caretTinggi
    End If

  If Err.Number <> 0 Then
    LogError Name, "txtSinopsis_KeyPress"
  End If

End Sub

Public Sub PlayerKomputer()
    On Error Resume Next
    Dim vMuteTemp As Boolean
    Dim vLstAllHeightTemp As Integer
    Dim vCameraAktif As Boolean
    vMuteTemp = False

    If vpbProsesEksekusi Then
        Exit Sub
    End If

    vpbProsesEksekusi = True

    If vpbfrmCamera Then
        If (frmCamera.Width = Screen.Width) Or (vVideo = 1) Then
            setAudioEndPointVolumeMasterVolumeLevelPercent 0
            vMuteTemp = True
        End If
    End If

    vCameraAktif = False
    If vVideo = 7 Then
        frmCamera.Width = 7680
        frmCamera.Left = Screen.Width - frmCamera.Width
        frmCamera.Top = 950
        frmCamera.Height = 5700
        frmCamera.VideoCap1.Top = 0
        frmCamera.VideoCap1.Left = 75
        frmCamera.VideoCap1.Height = 5610
        frmCamera.VideoCap1.Width = 7530
        frmCamera.VideoCap1.TVMute = True
        frmCamera.VideoCap1.Start
        vCameraAktif = True
    Else
        Unload frmCamera
    End If

    vVideo = 0
    tutup
    frmRoom.vpointer = 1
    frmRoom.Height = 11520
    txtChatAktif.Visible = False
    txtPlaying.Visible = True
    frmRoom.flsTitle.Movie = App.Path + "\picture\anim\title"
    screennormal

    If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
        selIndex = 1
        lstAll.ListItems(1).Selected = True
        Set lstAll.DropHighlight = lstAll.selectedItem
        lstAll.selectedItem.Selected = False
    End If

    txtSearch.Visible = True
    vLstAllHeightTemp = lstAll.Height
    lstAll.Height = 0
    lstAll.Visible = True
    lstAll.SetFocus
    txtSearch.SetFocus
    DoEvents

    '--------SKIN COY----------
    Skin1.RemoveSkin
    'Aktifkan transparent color
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY

    frmTransparent.LoadSkin (0)
    If vpbHits = True Then
        frmTransparent.GantiSkin (2)
    ElseIf vpbNew = True Then
        frmTransparent.GantiSkin (3)
    ElseIf vpbPopuler = True Then
        frmTransparent.GantiSkin (4)
    End If
    txtVol.Visible = True

    If vVideoAktif Then
    '    vVideo = 99
        frmVideo.Show
        frmVideo.WindowsMediaPlayer1.Controls.play
    '    vVideo = 0
    Else
        frmPromo.Show
        frmPromo.Refresh
    End If

    If vCameraAktif Then frmCamera.Show
    frmTransparent.Show
    frmRoom.Show
    lstAll.Height = vLstAllHeightTemp

    If vMuteTemp = True Then
        VolTemp = 0
        tmrVolume.Enabled = True
    End If

    Dim sqlt As String
    Dim MyRst As MYSQL_RS
    sqlt = "SELECT TYPE, TYPENAME FROM kategori ORDER BY TYPE"
    Set MyRst = MyConn.Execute(sqlt)

    cbokategori.Clear
    cbokategori.AddItem "ALL"
    While Not MyRst.EOF
        cbokategori.AddItem MyRst.Fields(1).value
        cbokategori.ItemData(cbokategori.NewIndex) = MyRst.Fields(0).value
        MyRst.MoveNext
    Wend
    Set MyRst = Nothing
    cbokategori.ListIndex = 0

    UkuranVideo = 1
    If vVideoAktif Then
        frmVideo.Top = 0
        frmVideo.Left = 0
        frmVideo.Width = 15360
        frmVideo.Height = 11520
        frmVideo.WindowsMediaPlayer1.Top = 0
        frmVideo.WindowsMediaPlayer1.Left = 0
        frmVideo.WindowsMediaPlayer1.Height = frmVideo.Height
        frmVideo.WindowsMediaPlayer1.Width = frmVideo.Width
    End If
    vpbProsesEksekusi = False
End Sub

Public Sub PlayerMidi()
On Error Resume Next

    If vpbProsesEksekusi Then
        Exit Sub
    End If

    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    vpbProsesEksekusi = True

    If vVideoAktif Then
        frmVideo.WindowsMediaPlayer1.Controls.pause
    End If

    flsTitle.Visible = False
    txtSearch.Visible = False
    txtCategory.Visible = False
    picKeyTempo.Visible = False
    tmrPicKeyTempo.Enabled = False
    txtRemoteCode.Visible = False

    flsLogo.Movie = App.Path + "\picture\anim\" + "mainmin"
    flsLogo.Top = 120
    flsLogo.Width = 1560
    flsLogo.Height = 840
    flsLogo.Left = 360
    flsLogo.Visible = True

    txtUser.Top = 220
    txtUser.Left = 13800
    txtUser.Height = 270
    txtUser.Width = 1575
    txtUser.BackColor = &HFFFFFF
    txtUser.ForeColor = &H22993

    txtTime.Top = 600
    txtTime.Left = 13800
    txtTime.Height = 270
    txtTime.Width = 1575
    txtTime.BackColor = &HFFFFFF
    txtTime.ForeColor = &H22993

    frmRoom.Height = 10
    DoEvents

    setAudioEndPointVolumeMasterVolumeLevelPercent 0


    vpointer = 1

    txtSinopsis.Visible = False
    txtArtis.Visible = False


    flsMovieCategory.Visible = False
    txtChatAktif.Visible = False
    flsChatNewMessage.Visible = False
    Picture3.Visible = False
    picKategori.Visible = False

    'Buka Transparent
    If vVideo = 0 Or vVideo = 5 Or vVideo = 7 Then
        Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
        frmTransparent.Hide
    End If

    '--------SKIN COY----------
    lokasi = App.Path
    Skin1.LoadSkin lokasi + "\skin\sknmidi.skn"
    Skin1.ApplySkinByName hWnd, "sknmidi"

    vVideo = 1

    cbokategori.Font.Size = 16
    cbokategori.Left = 5280
'    cbokategori.Top = 540
    cbokategori.Width = 5055
    cbokategori.BackColor = &HFFFFFF
    cbokategori.ForeColor = &H22993
    cbokategori.Visible = False

    picKategori.Left = 5280
    picKategori.Top = 600
    picKategori.Height = 400
    picKategori.Width = 4215

    txtSearch.Top = 9720
    txtSearch.Left = 4920
    txtSearch.Height = 495
    txtSearch.Width = 7065
    txtSearch.BackColor = &HFFFFFF
    txtSearch.ForeColor = &H22993

    txtVol.Top = 10200
    txtVol.Left = 14040
    txtVol.Height = 645
    txtVol.Width = 900
    txtVol.Font.Size = 27
    txtVol.BackColor = &HFFFFFF
    txtVol.ForeColor = &H22993

    Unload frmPromo
    lstMidi.Left = 720
    lstMidi.Top = 1320
    lstMidi.Height = 10
    lstMidi.Width = 13935
    lstMidi.Visible = True
    lstMidiMusic.Left = lstMidi.Left
    lstMidiMusic.Top = lstMidi.Top
    lstMidiMusic.Height = lstMidi.Height
    lstMidiMusic.Width = lstMidi.Width
    lstMidiMusic.Visible = False
    lstTV.Visible = False
    frmRoom.Height = 1450
    Form2.Height = 1095

    MidiAktif
    frmCamera.VideoCap1.TVMute = True

    txtVol.Visible = True

    setAudioEndPointVolumeMute 0

    VolTemp = 0
    tmrVolume.Enabled = True

    DoEvents
    Form2.Show
    frmRoom.Show
    vpbProsesEksekusi = False
End Sub

Public Sub PlayerTV()
    On Error Resume Next

    Dim LV As ListItem
    Dim VideoFormatIndex As Integer

    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    If vpbProsesEksekusi Then
        Exit Sub
    End If

    vpbProsesEksekusi = True

    frmRoom.ScreenSaverAktif = 0
    If vVideoAktif Then
        frmVideo.WindowsMediaPlayer1.Controls.pause
    End If

    Form2.Height = 0

    cbokategori.ListIndex = 0
    picKeyTempo.Visible = False
    tmrPicKeyTempo.Enabled = False
    txtPlaying.Visible = False

    txtUser.Top = 3240
    txtUser.Left = 7560
    txtUser.Height = 0
    txtUser.Width = 0
    txtUser.BackColor = &HFFFFFF
    txtUser.ForeColor = &H4000&

    txtTime.Top = 3240
    txtTime.Left = 7560
    txtTime.Height = 0
    txtTime.Width = 0
    txtTime.BackColor = &HFFFFFF
    txtTime.ForeColor = &H4000&

    txtVol.Top = 6120
    txtVol.Left = 14160
    txtVol.Height = 645
    txtVol.Width = 900
    txtVol.Font.Size = 27
    txtVol.BackColor = &HFFFFFF
    txtVol.ForeColor = &H6F1628
    txtVol.Visible = True

    lstTV.Font.Size = 20
    lstTV.Top = 1920
    lstTV.Left = 2280
    lstTV.Height = 4695
    lstTV.Width = 11625

    flsLogo.Top = 0
    flsLogo.Width = 1875
    flsLogo.Height = 1240
    flsLogo.Left = 0
    flsLogo.Visible = True

    flsTitle.Visible = False
    flsMovieCategory.Visible = False
    txtSearch.Visible = False
    txtChatAktif.Visible = False
    flsChatNewMessage.Visible = False
    txtRemoteCode.Visible = False

    lstAll.Visible = False
    lstPlaylist.Visible = False
    lstTV.Visible = True
    lstMidi.Visible = False
    lstMidiMusic.Visible = False
    lstChat.Visible = False
    lstMovie.Visible = False
    txtChat.Visible = False

    Picture3.Visible = False
    txtCategory.Visible = False
    picKategori.Visible = False
    cbokategori.Visible = False

    lstTV.ListItems.Clear
    Dim sqlm As String
    Dim Im As Integer
    Dim MyRsm As MYSQL_RS

    sqlm = "SELECT ID, CHANNEL, FREKWENSI FROM TVCHANNEL ORDER BY CHANNEL;"
    Set MyRsm = MyConn.Execute(sqlm)

    'Isi list data
    Do Until MyRsm.EOF
        Set LV = lstTV.ListItems.add(, , (MyRsm.Fields(0).value))
        LV.SubItems(1) = MyRsm.Fields(1).value
        LV.SubItems(2) = MyRsm.Fields(2).value
        MyRsm.MoveNext
    Loop

    ShowScrollBar lstTV.hWnd, SB_VERT, False

    lstAll.Visible = False
    lstTV.Visible = True
    lstTV.SetFocus
    frmBackground.Show
    frmTransparent.Show
    frmRoom.Show
    frmRoom.Height = 11520
    DoEvents

    '--------SKIN COY----------
    If (vVideo <> 0) Or (vVideo <> 5) Or (vVideo <> 7) Then
        Skin1.RemoveSkin
        'Aktifkan transparent color
        SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hWnd, &H244C&, 0&, LWA_COLORKEY
    End If

    frmTransparent.LoadSkin (7)

    flsLogo.Movie = App.Path + "\picture\tv\logo.swf"

    setAudioEndPointVolumeMasterVolumeLevelPercent 0

'    frmCamera.Show
    frmCamera.VideoCap1.Visible = True
    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.TVMute = True

    frmCamera.Left = 0
    frmCamera.Top = 0
    frmCamera.Width = Screen.Width
    frmCamera.Height = Screen.Height
    frmCamera.VideoCap1.Top = 0
    frmCamera.VideoCap1.Left = 0
    frmCamera.VideoCap1.Width = frmCamera.Width
    frmCamera.VideoCap1.Height = frmCamera.Height

    frmCamera.VideoCap1.TVMute = False

    VolTemp = 0
    tmrVolume.Enabled = True

    ControlCap ("Video Tuner")
    frmCamera.VideoCap1.CountryCode = 62
    frmCamera.VideoCap1.VideoStandard = 4
    VideoFormatIndex = frmCamera.VideoCap1.VideoFormats.FindVideoFormat("RGB555 (720x576)")
    If VideoFormatIndex <> -1 Then
          frmCamera.VideoCap1.VideoFormat = VideoFormatIndex
    End If

    frmCamera.VideoCap1.Start
    frmCamera.VideoCap1.AspectRatio = False

'    lstTV_Click


    vVideo = 7

    frmCamera.Show
    frmTransparent.Show

    DoEvents

    lstTV.SetFocus
    ShowScrollBar lstTV.hWnd, SB_VERT, False

    UkuranVideo = 1
    frmRoom.ScreenSaverAktif = 0
    vpbProsesEksekusi = False
End Sub

Public Sub prcRepeat()

    On Error Resume Next

    If Not (vVideoAktif) Then
        Exit Sub
    End If

    If getApa() = "buka" Then
        frmVideo.pbDurasiAkhir = 0
        frmVideo.WindowsMediaPlayer1.Controls.currentPosition = 1
        ScoreValid = True
    Else
        cmdStop_Click
        ScreenSaverAktif = 5
    End If


    If Err.Number <> 0 Then
      LogError Name, "prcRepeat"
    End If
End Sub

Public Sub prcVCD()
    On Error Resume Next
    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    frmConfirmasi.vpengirim = 7
    frmConfirmasi.Text1.text = ""
    frmConfirmasi.Text2 = "Record to VCD ?"
    frmConfirmasi.text3.text = ""
    frmConfirmasi.Show
End Sub

Public Sub prcCD()
    On Error Resume Next
    If vrekamstate Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If
    frmConfirmasi.vpengirim = 8
    frmConfirmasi.Text1.text = ""
    frmConfirmasi.Text2 = "Record Audio ?"
    frmConfirmasi.text3.text = ""
    frmConfirmasi.Show
    frmConfirmasi.Top = 1550
End Sub

Public Sub prcDVD()
    On Error Resume Next
    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Recording!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    frmConfirmasi.vpengirim = 11
    frmConfirmasi.Text1.text = ""
    If vpbRemoteStatus = 1 Then
        frmConfirmasi.Text2 = "Record to DVD ?"
    Else
        frmConfirmasi.Text2 = "Record Video ?"
    End If
    frmConfirmasi.text3.text = ""
    frmConfirmasi.Show
    frmConfirmasi.Top = 1550
End Sub

Public Sub prcHP()
    On Error Resume Next
    If vrekamstate = True Then
        frmConfirmasi.vpengirim = 6
        frmConfirmasi.Text1.text = ""
        frmConfirmasi.Text2.text = "Sedang Merekam!"
        frmConfirmasi.text3.text = ""
        frmConfirmasi.Show
        Exit Sub
    End If

    frmConfirmasi.vpengirim = 10
    frmConfirmasi.Text1.text = ""
    frmConfirmasi.Text2 = "Record to HP ?"
    frmConfirmasi.text3.text = ""
    frmConfirmasi.Show

'    frmCamera.Width = 7680
End Sub

Public Sub prcMute()
    On Error Resume Next
    If vpbMute = True Then
        vpbMute = False
        setAudioEndPointVolumeMute 0
    Else
        vpbMute = True
        setAudioEndPointVolumeMute 1
    End If
End Sub

Sub prcHits()
    On Error Resume Next
    vpbNew = False
    vpbPopuler = False
    Dim lokasi As String
    txtSearch.text = ""
    lokasi = App.Path + "\picture\normalscreen\"
    If (vpbHits = True) And (vpointer <> 3) Then
        vpbHits = False
        cmdSong_Click
    Else
        frmTransparent.GantiSkin (2)

        lstAll.Visible = True
        lstPlaylist.Visible = False

        If vpointer = 3 Then
            vpointer = vpointerTemp
        End If

        If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
            sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
        End If
        vpbHits = True
    End If
    CariTitle
End Sub

Sub prcPopuler(x As String)
    On Error Resume Next

    vpbJumlahNewPopuler = Val(x)
    vpbPopuler = True
    vpbNew = False
    vpbHits = False
    frmTransparent.GantiSkin (4)

    lstAll.Visible = True
    lstPlaylist.Visible = False

    If vpointer = 3 Then
        vpointer = vpointerTemp
    End If

    txtSearch.text = ""
    CariTitle
    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
End Sub

Sub prcNew(x As String)
    On Error Resume Next

    vpbJumlahNewPopuler = Val(x)
    vpbNew = True
    vpbPopuler = False
    vpbHits = False
    frmTransparent.GantiSkin (3)
    lstAll.Visible = True
    lstPlaylist.Visible = False

        If vpointer = 3 Then
            vpointer = vpointerTemp
        End If

    txtSearch.text = ""
    CariTitle
    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
End Sub

Sub prcSaran()
    On Error Resume Next
    frmSaran.Show
    frmSaran.txtPesan.SetFocus
    sMakeCaret frmSaran.txtPesan, 7, 30
End Sub

Sub prcKeyUp()
    On Error Resume Next

    Dim i As Integer
    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True

'    If vKey < 5 Then
'        incPitch 100
'        vKey = vKey + 1
'    End If

    If vKey < 4 Then
        tempPitch 0
        incPitch 100
        vKey = vKey + 1
    End If

    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i

    For i = 1 To 5
        If vKey > 0 Then
            If i <= vKey Then
                KotakTambah(i - 1).Visible = True
            Else
                KotakTambah(i - 1).Visible = False
            End If
        End If
        If vKey < 0 Then
            If i <= Abs(vKey) Then
                KotakKurang(i - 1).Visible = True
            Else
                KotakKurang(i - 1).Visible = False
            End If
        End If
    Next i
    picKeyTempo.Visible = True
End Sub

Sub prckeyDown()
    On Error Resume Next

    Dim i As Integer

    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True

    If vKey > -4 Then
        tempPitch 0
        incPitch -100
        vKey = vKey - 1
    End If

    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i

    For i = 1 To 5
        If vKey > 0 Then
            If i <= vKey Then
                KotakTambah(i - 1).Visible = True
            Else
                KotakTambah(i - 1).Visible = False
            End If
        End If
        If vKey < 0 Then
            If i <= Abs(vKey) Then
                KotakKurang(i - 1).Visible = True
            Else
                KotakKurang(i - 1).Visible = False
            End If
        End If
    Next i
    picKeyTempo.Visible = True
End Sub

Sub prcKeyReset()
    Dim lokasi As String
    On Error Resume Next
    lokasi = App.Path
    picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\key.jpg")

    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True

    Dim i As Integer
    SetPitch 0
    vKey = 0
    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i
    picKeyTempo.Visible = True
End Sub

Sub prcTempoUp()
    On Error Resume Next

    Dim i As Integer
    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True
    If vTempo < 5 Then
        IncTempo 0.05
        vTempo = vTempo + 1
    End If

    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i

    For i = 1 To 5
        If vTempo > 0 Then
            If i <= vTempo Then
                KotakTambah(i - 1).Visible = True
            Else
                KotakTambah(i - 1).Visible = False
            End If
        End If
        If vTempo < 0 Then
            If i <= Abs(vTempo) Then
                KotakKurang(i - 1).Visible = True
            Else
                KotakKurang(i - 1).Visible = False
            End If
        End If
    Next i
    picKeyTempo.Visible = True
End Sub

Sub prcTempoDown()
    On Error Resume Next

    Dim i As Integer

    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True

    If vTempo > -5 Then
        IncTempo -0.05
        vTempo = vTempo - 1
    End If

    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i

    For i = 1 To 5
        If vTempo > 0 Then
            If i <= vTempo Then
                KotakTambah(i - 1).Visible = True
            Else
                KotakTambah(i - 1).Visible = False
            End If
        End If
        If vTempo < 0 Then
            If i <= Abs(vTempo) Then
                KotakKurang(i - 1).Visible = True
            Else
                KotakKurang(i - 1).Visible = False
            End If
        End If
    Next i
    picKeyTempo.Visible = True
End Sub

Sub prcTempoReset()
    Dim lokasi As String
    On Error Resume Next
    lokasi = App.Path
    picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\tempo.jpg")

    Dim i As Integer
    tmrPicKeyTempo.Enabled = False
    tmrPicKeyTempo.Enabled = True

    SetTempo 1
    frmRoom.vTempo = 0

    For i = 1 To 5
        KotakTambah(i - 1).Visible = False
        KotakKurang(i - 1).Visible = False
    Next i

    picKeyTempo.Visible = True
End Sub

Private Sub TambahLstAll()

    On Error Resume Next

    Dim cariapa As String
    Dim sqlm As String
    Dim LV As UniToolbox2.ListItem
    Dim Title2 As Boolean
    Dim Text1 As String
    Dim Path As String

    Title2 = False

    Select Case vpointer
        Case 2
            cariapa = "spSinger"
        Case 6
            cariapa = "spCode"
        Case 7
            cariapa = "spAbrTitle"
        Case 8
            cariapa = "spAbrSinger"
        Case Else
            cariapa = "spTitle"
    End Select

    Dim textbersih As String
    textbersih = txtSearch.text
    textbersih = Replace$(textbersih, " ", "")
    textbersih = Replace$(textbersih, "'", "")
    textbersih = Replace$(textbersih, ",", "")
    textbersih = Replace$(textbersih, ".", "")
    textbersih = mysql_escape_string(textbersih)

    If vpbHits = True Then
        If (cbokategori.ListIndex = 0) Then
            cariapa = cariapa & "Hits"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
        Else
            cariapa = cariapa & "HitsKategori"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
        End If
    ElseIf vpbNew = True Then
        cariapa = cariapa & "New"
        If vLstAllRow < vpbJumlahNewPopuler Then
            If (cbokategori.ListIndex = 0) Then
                If textbersih = "" Then
                    cariapa = cariapa & "Empty"
                    sqlm = "CALL " & cariapa & "(" & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            Else
                cariapa = cariapa & "Type"
                If textbersih = "" Then
                    cariapa = cariapa & "Empty"
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            End If
        Else
'                lErr = LockWindowUpdate(0)
            Exit Sub
        End If
    ElseIf vpbPopuler = True Then
        cariapa = cariapa & "Popular"
        If vLstAllRow < vpbJumlahNewPopuler Then
            If (cbokategori.ListIndex = 0) Then
                If textbersih = "" Then
                    cariapa = cariapa & "Empty"
                    sqlm = "CALL " & cariapa & "(" & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            Else
                cariapa = cariapa & "Type"
                If textbersih = "" Then
                    cariapa = cariapa & "Empty"
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            End If
        Else
'               lErr = LockWindowUpdate(0)
            Exit Sub
        End If
    Else
        If (cbokategori.ListIndex = 0) Then
           sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
        Else
            cariapa = cariapa & "K"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
        End If
    End If

    Set rsAdo = New ADODB.Recordset
    rsAdo.Open sqlm, KoneksiAdoDB, adOpenStatic, adLockOptimistic

    If frmUser.settingStatusAutoAbbreviation = 1 Then
      If rsAdo.recordCount = 0 Then
        Select Case vpointer
          Case 1
            vpointer = 7
            TambahLstAll
            If lstAll.ListItems.Count > 0 Then
              Path = App.Path + "\picture\anim\abr-title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 1
            Exit Sub
          Case 2
            vpointer = 8
            TambahLstAll
            If lstAll.ListItems.Count <> 0 Then
              Path = App.Path + "\picture\anim\abr-artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 2
            Exit Sub
        End Select
      End If
    End If

    If (rsAdo.recordCount > 0) Then
        vLstAllRow = vLstAllRow + 1
        If Not lstAll.ListItems(1).SubItems(3) = lstAll.ListItems(2).SubItems(3) Then
            vlstAllRecord = vlstAllRecord - 1
        End If

        'HAPUS YG DIATAS CEK CINANYA DULU
        If lstAll.ListItems(2).SubItems(7) = "1" Then
            lstAll.ListItems.Remove (2)
        Else
            lstAll.ListItems.Remove (1)
        End If
        'Isi list data
        vlstAllRecord = vlstAllRecord + 1
        Set LV = lstAll.ListItems.add(, , (rsAdo.Fields("CODE").value))

        Text1 = UCase$(textbersih)
        If vpointer = 1 Then
            If Text1 = UCase$(Left$(rsAdo.Fields("TITLE3").value, Len(textbersih))) Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            ElseIf rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 2 Then
            If rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            If Text1 = UCase$(Left$(rsAdo.Fields("SINGER3").value, Len(textbersih))) Then
              LV.SubItems(2) = rsAdo.Fields("SINGER").value
            ElseIf rsAdo.Fields("SINGER4").value = "" Then
              LV.SubItems(2) = rsAdo.Fields("SINGER").value
            Else
              LV.SubItems(2) = rsAdo.Fields("SINGER4").value
            End If
        ElseIf vpointer = 7 Then
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("TITLE").value), Len(Text1))) Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          ElseIf rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 8 Then
          If rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("SINGER").value), Len(Text1))) Then
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
          ElseIf rsAdo.Fields("SINGER4").value = "" Then
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
          Else
            LV.SubItems(2) = rsAdo.Fields("SINGER4").value
          End If
        Else
            If rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
        End If

        LV.SubItems(3) = rsAdo.Fields(2).value
        LV.SubItems(4) = rsAdo.Fields(3).value
        LV.SubItems(5) = rsAdo.Fields(4).value
        LV.SubItems(6) = rsAdo.Fields(5).value

        If (Not rsAdo.Fields("TITLE2").value = "") Then
            Title2 = True
            If Not lstAll.ListItems(1).SubItems(3) = lstAll.ListItems(2).SubItems(3) Then
                vlstAllRecord = vlstAllRecord - 1
            End If
            lstAll.ListItems.Remove (1)
            Set LV = lstAll.ListItems.add(, , "")
            LV.SubItems(1) = rsAdo.Fields("TITLE2").value
            LV.SubItems(2) = rsAdo.Fields("SINGER2").value
            LV.SubItems(3) = rsAdo.Fields(2).value
            LV.SubItems(4) = rsAdo.Fields(3).value
            LV.SubItems(5) = rsAdo.Fields(4).value
            LV.SubItems(6) = rsAdo.Fields(5).value
            LV.SubItems(7) = "1"
            LV.Checked = True
            LV.ForeColor = &HFFFFFF
        End If
    End If

    Set rsAdo = Nothing

    Select Case vpointer
      Case 1
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\title"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
      Case 2
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\artist"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
    End Select

    If lstAll.ListItems.Count > 0 Then
        If Title2 = True Then
            selIndex = 8
            lstAll.ListItems(8).Selected = True
        Else
            selIndex = 9
            lstAll.ListItems(9).Selected = True
        End If
        Set lstAll.DropHighlight = lstAll.selectedItem
        lstAll.selectedItem.Selected = False
    End If

    If Err.Number <> 0 Then
        LogError Name, "TambahLstAll"
    End If
End Sub

Public Sub CariTitle()

    On Error Resume Next

    Dim cariapa As String
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim textbersih As String
    Dim LV As UniToolbox2.ListItem
    Dim i, x As Integer
    Dim Text1 As String
    Dim Path As String

    Select Case vpointer
        Case 2
            cariapa = "spSinger"
        Case 6
            cariapa = "spCode"
        Case 7
            cariapa = "spAbrTitle"
        Case 8
            cariapa = "spAbrSinger"
        Case Else
            cariapa = "spTitle"
    End Select

    textbersih = txtSearch.text
    textbersih = Replace$(textbersih, " ", "")
    textbersih = Replace$(textbersih, "'", "")
    textbersih = Replace$(textbersih, ",", "")
    textbersih = Replace$(textbersih, ".", "")
    textbersih = mysql_escape_string(textbersih)

    If vpbHits = True Then
        If (cbokategori.ListIndex = 0) Then
            cariapa = cariapa & "Hits"
            Sql = "CALL " & cariapa & "('" & textbersih & "%',0,9)"
        Else
            cariapa = cariapa & "HitsKategori"
            Sql = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ",0,9)"
        End If
    ElseIf vpbNew = True Then
        cariapa = cariapa & "New"
        If (cbokategori.ListIndex = 0) Then
            If textbersih = "" Then
                cariapa = cariapa & "Empty"
                Sql = "CALL " & cariapa & "(0,9)"
            Else
                Sql = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & " 0,9)"
            End If
        Else
            cariapa = cariapa & "Type"
            If textbersih = "" Then
                cariapa = cariapa & "Empty"
                Sql = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", 0,9)"
            Else
                Sql = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%', 0,9)"
            End If
        End If
    ElseIf vpbPopuler = True Then
        cariapa = cariapa & "Popular"
        If (cbokategori.ListIndex = 0) Then
            If textbersih = "" Then
                cariapa = cariapa + "Empty"
                Sql = "CALL " & cariapa & "(0,9)"
            Else
                Sql = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & " 0,9)"
            End If
        Else
            cariapa = cariapa & "Type"
            If textbersih = "" Then
                cariapa = cariapa + "Empty"
                Sql = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", 0,9)"
            Else
                Sql = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%', 0,9)"
            End If
        End If
    Else
        If (cbokategori.ListIndex = 0) Then
             Sql = "CALL " & cariapa & "('" & textbersih & "%',0,9)"
        Else
            cariapa = cariapa & "K"
            Sql = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ",0,9)"
        End If
    End If

    Set rsAdo = New ADODB.Recordset
    rsAdo.Open Sql, KoneksiAdoDB, adOpenStatic, adLockOptimistic

    If Not (rsAdo.State = adStateOpen) Then
        KonekServer
        rsAdo.Open Sql, KoneksiAdoDB, adOpenStatic, adLockOptimistic
    End If

    If frmUser.settingStatusAutoAbbreviation = 1 Then
      If rsAdo.recordCount = 0 Then
        Select Case vpointer
          Case 1
            vpointer = 7
            CariTitle
            If lstAll.ListItems.Count > 0 Then
              Path = App.Path + "\picture\anim\abr-title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 1
            GoTo hell
          Case 2
            vpointer = 8
            CariTitle
            If lstAll.ListItems.Count <> 0 Then
              Path = App.Path + "\picture\anim\abr-artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 2
            GoTo hell
        End Select
      End If
    End If

    lErr = LockWindowUpdate(lstAll.hWnd)
    lstAll.ListItems.Clear
    i = 0
    x = 0
    Do Until rsAdo.EOF
        'Isi list data
        i = i + 1
        x = x + 1
        Set LV = lstAll.ListItems.add(, , (rsAdo.Fields("CODE").value))

        Text1 = UCase$(textbersih)
        If vpointer = 1 Then
            If Text1 = UCase$(Left$(rsAdo.Fields("TITLE3").value, Len(textbersih))) Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            ElseIf rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 2 Then
            If rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            If Text1 = UCase$(Left$(rsAdo.Fields("SINGER3").value, Len(textbersih))) Then
              LV.SubItems(2) = rsAdo.Fields("SINGER").value
            ElseIf rsAdo.Fields("SINGER4").value = "" Then
              LV.SubItems(2) = rsAdo.Fields("SINGER").value
            Else
              LV.SubItems(2) = rsAdo.Fields("SINGER4").value
            End If
        ElseIf vpointer = 7 Then
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("TITLE").value), Len(Text1))) Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          ElseIf rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 8 Then
          If rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("SINGER").value), Len(Text1))) Then
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
          ElseIf rsAdo.Fields("SINGER4").value = "" Then
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
          Else
            LV.SubItems(2) = rsAdo.Fields("SINGER4").value
          End If
        Else
          If rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          LV.SubItems(2) = rsAdo.Fields("SINGER").value
        End If

        LV.SubItems(3) = rsAdo.Fields(2).value
        LV.SubItems(4) = rsAdo.Fields(3).value
        LV.SubItems(5) = rsAdo.Fields(4).value
        LV.SubItems(6) = rsAdo.Fields(5).value

        If (Not rsAdo.Fields("TITLE2").value = "") And (i < 9) Then
            i = i + 1
            Set LV = lstAll.ListItems.add(, , "")
            LV.SubItems(1) = rsAdo.Fields("TITLE2").value
            LV.SubItems(2) = rsAdo.Fields("SINGER2").value
            LV.SubItems(3) = rsAdo.Fields(2).value
            LV.SubItems(4) = rsAdo.Fields(3).value
            LV.SubItems(5) = rsAdo.Fields(4).value
            LV.SubItems(6) = rsAdo.Fields(5).value
            LV.SubItems(7) = "1"
            LV.Checked = True
            LV.ForeColor = &HFFFFFF
            '&H8E4DA4
        End If
        If i >= 9 Then rsAdo.MoveLast
        rsAdo.MoveNext
    Loop
    vLstAllRow = x
    vlstAllRecord = x
    lErr = LockWindowUpdate(0)
    Set rsAdo = Nothing

    Select Case vpointer
      Case 1
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\title"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
      Case 2
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\artist"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
    End Select

    If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
            Set lstAll.DropHighlight = lstAll.selectedItem
            selIndex = lstAll.selectedItem.index
            lstAll.selectedItem.Selected = False
    End If

hell:

    If Err.Number <> 0 Then
      LogError Name, "CariTitle: " & Sql
    End If

End Sub

Private Sub KurangLstAll()

    On Error Resume Next

    Dim cariapa As String
    Dim sqlm As String
    Dim LV As UniToolbox2.ListItem
    Dim textbersih As String
    Dim i As Integer
    Dim Text1 As String
    Dim Path As String

    Select Case vpointer
        Case 2
            cariapa = "spSinger"
        Case 6
            cariapa = "spCode"
        Case 7
            cariapa = "spAbrTitle"
        Case 8
            cariapa = "spAbrSinger"
        Case Else
            cariapa = "spTitle"
    End Select

    i = vLstAllRow - vlstAllRecord
    If i <= 0 Then
        Exit Sub
    End If

'    lErr = LockWindowUpdate(lstAll.hWnd)

    vLstAllRow = vLstAllRow - vlstAllRecord - 1

    textbersih = txtSearch.text
    textbersih = Replace$(textbersih, " ", "")
    textbersih = Replace$(textbersih, "'", "")
    textbersih = Replace$(textbersih, ",", "")
    textbersih = Replace$(textbersih, ".", "")
    textbersih = mysql_escape_string(textbersih)

    If vpbHits = True Then
        If (cbokategori.ListIndex = 0) Then
            cariapa = cariapa & "Hits"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
        Else
            cariapa = cariapa & "HitsKategori"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
        End If
    ElseIf vpbNew = True Then
        cariapa = cariapa & "New"
        If vLstAllRow < vpbJumlahNewPopuler Then
            If (cbokategori.ListIndex = 0) Then
                If textbersih = "" Then
                    cariapa = cariapa + "Empty"
                    sqlm = "CALL " & cariapa & "(" & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            Else
                cariapa = cariapa + "Type"
                If textbersih = "" Then
                    cariapa = cariapa + "Empty"
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            End If
        Else
'               lErr = LockWindowUpdate(0)
            Exit Sub
        End If
    ElseIf vpbPopuler = True Then
        cariapa = cariapa & "Popular"
        If vLstAllRow < vpbJumlahNewPopuler Then
            If (cbokategori.ListIndex = 0) Then
                If textbersih = "" Then
                    cariapa = cariapa + "Empty"
                    sqlm = "CALL " & cariapa & "(" & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            Else
                cariapa = cariapa + "Type"
                If textbersih = "" Then
                    cariapa = cariapa + "Empty"
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
                Else
                    sqlm = "CALL " & cariapa & "(" & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vpbJumlahNewPopuler) & ", '" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
                End If
            End If
        Else
'                lErr = LockWindowUpdate(0)
            Exit Sub
        End If
    Else
        If (cbokategori.ListIndex = 0) Then
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & Str(vLstAllRow) & "," & Str(1) & ")"
        Else
            cariapa = cariapa & "K"
            sqlm = "CALL " & cariapa & "('" & textbersih & "%'," & cbokategori.ItemData(cbokategori.ListIndex) & ", " & Str(vLstAllRow) & "," & Str(1) & ")"
        End If
    End If

    Set rsAdo = New ADODB.Recordset
    rsAdo.Open sqlm, KoneksiAdoDB, adOpenStatic, adLockOptimistic

    If frmUser.settingStatusAutoAbbreviation = 1 Then
      If rsAdo.recordCount = 0 Then
        Select Case vpointer
          Case 1
            vpointer = 7
            KurangLstAll
            If lstAll.ListItems.Count > 0 Then
              Path = App.Path + "\picture\anim\abr-title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\title"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 1
            Exit Sub
          Case 2
            vpointer = 8
            KurangLstAll
            If lstAll.ListItems.Count <> 0 Then
              Path = App.Path + "\picture\anim\abr-artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            Else
              Path = App.Path + "\picture\anim\artist"
              If flsTitle.Movie <> Path Then
                flsTitle.Movie = Path
              End If
            End If
            vpointer = 2
            Exit Sub
        End Select
      End If
    End If

    If rsAdo.recordCount > 0 Then
        If Not lstAll.ListItems(lstAll.ListItems.Count).SubItems(3) = lstAll.ListItems(lstAll.ListItems.Count - 1).SubItems(3) Then
             vlstAllRecord = vlstAllRecord - 1
        End If
        lstAll.ListItems.Remove (lstAll.ListItems.Count)
        'Isi list data
        Set LV = lstAll.ListItems.add(1, , (rsAdo.Fields("CODE").value))
        vlstAllRecord = vlstAllRecord + 1

        Text1 = UCase$(textbersih)
        If vpointer = 1 Then
            If Text1 = UCase$(Left$(rsAdo.Fields("TITLE3").value, Len(textbersih))) Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            ElseIf rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 2 Then
            If rsAdo.Fields("TITLE4").value = "" Then
              LV.SubItems(1) = rsAdo.Fields("TITLE").value
            Else
              LV.SubItems(1) = rsAdo.Fields("TITLE4").value
            End If
            If Text1 = UCase$(Left$(rsAdo.Fields("SINGER3").value, Len(textbersih))) Then
              LV.SubItems(2) = rsAdo.Fields("SINGER").value
            Else
              LV.SubItems(2) = rsAdo.Fields("SINGER4").value
            End If
        ElseIf vpointer = 7 Then
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("TITLE").value), Len(Text1))) Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          ElseIf rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          LV.SubItems(2) = rsAdo.Fields("SINGER").value
        ElseIf vpointer = 8 Then
          If rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          If Text1 = UCase$(Left$(getAbbreviation(rsAdo.Fields("SINGER").value), Len(Text1))) Then
            LV.SubItems(2) = rsAdo.Fields("SINGER").value
          Else
            LV.SubItems(2) = rsAdo.Fields("SINGER4").value
          End If
        Else
          If rsAdo.Fields("TITLE4").value = "" Then
            LV.SubItems(1) = rsAdo.Fields("TITLE").value
          Else
            LV.SubItems(1) = rsAdo.Fields("TITLE4").value
          End If
          LV.SubItems(2) = rsAdo.Fields("SINGER").value
        End If

        LV.SubItems(3) = rsAdo.Fields(2).value
        LV.SubItems(4) = rsAdo.Fields(3).value
        LV.SubItems(5) = rsAdo.Fields(4).value
        LV.SubItems(6) = rsAdo.Fields(5).value

        If (Not rsAdo.Fields("TITLE2").value = "") Then
            If Not lstAll.ListItems(lstAll.ListItems.Count).SubItems(3) = lstAll.ListItems(lstAll.ListItems.Count - 1).SubItems(3) Then
                vlstAllRecord = vlstAllRecord - 1
            End If

            lstAll.ListItems.Remove (lstAll.ListItems.Count)
            Set LV = lstAll.ListItems.add(2, , "")
            LV.SubItems(1) = rsAdo.Fields("TITLE2").value
            LV.SubItems(2) = rsAdo.Fields("SINGER2").value
            LV.SubItems(3) = rsAdo.Fields(2).value
            LV.SubItems(4) = rsAdo.Fields(3).value
            LV.SubItems(5) = rsAdo.Fields(4).value
            LV.SubItems(6) = rsAdo.Fields(5).value
            LV.SubItems(7) = "1"
            LV.Checked = True
            LV.ForeColor = &HFFFFFF
        End If
    End If

    Set rsAdo = Nothing

    vLstAllRow = vLstAllRow + vlstAllRecord

    Select Case vpointer
      Case 1
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\title"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
      Case 2
        If lstAll.ListItems.Count > 0 Then
          Path = App.Path + "\picture\anim\artist"
          If flsTitle.Movie <> Path Then
            flsTitle.Movie = Path
          End If
        End If
    End Select

    If lstAll.ListItems.Count > 0 Then
        selIndex = 1
        lstAll.ListItems(1).Selected = True
        Set lstAll.DropHighlight = lstAll.selectedItem
        lstAll.selectedItem.Selected = False
    End If
'   lErr = LockWindowUpdate(0)

    If Err.Number <> 0 Then
        LogError Name, "KurangLstAll"
    End If
End Sub

Private Sub TambahLstMovie(Lview As ListView)
    On Error Resume Next
Dim cariapa As String
Dim sqlm As String
Dim id As Long
Dim chm As ColumnHeader
Dim LV As ListItem
Dim MyRsm As MYSQL_RS

lErr = LockWindowUpdate(lstMovie.hWnd)

    Select Case vpointer
    Case 2
        cariapa = "ARTIS" & " LIKE '%"
    Case Else
        cariapa = "TITLE LIKE '"
    End Select

    Dim textbersih As String
    textbersih = mysql_escape_string(txtSearch.text)

    sqlm = "SELECT ID, TITLE FROM FILM WHERE " & cariapa & textbersih & "%' "
              If vMovieKategori = 1 Then sqlm = sqlm & " AND drama  = 1 "
              If vMovieKategori = 2 Then sqlm = sqlm & " AND komedi = 1 "
              If vMovieKategori = 3 Then sqlm = sqlm & " AND action = 1 "
              If vMovieKategori = 4 Then sqlm = sqlm & " AND horor  = 1 "
              If vMovieKategori = 5 Then sqlm = sqlm & " AND kartun = 1 "
    If cbokategori.ListIndex > 0 Then
        sqlm = sqlm & " AND negara = " & cbokategori.ItemData(cbokategori.ListIndex)
    End If
    sqlm = sqlm & " ORDER BY TITLE LIMIT " & Str(vLstAllRow) & ", " & Str(1) & ";"

    vLstAllRow = vLstAllRow + 1

    Set MyRsm = MyConn.Execute(sqlm)
    If MyRsm.recordCount > 0 Then
        Lview.ListItems.Remove (1)
        'Isi list data
        Set LV = Lview.ListItems.add(, , (MyRsm.Fields(0).value))
        LV.SubItems(1) = MyRsm.Fields(1).value
        LV.SubItems(2) = MyRsm.Fields(0).value
    End If
    Set MyRsm = Nothing

    If Lview.ListItems.Count > 0 Then
        Lview.ListItems(Lview.ListItems.Count).Selected = True
        Lview.selectedItem.Selected = False
        Set Lview.DropHighlight = Lview.selectedItem
    End If

   lErr = LockWindowUpdate(0)
'   DoEvents
End Sub

Private Sub KurangLstMovie(Lview As ListView)
    On Error Resume Next
Dim cariapa As String
Dim sqlm As String
Dim chm As ColumnHeader
Dim LV As ListItem
Dim MyRsm As MYSQL_RS
Dim textbersih As String

    Select Case vpointer
    Case 2
        cariapa = "SINGER" & " LIKE '%"
    Case Else
        cariapa = "TITLE LIKE '"
    End Select

    lErr = LockWindowUpdate(Lview.hWnd)

    vLstAllRow = vLstAllRow - Lview.ListItems.Count

    If vLstAllRow <= 0 Then
        lErr = LockWindowUpdate(0)
         vLstAllRow = Lview.ListItems.Count
        Exit Sub
    End If

    vLstAllRow = vLstAllRow - 1
    textbersih = mysql_escape_string(txtSearch.text)
    sqlm = "SELECT ID, TITLE FROM FILM WHERE " & cariapa & textbersih & "%' "
              If vMovieKategori = 1 Then sqlm = sqlm & " AND drama  = 1 "
              If vMovieKategori = 2 Then sqlm = sqlm & " AND komedi = 1 "
              If vMovieKategori = 3 Then sqlm = sqlm & " AND action = 1 "
              If vMovieKategori = 4 Then sqlm = sqlm & " AND horor  = 1 "
              If vMovieKategori = 5 Then sqlm = sqlm & " AND kartun = 1 "
    If cbokategori.ListIndex > 0 Then
        sqlm = sqlm & " AND negara = " & cbokategori.ItemData(cbokategori.ListIndex)
    End If
    sqlm = sqlm & " ORDER BY TITLE LIMIT " & Str(vLstAllRow) & ", " & Str(1) & ";"

    vLstAllRow = vLstAllRow + Lview.ListItems.Count

    Set MyRsm = MyConn.Execute(sqlm)
    If MyRsm.recordCount > 0 Then
        Lview.ListItems.Remove (Lview.ListItems.Count)
        'Isi list data
        Set LV = Lview.ListItems.add(1, , (MyRsm.Fields(0).value))
        LV.SubItems(1) = MyRsm.Fields(1).value
        LV.SubItems(2) = MyRsm.Fields(0).value
    End If

    Set MyRsm = Nothing

        If Lview.ListItems.Count > 0 Then
            Lview.ListItems(1).Selected = True
            Lview.selectedItem.Selected = False
            Set Lview.DropHighlight = Lview.selectedItem
        End If
   lErr = LockWindowUpdate(0)
   DoEvents
End Sub

Sub UpdateChat()
    On Error Resume Next
    Dim Sql As String
    Dim myrs As MYSQL_RS
    Dim LV As ListItem

    If vrekamstate Then
        Exit Sub
    End If


    If Not vVideo = 3 Then
        If Not vVideo = 0 Then
            Exit Sub
        End If
        If flsChatNewMessage.Visible Then
            Exit Sub
        End If
        Dim lokasi As String
        lokasi = App.Path
        flsChatNewMessage.Movie = lokasi + "\picture\anim\newmessage"
        flsChatNewMessage.Visible = True
        Exit Sub
    End If

    Sql = " SELECT chat.pesan, ra.userroom as asal, rt.userroom as tujuan FROM chat inner join room ra on chat.RoomAsal = ra.idroom " & _
          " inner join room rt on chat.roomtujuan = rt.idroom " & _
          " Where chat.roomtujuan = " & Str$(txtCompName.Tag) & _
          " order by chat.id desc limit 0,10"
    Set myrs = MyConn.Execute(Sql)
    myrs.MoveLast
    Do Until myrs.BOF
        If (myrs.Fields(1).value = txtUser.text) Then
            Set LV = lstChat.ListItems.add(, , (myrs.Fields(1).value))
            LV.SubItems(1) = myrs.Fields(0).value
        Else
            Set LV = lstChat.ListItems.add(, , (myrs.Fields(1).value))
            LV.SubItems(1) = myrs.Fields(0).value
        End If
        myrs.MovePrevious
    Loop

    myrs.CloseRecordset
    Sql = "delete from chat where roomtujuan = '" & Str$(txtCompName.Tag) & "';"
    Set myrs = MyConn.Execute(Sql)

    myrs.CloseRecordset
    Sql = "UPDATE room SET CHAT= 0 WHERE IDROOM= '" & Str$(txtCompName.Tag) & "';"
    Set myrs = MyConn.Execute(Sql)
    Set myrs = Nothing
    If lstChat.ListItems.Count > 0 Then
        lstChat.selectedItem.Selected = False
    End If
End Sub

Sub TerimaLagu()

    On Error Resume Next

    Dim LV As ListItem
    Dim Sql As String
    Dim myrs As MYSQL_RS


    Sql = "UPDATE room SET kirimlagu = 0 WHERE IDROOM='" & Trim(Str$(txtCompName.Tag)) & "';"
    MyConn.Execute Sql

    Sql = " SELECT masters.title, masters.singer, masters.IDMUSIC, masters.PATH, masters.ANALOG, masters.VOL FROM room " & _
          " inner join masters on room.kirimidlagu = masters.idmusic where room.roomname ='" & txtCompName & "'"
    Set myrs = MyConn.Execute(Sql)


    Set LV = lstPlaylist.ListItems.add(1, , myrs.Fields(0).value)
    LV.SubItems(1) = myrs.Fields(1).value
    LV.SubItems(2) = myrs.Fields(2).value
    LV.SubItems(3) = myrs.Fields(3).value
    LV.SubItems(4) = myrs.Fields(4).value
    LV.SubItems(5) = myrs.Fields(5).value

    DoEvents

    If PlaySong = True Then

        promoSongPlayCurrent = promoSongPlayCurrent - 1
        cmdStop_Click
    Else
        Unload frmPromo
        PlayLstPlaylist
        If lstPlaylist.ListItems.Count > 0 Then
            lErr = LockWindowUpdate(lstPlaylist.hWnd)
            lstPlaylist.ListItems.Remove (1)
            ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
            lErr = LockWindowUpdate(0)
        End If
    End If

    savePlayList

    ClientRemotePlaylist


lblEnd:

    If Err.Number <> 0 Then
      LogError Me.Name, "TerimaLagu"
    End If
End Sub

Public Sub PrioritasLagu()
    On Error Resume Next
        If lstAll.Visible Then
            If lstAll.ListItems.Count = 0 Then
                Exit Sub
            Else
                cmdRequest_Click
                If vpbBlackBox = 2 Then
                    frmUser.turnDiscoLampOn
                End If
                If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
                    Set lstAll.DropHighlight = lstAll.selectedItem
                    selIndex = lstAll.selectedItem.index
                    lstAll.selectedItem.Selected = False
                End If
            End If
            txtSearch.SelStart = 0
            txtSearch.SelLength = Len(txtSearch.text)
        End If
End Sub

Public Sub Maksimal()
    On Error Resume Next

    UkuranVideo = 1
    Form2.Height = 0

    lblRecording.Top = 360
    lblRecording.Left = 12960
    lblRecording.Height = 585
    lblRecording.Width = 2280

    btnRecStop.Top = lblRecording.Top
    btnRecStop.Left = lblRecording.Left
    btnRecStop.Height = lblRecording.Height
    btnRecStop.Width = lblRecording.Width

    If vVideo = 0 Then
        picKeyTempo.Visible = False
        Picture3.Visible = False

        frmTransparent.Show
        txtCategory.Visible = True
        lstAll.Top = 1920
        lstPlaylist.Top = lstAll.Top
        txtSearch.Top = 380
        flsTitle.Top = 90
        lstAll.Height = 4695
        lstPlaylist.Height = lstAll.Height
        txtVol.Top = 6120
        txtVol.Left = 14160
        txtVol.Height = 645
        txtVol.Width = 900

        txtUser.Top = 520
        txtUser.Left = 6960
        txtUser.Height = 0
        txtUser.Width = 0
        txtTime.Top = 520
        txtTime.Left = 6960
        txtTime.Height = 0
        txtTime.Width = 0

        If vpointer = 3 Then
            frmTransparent.GantiSkin (5)
        Else
            If vpbHits = True Then
                frmTransparent.GantiSkin (2)
            ElseIf vpbNew = True Then
                frmTransparent.GantiSkin (3)
            ElseIf vpbPopuler = True Then
                frmTransparent.GantiSkin (4)
            Else
                frmTransparent.GantiSkin (0)
            End If
        End If

        Dim i As Integer
        For i = 0 To 4
            KotakTambah(i).Width = 400
            KotakKurang(i).Width = 400
            KotakTambah(i).Top = 480
            KotakKurang(i).Top = 480
        Next i
        KotakTambah(0).Left = 5880
        KotakKurang(0).Left = 4340
        For i = 1 To 4
            KotakTambah(i).Left = KotakTambah(i - 1).Left + KotakTambah(i - 1).Width
            KotakKurang(i).Left = KotakKurang(i - 1).Left - KotakKurang(i - 1).Width
        Next i

        lokasi = App.Path + "\picture\normalscreen\"
        flsLogo.Movie = lokasi + "logo"
        flsLogo.Top = 0
        flsLogo.Width = 1875
        flsLogo.Height = 1240
        flsLogo.Left = 0
        flsTitle.Visible = True
        txtRemoteCode.Visible = modProject.frmRoomRemoteCode
        frmRoom.Show
        frmRoom.Height = 11520
    End If

    If vVideo = 5 Then
        txtVol.Top = 4860
        txtVol.Left = 13580
        txtVol.Height = 645
        txtVol.Width = 900
        txtSearch.Top = 4920
        txtCategory.Visible = True

        txtUser.Top = 3240
        txtUser.Left = 7560
        txtUser.Height = 0
        txtUser.Width = 0
        txtTime.Top = 3240
        txtTime.Left = 7560
        txtTime.Height = 0
        txtTime.Width = 0

        frmTransparent.Show
        cbokategori.Visible = False
        lstMovie.Top = 1920

        frmTransparent.GantiSkinMovie (0)
        flsLogo.Movie = App.Path + "\picture\movie\logo.swf"

        flsMovieCategory.Visible = True

        txtSearch.Top = 380
        flsTitle.Top = 90
        txtVol.Top = 6120
        txtVol.Left = 14160
        txtVol.Height = 645
        txtVol.Width = 900

        flsLogo.Top = 0
        flsLogo.Width = 1875
        flsLogo.Height = 1240
        flsLogo.Left = 0
        frmRoom.Height = 11520

        frmTransparent.Show
        txtRemoteCode.Visible = False
        frmRoom.Show
    End If

    If vVideo = 7 Then
        frmTransparent.GantiSkinTV (0)
        flsLogo.Movie = App.Path + "\picture\tv\logo.swf"

        txtUser.Top = 3240
        txtUser.Left = 7560
        txtUser.Height = 0
        txtUser.Width = 0
        txtTime.Top = 3240
        txtTime.Left = 7560
        txtTime.Height = 0
        txtTime.Width = 0

        txtVol.Top = 6120
        txtVol.Left = 14160
        txtVol.Height = 645
        txtVol.Width = 900

        lstTV.Top = 1920

        flsLogo.Top = 0
        flsLogo.Width = 1875
        flsLogo.Height = 1240
        flsLogo.Left = 0
        frmRoom.Height = 11520

        frmTransparent.Show
        frmRoom.Show

        lstTV.SetFocus
    End If
End Sub

Public Sub Minimal()
    On Error Resume Next

    UkuranVideo = 2

    If (vrekamstate) Then
        frmCamera.Show
    End If

    If vVideo = 0 Then
        Form2.Height = 900
        Form2.Show
        picKeyTempo.Visible = False
        frmRoom.Height = 1450
        frmTransparent.GantiSkin (1)

        txtUser.Top = 220
        txtUser.Left = 13800
        txtUser.Height = 270
        txtUser.Width = 1575
        txtTime.Top = 600
        txtTime.Left = 13800
        txtTime.Height = 270
        txtTime.Width = 1575

        txtSearch.Top = Screen.Height + 8000
        flsTitle.Top = Screen.Height + 8000
        txtVol.Top = Screen.Height + 8000
        lstAll.Top = Screen.Height + 8000 '2120
        lstPlaylist.Top = lstAll.Top

        txtRemoteCode.Visible = False

        lokasi = App.Path + "\picture\anim\"
        flsLogo.Movie = lokasi + "mainmin"
        flsLogo.Top = 120
        flsLogo.Width = 1560
        flsLogo.Height = 840
        flsLogo.Left = 360
        Picture3.Visible = True

        lblRecording.Top = 1440
        lblRecording.Left = 12960
        lblRecording.Height = 585
        lblRecording.Width = 2280

        btnRecStop.Top = lblRecording.Top
        btnRecStop.Left = lblRecording.Left
        btnRecStop.Height = lblRecording.Height
        btnRecStop.Width = lblRecording.Width

        txtCategory.Visible = False

        Dim i As Integer
        For i = 0 To 4
            KotakTambah(i).Width = 400
            KotakKurang(i).Width = 400
            KotakTambah(i).Top = 80
            KotakKurang(i).Top = 80
        Next i
        KotakTambah(0).Left = 3180
        KotakKurang(0).Left = 1630
        For i = 1 To 4
            KotakTambah(i).Left = KotakTambah(i - 1).Left + KotakTambah(i - 1).Width
            KotakKurang(i).Left = KotakKurang(i - 1).Left - KotakKurang(i - 1).Width
        Next i

        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.text)
        If vpbfrmCamera Then
            frmCamera.Show
        End If
        frmTransparent.Show
        frmRoom.Show
    End If

    If vVideo = 5 Then
        Form2.Height = 1095
        Form2.Show
        frmRoom.Height = 1450

        lokasi = App.Path + "\picture\anim\"
        flsLogo.Movie = lokasi + "mainmin"

        flsMovieCategory.Visible = False

        flsLogo.Top = 120
        flsLogo.Width = 1560
        flsLogo.Height = 840
        flsLogo.Left = 360

        txtUser.Top = 220
        txtUser.Left = 13800
        txtUser.Height = 270
        txtUser.Width = 1575
        txtTime.Top = 600
        txtTime.Left = 13800
        txtTime.Height = 270
        txtTime.Width = 1575

        frmTransparent.GantiSkinMovie (1)

        txtSearch.Top = frmRoom.Height + 8000
        flsTitle.Top = frmRoom.Height + 8000
        txtVol.Top = frmRoom.Height + 8000

        txtRemoteCode.Visible = False

        txtCategory.Visible = False
        lstMovie.Top = 6825

        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.text)
        frmTransparent.Show
        frmRoom.Show
    End If

    If vVideo = 7 Then
        Form2.Height = 1095
        Form2.Show
        frmRoom.Height = 1450

        txtUser.Top = 220
        txtUser.Left = 13800
        txtUser.Height = 270
        txtUser.Width = 1575
        txtTime.Top = 600
        txtTime.Left = 13800
        txtTime.Height = 270
        txtTime.Width = 1575

        flsLogo.Movie = App.Path + "\picture\anim\mainmin"
        frmTransparent.GantiSkinTV (1)

        flsLogo.Top = 120
        flsLogo.Width = 1560
        flsLogo.Height = 840
        flsLogo.Left = 360

        txtVol.Top = frmRoom.Height + 8000
        flsTitle.Top = frmRoom.Height + 8000

        txtRemoteCode.Visible = False

        frmTransparent.Show
        frmRoom.Show
        lstTV.SetFocus
    End If

    If frmRoom.vVideoAktif Then
        frmVideo.Top = 0
        frmVideo.Left = 0
        frmVideo.Width = 15360
        frmVideo.Height = 11520
        frmVideo.WindowsMediaPlayer1.Top = 0
        frmVideo.WindowsMediaPlayer1.Left = 0
        frmVideo.WindowsMediaPlayer1.Height = frmVideo.Height
        frmVideo.WindowsMediaPlayer1.Width = frmVideo.Width
    End If

End Sub

Public Sub MinimalKey()
    On Error Resume Next
    If UkuranVideo = 2 Or UkuranVideo = 4 Then
        If vVideo = 0 Then
            picKeyTempo.Top = 400
            picKeyTempo.Left = 5040
            picKeyTempo.Height = 480
            picKeyTempo.Width = 5280
            lokasi = App.Path
            picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\keytempomin.jpg")
            frmTransparent.GantiSkin (9)
            Picture3.Visible = False
        End If
    ElseIf UkuranVideo = 1 Then
        flsTitle.Visible = False
        txtRemoteCode.Visible = False
        picKeyTempo.Top = 0
        picKeyTempo.Left = 2340
        picKeyTempo.Height = 1100
        picKeyTempo.Width = 9555
        lokasi = App.Path
        picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\keytempomain.jpg")
        frmTransparent.GantiSkin (7)
        txtCategory.Visible = False
    End If
End Sub

Public Sub MinimalTempo()
    On Error Resume Next
    If UkuranVideo = 2 Or UkuranVideo = 4 Then
        If vVideo = 0 Then
            picKeyTempo.Top = 400
            picKeyTempo.Left = 5040
            picKeyTempo.Height = 480
            picKeyTempo.Width = 5280
            lokasi = App.Path
            picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\keytempomin.jpg")
            frmTransparent.GantiSkin (10)
            Picture3.Visible = False
        End If
    ElseIf UkuranVideo = 1 Then
        flsTitle.Visible = False
        txtRemoteCode.Visible = False
        picKeyTempo.Top = 0
        picKeyTempo.Left = 2340
        picKeyTempo.Height = 1100
        picKeyTempo.Width = 9555
        lokasi = App.Path
        picKeyTempo.Picture = LoadPicture(lokasi + "\Picture\normalscreen\keytempomain.jpg")
        frmTransparent.GantiSkin (8)
        txtCategory.Visible = False
    End If
End Sub

Public Sub MinimalVolume()
    On Error Resume Next
    If UkuranVideo = 2 Or UkuranVideo = 3 Then
        UkuranVideo = 4
        If vVideo = 0 Then
            txtSearch.Top = Screen.Height + 8000
            lstAll.Height = 1695
            lstPlaylist.Height = lstAll.Height
            picKeyTempo.Visible = False
            frmTransparent.GantiSkin (6)
            txtVol.Top = 200
            txtVol.Left = 13800
            txtVol.Height = 645
            txtVol.Width = 1260
        End If
        If vVideo = 5 Then
            txtSearch.Top = Screen.Height + 8000
            frmTransparent.GantiSkin (6)
            txtVol.Top = 200
            txtVol.Left = 13800
            txtVol.Height = 645
            txtVol.Width = 1260
        End If
        If vVideo = 7 Then
            frmTransparent.GantiSkin (6)
            txtVol.Top = 200
            txtVol.Left = 13800
            txtVol.Height = 645
            txtVol.Width = 1260
        End If
        If vVideo = 1 Then
            Skin1.ApplySkinByName hWnd, "volume"
           ' hot
            txtVol.Top = 5520
            txtVol.Left = 7680
            txtVol.Height = 645
            txtVol.Width = 900
        End If
    End If
End Sub

Public Sub PlayLagu()

    On Error Resume Next

    If lstAll.Visible Then
        If lstAll.ListItems.Count = 0 Then
            GoTo lblEnd
        Else
            cmdRequest_Click
            If vpbBlackBox = 2 Then
                frmUser.turnDiscoLampOn
            End If
            If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
                Set lstAll.DropHighlight = lstAll.selectedItem
                selIndex = lstAll.selectedItem.index
                lstAll.selectedItem.Selected = False
            End If
        End If
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.text)
    End If
    ScreenSaverAktif = 5

lblEnd:

    If Err.Number <> 0 Then
      LogError Me.Name, "PlayLagu"
    End If

End Sub

Public Sub prcPageDownLstAll()
    On Error Resume Next

    Dim idMusic As Long
    Dim idMusic2 As Long
    Dim i As Integer
    Dim awal As Boolean

    If lstAll.ListItems.Count < 9 Then
        Exit Sub
    End If
    idMusic = lstAll.ListItems(lstAll.ListItems.Count).SubItems(3)

    lErr = LockWindowUpdate(lstAll.hWnd)
    awal = True
    i = 1
    Do While Not ((i > 8) Or (idMusic = lstAll.ListItems(1).SubItems(3)) Or (idMusic2 = lstAll.ListItems(1).SubItems(3)))
        TambahLstAll
        If awal Then
            idMusic2 = lstAll.ListItems(lstAll.ListItems.Count).SubItems(3)
            awal = False
        End If
        i = i + 1
    Loop
    lErr = LockWindowUpdate(0)
End Sub

Public Sub prcPageUpLstAll()
    On Error Resume Next

    Dim idMusic As Long
    Dim i As Integer

    If lstAll.ListItems.Count < 9 Then
        Exit Sub
    End If

    idMusic = lstAll.ListItems(1).SubItems(3)
    lErr = LockWindowUpdate(lstAll.hWnd)
    i = 1
    Do While Not ((i > 8) Or (idMusic = lstAll.ListItems(lstAll.ListItems.Count).SubItems(3)))
        KurangLstAll
        i = i + 1
    Loop
    lErr = LockWindowUpdate(0)
End Sub

Public Sub prcDownLstAll()
    On Error Resume Next
    If lstAll.Visible Then
        If lstAll.ListItems.Count = 0 Then
            Exit Sub
        End If
        If selIndex < 9 Then
            If selIndex >= lstAll.ListItems.Count Then
                selIndex = lstAll.ListItems.Count
            Else
                selIndex = selIndex + 1
            End If
            lstAll.ListItems(selIndex).Selected = True
        Else
            TambahLstAll
        End If
        'LEWATKAN KALAU HURUF CINA
        If lstAll.ListItems(selIndex).SubItems(7) = "1" Then
            If selIndex < 9 Then
                If selIndex >= lstAll.ListItems.Count Then
                    selIndex = lstAll.ListItems.Count
                Else
                    selIndex = selIndex + 1
                End If
                lstAll.ListItems(selIndex).Selected = True
            Else
                TambahLstAll
            End If
        End If
        If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
                Set lstAll.DropHighlight = lstAll.selectedItem
                selIndex = lstAll.selectedItem.index
                lstAll.selectedItem.Selected = False
        End If
    End If
End Sub

Public Sub prcUplstAll()
    On Error Resume Next
    If lstAll.Visible Then
        If lstAll.ListItems.Count = 0 Then
            Exit Sub
        End If
        If selIndex > 1 Then
            selIndex = selIndex - 1
            lstAll.ListItems(selIndex).Selected = True
        Else
            KurangLstAll
        End If
        If lstAll.ListItems(selIndex).SubItems(7) = "1" Then
            If selIndex > 1 Then
                selIndex = selIndex - 1
                lstAll.ListItems(selIndex).Selected = True
            Else
                KurangLstAll
            End If
        End If
        If (lstAll.ListItems.Count > 0) And (Not lstAll.selectedItem Is Nothing) Then
                Set lstAll.DropHighlight = lstAll.selectedItem
                selIndex = lstAll.selectedItem.index
                lstAll.selectedItem.Selected = False
        End If
    End If
End Sub

Public Sub prcLoginMember()
    On Error Resume Next
    frmConfirmasi.vpengirim = 0
    frmConfirmasi.Text1 = ""
    frmConfirmasi.Text2 = ""
    frmConfirmasi.text3 = ""
    frmTransparent.Show
    frmRoom.Show
    frmConfirmasi.Show
    frmConfirmasi.Text1.SetFocus
    frmConfirmasi.Top = 1550
End Sub

Public Sub prcAddMemberPlaylist()

    On Error Resume Next

    If vpbMember = "" Then
        prcLoginMember
        Exit Sub
    End If

    If (lstAll.ListItems.Count = 0) Then
        Exit Sub
    End If

    If UkuranVideo <> 1 Then
        ScreenSaverAktif = 0
        Maksimal
        Exit Sub
    End If

    If lstAll.Visible = False Then
        frmRoom.ScreenSaverAktif = 0
        frmRoom.Maksimal
        prcSonglist
        Exit Sub
    End If

    cmdRequest_Click
    DoEvents

    frmCall.vpengirim = 0
    frmCall.Show
    frmCall.Timer1.Enabled = True

    If (lstAll.Visible = True) Then
        Dim Sql As String
        If (lstAll.ListItems(selIndex).SubItems(3) <> "") Then
            Sql = "INSERT INTO memberplaylist (idmember, idmusic) " & _
                  "VALUES('" & Trim(vpbMember) & "','" & lstAll.ListItems(selIndex).SubItems(3) & "');"
            MyConn.Execute Sql
        End If

    End If


    If Err.Number <> 0 Then
      LogError Name, "prcAddMemberPlaylist"
    End If
End Sub

Public Sub prcDeleteMemberPlaylist()
    On Error Resume Next
    If vpbMember = "" Then
        prcLoginMember
        Exit Sub
    End If

    If vpointer <> 3 Then
        frmRoom.ScreenSaverAktif = 0
        If vVideo <> 1 Then
            frmRoom.Maksimal
            DoEvents
        End If
        frmRoom.pencet_btnplaylist
        Exit Sub
    End If

    frmCall.vpengirim = 1
    frmCall.Show
    frmCall.Timer1.Enabled = True

    Dim Sql As String
    If (vpointer = 3) And (lstPlaylist.ListItems.Count > 0) Then
        Sql = "DELETE FROM  memberplaylist " & _
              "WHERE (idmember ='" & Trim(vpbMember) & "') AND (idmusic = '" & lstPlaylist.selectedItem.SubItems(2) & "');"
        MyConn.Execute Sql
        frmRoom.cmdDelete_Click
    End If
End Sub

Public Sub prcClearMemberPlaylist()
    On Error Resume Next
    If vpbMember = "" Then
        prcLoginMember
        Exit Sub
    End If

'    If vpointer <> 3 Then
'        frmRoom.ScreenSaverAktif = 0
'        frmRoom.Maksimal
'        frmRoom.pencet_btnplaylist
'        DoEvents
'    End If

    frmCall.vpengirim = 2
    frmCall.Show
    frmCall.Timer1.Enabled = True

    Dim Sql As String
    Sql = "DELETE FROM  memberplaylist " & _
          "WHERE (idmember ='" & Trim(vpbMember) & "');"
    MyConn.Execute Sql

'    cmdClear_Click
End Sub

Public Sub prcHilangkanScrollbarPlaylist()
    On Error Resume Next
    ShowScrollBar lstPlaylist.hWnd, SB_VERT, False
    lErr = LockWindowUpdate(0)
End Sub


Private Sub ServerRemotePlay(idMusic As String)

  On Error Resume Next

  Dim PathLagu As String
  Dim PathLaguUtama As String
  Dim PathLaguBackup As String
  Dim ECHO As ICMP_ECHO_REPLY

  If (vVideo <> 0) And (vVideo <> 2) Then
    PlayerKomputer
    Unload frmCamera
    DoEvents
  End If

  If vpbBlackBox = 2 Then
      frmUser.turnDiscoLampOn
  End If

  frmVideo.WindowsMediaPlayer1.URL = ""
  frmVideo.WindowsMediaPlayer1.Controls.stop
  Unload frmVideo

  frmVideo.Show

  frmVideo.WindowsMediaPlayer1.Visible = True

  Dim Sql As String
  Dim rs As MYSQL_RS
  Sql = "SELECT PATH,ANALOG,VOL FROM masters WHERE IDMUSIC=" & idMusic
  Set rs = MyConn.Execute(Sql)

  PathLagu = Replace$(rs.Fields(0).value, "/", "\")
  If AktifServerStatus = 1 Then
      Call Ping(vpbServerBackup, ECHO)
      If ECHO.status = 0 Then
          PathLaguUtama = "\\" & vpbServerBackup & "\Data\" & PathLagu
      Else
          PathLaguUtama = "\\" & vpbServerUtama & "\Data\" & PathLagu
      End If
          PathLaguBackup = "\\" & vpbServerUtama & "\Data\" & PathLagu
  Else
      Call Ping(vpbServerUtama, ECHO)
      If ECHO.status = 0 Then
          PathLaguUtama = "\\" & vpbServerUtama & "\Data\" & PathLagu
      Else
          PathLaguUtama = "\\" & vpbServerBackup & "\Data\" & PathLagu
      End If
          PathLaguBackup = "\\" & vpbServerBackup & "\Data\" & PathLagu
  End If

  If FileExists(PathLaguUtama) Then
      frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
  ElseIf FileExists(PathLaguBackup) Then
      frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
  Else
      rs.CloseRecordset
      Set rs = Nothing
      GoTo lblEnd
  End If

  frmVideo.WindowsMediaPlayer1.Controls.play

  frmVideo.WindowsMediaPlayer1.settings.volume = rs.Fields("VOL").value

  Select Case rs.Fields("ANALOG").value
    Case "ML"
      tmrVokal.Enabled = False
      tmrNonVocalML.Enabled = True
      tmrNonVocalMR.Enabled = False
      vVocal = 2
    Case "MR"
      tmrVokal.Enabled = False
      tmrNonVocalML.Enabled = False
      tmrNonVocalMR.Enabled = True
      vVocal = 3
    Case "ST"
      tmrVokal.Enabled = False
      tmrNonVocalML.Enabled = False
      tmrNonVocalMR.Enabled = False
      vVocal = 1
  End Select
  vVolvocal = 0

  rs.CloseRecordset
  Set rs = Nothing

  ScoreValid = True
  PlaySong = True


lblEnd:

  If Err.Number <> 0 Then
    LogError Me.Name, "ServerRemotePlay"
  End If
End Sub

Private Sub ServerRemoteRemove(playOrder As Long)

  On Error Resume Next

  lstPlaylist.ListItems.Remove playOrder

  savePlayList

  If lstPlaylist.ListItems.Count >= playOrder Then
    Set lstPlaylist.selectedItem = lstPlaylist.ListItems(playOrder)
    Set lstPlaylist.DropHighlight = lstPlaylist.ListItems(playOrder)
  Else
    Set lstPlaylist.selectedItem = lstPlaylist.ListItems(lstPlaylist.ListItems.Count)
    Set lstPlaylist.DropHighlight = lstPlaylist.ListItems(lstPlaylist.ListItems.Count)
  End If

  If lstPlaylist.ListItems.Count = 0 Then
    txtNextSong.text = "NO NEXT SONG"
  Else
    txtNextSong.text = lstPlaylist.ListItems.Item(1)
  End If

  DoEvents

  ClientRemotePlaylist

lblEnd:

  If Err.Number <> 0 Then
    LogError Me.Name, "ServerRemoteRemove"
  End If

End Sub

Sub ClientRemotePlaylist()
      On Error Resume Next
  Dim a As Long

  For a = 1 To ClientRemoteIp.Count
    wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
    wsClientRemote.SendData "playlist"
  Next
End Sub

Sub ClientRemoteEffect()
      On Error Resume Next
  Dim effect As String
  Dim a As Long

  If ClientRemoteIp.Count > 0 Then

    If vEffectAmpli = 1 Or vEffectAmpli = 5 Or vEffectAmpli = 9 Then
      effect = "1"
    ElseIf vEffectAmpli = 2 Or vEffectAmpli = 6 Or vEffectAmpli = 10 Then
      effect = "2"
    ElseIf vEffectAmpli = 3 Or vEffectAmpli = 7 Or vEffectAmpli = 11 Then
      effect = "3"
    ElseIf vEffectAmpli = 4 Or vEffectAmpli = 8 Or vEffectAmpli = 12 Then
      effect = "4"
    Else
      effect = "2"
    End If

    For a = 1 To ClientRemoteIp.Count
      wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
      wsClientRemote.SendData "effect=" & effect
    Next
  End If
End Sub

Sub ClientRemoteVocal()
      On Error Resume Next
  Dim value As String
  Dim a As Long

  If ClientRemoteIp.Count > 0 Then

    If vVocalterus = True Then
      value = "1"
    Else
      value = "0"
    End If

    For a = 1 To ClientRemoteIp.Count
      wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
      wsClientRemote.SendData "vocal=" & value
    Next
  End If
End Sub

Sub ClientRemoteLamp()
    On Error Resume Next
  Dim a As Long

  For a = 1 To ClientRemoteIp.Count
    wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
    wsClientRemote.SendData "lamp=" & vpbBlackBox
  Next
End Sub

Sub ClientRemoteScore()
    On Error Resume Next
  Dim value As String
  Dim a As Long

  If ClientRemoteIp.Count > 0 Then

    If ScoreSetup = True Then
      value = "1"
    Else
      value = "0"
    End If

    For a = 1 To ClientRemoteIp.Count
      wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
      wsClientRemote.SendData "score=" & value
    Next
  End If
End Sub

Private Sub ClientRemoteVolume()
    On Error Resume Next

  Dim a As Long

  For a = 1 To ClientRemoteIp.Count
    wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
    wsClientRemote.SendData "volume=" & txtVol.text
  Next
End Sub

Private Sub ClientRemoteLampScore()

  On Error Resume Next

  Dim value As String
  Dim a As Long

  If ClientRemoteIp.Count > 0 Then

    If ScoreSetup = True Then
      value = "1"
    Else
      value = "0"
    End If

    For a = 1 To ClientRemoteIp.Count
      wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
      wsClientRemote.SendData "lampScore=" & vpbBlackBox & ";" & value
    Next
  End If
End Sub

Private Sub ClientRemoteEffectVocalVolume()

  On Error Resume Next

  Dim effect As String
  Dim a As Long

  If ClientRemoteIp.Count > 0 Then

    If vEffectAmpli = 1 Or vEffectAmpli = 5 Or vEffectAmpli = 9 Then
      effect = "1"
    ElseIf vEffectAmpli = 2 Or vEffectAmpli = 6 Or vEffectAmpli = 10 Then
      effect = "2"
    ElseIf vEffectAmpli = 3 Or vEffectAmpli = 7 Or vEffectAmpli = 11 Then
      effect = "3"
    ElseIf vEffectAmpli = 4 Or vEffectAmpli = 8 Or vEffectAmpli = 12 Then
      effect = "4"
    Else
      effect = "2"
    End If

    Dim vocal As String
    If vVocalterus = True Then
      vocal = "1"
    Else
      vocal = "0"
    End If

    For a = 1 To ClientRemoteIp.Count
      wsClientRemote.RemoteHost = ClientRemoteIp.Item(a)
      wsClientRemote.SendData "effectVocalVolume=" & effect & ";" & vocal & ";" & txtVol.text
    Next
  End If
End Sub


Private Sub wsServerRemote_DataArrival(ByVal bytesTotal As Long)

  On Error Resume Next

  Dim Data As String

  Dim idMember As String
  Dim pair() As String
  Dim idPlaylist As String
  Dim idMusic As String
  Dim Sql As String
  Dim a As Long
  Dim playOrder As Long
  Dim rs As MYSQL_RS
  Dim li As ListItem

  Dim duplicateLastPlaylist As Boolean

  Dim command() As String

  wsServerRemote.GetData Data

  command = Split(Data, "=")

  If command(0) <> txtRemoteCode.text Then
    GoTo lblEnd
  End If

  For a = 1 To ClientRemoteIp.Count
    If ClientRemoteIp.Item(a) = wsServerRemote.RemoteHostIP Then
      Exit For
    End If
  Next
  If a > ClientRemoteIp.Count Then
    ClientRemoteIp.add wsServerRemote.RemoteHostIP
  End If

  ScreenSaverAktif = 5

  Select Case command(1)

    Case "call"

      frmCall.vpengirim = 3
      frmCall.Show
      frmCall.Timer1.Enabled = True


    Case "clear"

      If txtLogin.text = "" Then
        idMember = "00000"
      Else
        idMember = txtLogin.text
      End If

      Sql = "DELETE FROM playlist WHERE USERID='" & idMember & "' AND room='" & txtCompName.text & "'"
      MyConn.Execute Sql
      If MyConnBackup.State = MY_CONN_OPEN Then
        MyConnBackup.Execute Sql
      End If

      lstPlaylist.ListItems.Clear
      txtNextSong.text = "NO NEXT SONG"


    Case "effect"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      Unload frmConfirmasi

      vEffectAmpli = command(2)
      For a = 1 To 5
        RemoteAmpli "Preset" & vEffectAmpli, 1
        Sleep 3
      Next

      frmConfirmasi.vpengirim = 12
      frmConfirmasi.Show

      ClientRemoteEffect


    Case "getLamp"
      ClientRemoteLamp


    Case "getScore"
      ClientRemoteScore


    Case "getVolume"
      ClientRemoteVolume


    Case "getEffectVocalVolume"
      ClientRemoteEffectVocalVolume


    Case "keyDown"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      MinimalKey
      prckeyDown

    Case "keyUp"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      MinimalKey
      prcKeyUp


    Case "lampOff"

      vpbBlackBox = 0
      frmUser.turnDiscoLampOff

    Case "lampOn"

      vpbBlackBox = 2
      frmUser.turnDiscoLampOn


    Case "micDown"

      RemoteAmpli "MICDN", 0

    Case "micUp"

      RemoteAmpli "MICUP", 0


    Case "pause"

      If cmdPause(0).Visible = True Then
        cmdPause_Click 0
      Else
        cmdPause_Click 1
      End If


    Case "pauseMovie"

      If cmdPause(1).Visible = True Then
          cmdPause_Click 1
      Else
          cmdPause_Click 0
      End If


    Case "play"

      If command(2) = "-1" Then
        If cmdPause(1).Visible = True Then
          cmdPause_Click 1
        End If
      Else
        ServerRemotePlay command(2)
      End If


    Case "playPlaylist"

      If command(2) = "-1:-1" Then
        If cmdPause(1).Visible = True Then
          cmdPause_Click 1
        End If
      Else

        pair = Split(command(2), ":")

        ServerRemotePlay pair(0)

        ServerRemoteRemove CLng(pair(1)) + 1
      End If


    Case "playMovie"

      If command(2) = "-1" Then
        If cmdPause(1).Visible = True Then
          cmdPause_Click 1
        End If
      Else

        If vVideo <> 5 Then
          moviestart
          DoEvents
        End If

        Dim PathLagu As String
        Dim PathLaguUtama As String
        Dim PathLaguBackup As String
        Dim ECHO As ICMP_ECHO_REPLY

        Maksimal

        vTempo = 0
        vKey = 0

        frmUser.turnDiscoLampOff

        Sql = "SELECT PATH, VOL FROM film Where ID = '" & command(2) & "';"
        Set rs = MyConn.Execute(Sql)

        PathLagu = Replace$(rs.Fields(0).value, "/", "\")
        If AktifServerStatus = 1 Then
            Call Ping(vpbServerBackup, ECHO)
            If ECHO.status = 0 Then
                PathLaguUtama = "\\" & vpbServerBackup & "\Movie\" & PathLagu
            Else
                PathLaguUtama = "\\" & vpbServerUtama & "\Movie\" & PathLagu
            End If
                PathLaguBackup = "\\" & vpbServerUtama & "\Movie\" & PathLagu
        Else
            Call Ping(vpbServerUtama, ECHO)
            If ECHO.status = 0 Then
                PathLaguUtama = "\\" & vpbServerUtama & "\Movie\" & PathLagu
            Else
                PathLaguUtama = "\\" & vpbServerBackup & "\Movie\" & PathLagu
            End If
                PathLaguBackup = "\\" & vpbServerBackup & "\Movie\" & PathLagu
        End If

        Unload frmPromo
        Unload frmVideo
        frmVideo.Show

        If FileExists(PathLaguUtama) Then
            ScoreValid = False
            tmrVokal.Enabled = False
            tmrNonVocalML.Enabled = False
            tmrNonVocalMR.Enabled = False
            frmVideo.WindowsMediaPlayer1.URL = ""
            frmVideo.WindowsMediaPlayer1.Controls.stop
            frmVideo.WindowsMediaPlayer1.Visible = True
            PlaySong = False
            frmVideo.WindowsMediaPlayer1.URL = PathLaguUtama
            frmVideo.WindowsMediaPlayer1.Controls.play
            PlaySong = True
            frmVideo.WindowsMediaPlayer1.settings.volume = rs.Fields(1).value
        ElseIf FileExists(PathLaguBackup) Then
            ScoreValid = False
            tmrVokal.Enabled = False
            tmrNonVocalML.Enabled = False
            tmrNonVocalMR.Enabled = False
            frmVideo.WindowsMediaPlayer1.URL = ""
            frmVideo.WindowsMediaPlayer1.Controls.stop
            frmVideo.WindowsMediaPlayer1.Visible = True
            PlaySong = False
            frmVideo.WindowsMediaPlayer1.URL = PathLaguBackup
            frmVideo.WindowsMediaPlayer1.Controls.play
            PlaySong = True
            vtetapfokus = 0
            frmVideo.WindowsMediaPlayer1.settings.volume = rs.Fields(1).value
        End If

        vtetapfokus = 0

        Set rs = Nothing

        frmTransparent.Show

        If (frmRoom.vpointer = 1) Or (frmRoom.vpointer = 2) Or (frmRoom.vpointer = 6) Then
            sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
        End If
      End If


    Case "playTelevision"

      If command(2) <> "-1" Then

        If vVideo <> 7 Then
          PlayerTV
          DoEvents
        End If

        setAudioEndPointVolumeMasterVolumeLevelPercent 0

        If command(2) = frmCamera.VideoCap1.channel Then
          '---------PAUSE COMP------------
          If PlaySong = True Then
              frmVideo.WindowsMediaPlayer1.Controls.pause
          End If
          Minimal
          ScreenSaverAktif = 99
        Else
          frmCamera.VideoCap1.channel = command(2)
        End If

        VolTemp = 0
        tmrVolume.Enabled = True

      End If


    Case "prior"

      playOrder = command(2) + 1

      Dim title As String
      Dim singer As String
      Dim Path As String
      Dim analog As String
      Dim vol As String

      title = lstPlaylist.ListItems(playOrder).text
      singer = lstPlaylist.ListItems(playOrder).ListSubItems(1).text
      idMusic = lstPlaylist.ListItems(playOrder).ListSubItems(2).text
      Path = lstPlaylist.ListItems(playOrder).ListSubItems(3).text
      analog = lstPlaylist.ListItems(playOrder).ListSubItems(4).text
      vol = lstPlaylist.ListItems(playOrder).ListSubItems(5).text

      lstPlaylist.ListItems.Remove playOrder

      Set li = lstPlaylist.ListItems.add(playOrder - 1, , title)
      li.SubItems(1) = singer
      li.SubItems(2) = idMusic
      li.SubItems(3) = Path
      li.SubItems(4) = analog
      li.SubItems(5) = vol

      Set lstPlaylist.selectedItem = lstPlaylist.ListItems(playOrder - 1)
      Set lstPlaylist.DropHighlight = lstPlaylist.ListItems(playOrder - 1)

      txtNextSong.text = lstPlaylist.ListItems.Item(1)

      DoEvents


      savePlayList

      ClientRemotePlaylist

    Case "recAudio"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      prcCD
      frmConfirmasi.prcKonfirm

    Case "recVideo"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      prcDVD
      frmConfirmasi.prcKonfirm


    Case "remove"

      ServerRemoteRemove CLng(command(2)) + 1


    Case "repeat"

      prcRepeat


    Case "reserve"

      If PlaySong = False Then

        ServerRemotePlay command(2)

      Else

        duplicateLastPlaylist = False
        If lstPlaylist.ListItems.Count > 0 Then
          If command(2) = lstPlaylist.ListItems(lstPlaylist.ListItems.Count).SubItems(2) Then
            duplicateLastPlaylist = True
          End If
        End If

        If duplicateLastPlaylist = False Then

          Sql = "SELECT TITLE,SINGER,PATH,ANALOG,VOL FROM masters WHERE IDMUSIC=" & command(2)
          Set rs = MyConn.Execute(Sql)

          Dim LV As ListItem
          Set LV = lstPlaylist.ListItems.add(, , rs.Fields(0).value)
          LV.SubItems(1) = rs.Fields(1).value
          LV.SubItems(2) = command(2)
          LV.SubItems(3) = rs.Fields(2).value
          LV.SubItems(4) = rs.Fields(3).value
          LV.SubItems(5) = rs.Fields(4).value

          txtNextSong.text = lstPlaylist.ListItems.Item(1)

          DoEvents

          savePlayList

          ClientRemotePlaylist
        End If
      End If


    Case "scoreOff"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      ScoreSetup = False
      frmVocal.VocalAktif = 3
      frmVocal.Show

      tmrMainVocal.Enabled = True

    Case "scoreOn"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      ScoreSetup = True
      frmVocal.VocalAktif = 2
      frmVocal.Show

      tmrMainVocal.Enabled = True


    Case "seekBackward"
      cmdFast_Click

    Case "seekForward"
      cmdSlow_Click


    Case "seekMovieBackward"
      cmdFast_Click

    Case "seekMovieForward"
      cmdSlow_Click


    Case "start"
      frmWelcome.StartKaraoke


    Case "stop"
      cmdStop_Click


    Case "stopMovie"

      cmdStop_Click


    Case "tempoDown"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      MinimalTempo
      prcTempoDown

    Case "tempoUp"

      If (vVideo <> 0) And (vVideo <> 2) Then
        GoTo lblEnd
      End If

      MinimalTempo
      prcTempoUp


    Case "vocal"

      If Not (vpbfrmVocal) Then

        If vpbBolehMainVocal = True Then

          setvocal

          If vVocalterus = False Then
            frmVocal.VocalAktif = 0
            frmVocal.Show
          Else
            frmVocal.VocalAktif = 1
            frmVocal.Show
          End If

          vVocalterus = Not (vVocalterus) 'True

          tmrMainVocal.Enabled = True
        End If
      End If


    Case "volUp"
      setvolumenaik

    Case "volDown"
      setvolumeturun


    Case "volMovUp"
      setvolumenaik

    Case "volMovDown"
      setvolumeturun


    Case "volMic1"

      If vpbAmpliAuto = 2 Then

        Minimal

        Unload frmConfirmasi

        frmConfirmasi.vpengirim = 14
        frmConfirmasi.Show
        frmConfirmasi.Text1.SetFocus

        frmConfirmasi.tmrAktif.Enabled = False
        frmConfirmasi.tmrAktif.Enabled = True

        Select Case vEffectAmpli
          Case 1
            vEffectAmpli = 5
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 2
            vEffectAmpli = 6
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 3
            vEffectAmpli = 7
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 4
            vEffectAmpli = 8
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 9
            vEffectAmpli = 5
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 10
            vEffectAmpli = 6
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 11
            vEffectAmpli = 7
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
          Case 12
            vEffectAmpli = 8
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
        End Select

      End If

    Case "volMic2"

      If vpbAmpliAuto = 2 Then

        Minimal

        Unload frmConfirmasi

        frmConfirmasi.vpengirim = 14
        frmConfirmasi.Show
        frmConfirmasi.Text1.SetFocus

        frmConfirmasi.tmrAktif.Enabled = False
        frmConfirmasi.tmrAktif.Enabled = True

        Select Case vEffectAmpli
          Case 5
            vEffectAmpli = 1
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 6
            vEffectAmpli = 2
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 7
            vEffectAmpli = 3
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 8
            vEffectAmpli = 4
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 9
            vEffectAmpli = 1
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 10
            vEffectAmpli = 2
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 11
            vEffectAmpli = 3
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
          Case 12
            vEffectAmpli = 4
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
        End Select

      End If

    Case "volMic3"

      If vpbAmpliAuto = 2 Then

        Minimal

        Unload frmConfirmasi

        frmConfirmasi.vpengirim = 14
        frmConfirmasi.Show
        frmConfirmasi.Text1.SetFocus

        frmConfirmasi.tmrAktif.Enabled = False
        frmConfirmasi.tmrAktif.Enabled = True

        Select Case vEffectAmpli
          Case 1
            vEffectAmpli = 9
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 2
            vEffectAmpli = 10
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 3
            vEffectAmpli = 11
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 4
            vEffectAmpli = 12
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 5
            vEffectAmpli = 9
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 6
            vEffectAmpli = 10
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 7
            vEffectAmpli = 11
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
          Case 8
            vEffectAmpli = 12
            frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
        End Select

      End If

  End Select


lblEnd:

  If Err.Number <> 0 Then
    LogError Name, "wsServerRemote_DataArrival"
  End If
End Sub

Function getApa() As String

  On Error Resume Next

  Dim rs As MyVbQL.MYSQL_RS

  Set rs = MyConn.Execute("SELECT APA FROM room WHERE ROOMNAME ='" & txtCompName.text & "'")

  getApa = rs.Fields(0).value

  rs.CloseRecordset
  Set rs = Nothing


  If Err.Number <> 0 Then
    LogError Name, "getApa"
  End If
End Function

Sub savePlayList()

  On Error Resume Next

  Dim roomName As String
  Dim idMember As String
  Dim Sql As String
  Dim itemCount As Long
  Dim a As Long

  roomName = txtCompName.text

  If txtLogin = "" Then
      idMember = "00000"
  Else
      idMember = txtLogin.text
  End If


  Sql = "delete from playlist where ROOM='" & roomName & "' and USERID='" & idMember & "'"

  MyConn.Execute Sql
  DoEvents

  If MyConnBackup.State = MY_CONN_OPEN Then
    MyConnBackup.Execute Sql
    DoEvents
  End If


  itemCount = lstPlaylist.ListItems.Count

  If itemCount > 0 Then

    Sql = "INSERT INTO playlist (ROOM, IDMUSIC, USERID, playOrder) VALUES"

    For a = 1 To itemCount

      If a > 1 Then
        Sql = Sql & " ,"
      End If

      Sql = Sql & " ('" & roomName & "', '" & lstPlaylist.ListItems(a).ListSubItems(2).text & "', '" & idMember & "', " & a & ")"
    Next

    DoEvents

    MyConn.Execute Sql
    DoEvents

    If MyConnBackup.State = MY_CONN_OPEN Then
      MyConnBackup.Execute Sql
      DoEvents
    End If
  End If


  If Err.Number <> 0 Then
    LogError Name, "savePlayList"
  End If
End Sub

Sub updateStructure()

  On Error Resume Next

  Dim mysqlConnection As New MYSQL_CONNECTION


  mysqlConnection.OpenConnection vpbServerUtama, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
  DoEvents
  If mysqlConnection.Error.Number = 0 Then
    updateStructure2 mysqlConnection
  End If


  If vpbServerBackup <> vpbServerBackup Then
    mysqlConnection.OpenConnection vpbServerBackup, "karaoke", vpbServerKeyMySQL, "karaoke", "3306"
    DoEvents
    If mysqlConnection.Error.Number = 0 Then
      updateStructure2 mysqlConnection
    End If
  End If
End Sub

Sub updateStructure2(mysql As MYSQL_CONNECTION)

  On Error Resume Next

  Dim database_version As Long
  Dim rs As MyVbQL.MYSQL_RS
  Dim valueString As String
  Dim recordCount As Long
  
  database_version = 0
  Set rs = mysql.Execute("SHOW TABLES WHERE Tables_in_karaoke = 'setting'")
  DoEvents
  If rs.recordCount > 0 Then
    Set rs = Nothing
    Set rs = mysql.Execute("select value from setting where name = 'database_version'")
    DoEvents
    database_version = CLng(rs.Fields("value").value)
  End If
  Set rs = Nothing
  
  If database_version < 1 Then
    mysql.Execute "CREATE TABLE `setting` (`name` varchar(255) NOT NULL, `value` varchar(255) NOT NULL, PRIMARY KEY (`name`)) ENGINE=InnoDB DEFAULT CHARSET=utf8"
    DoEvents
    mysql.Execute "insert setting set name = 'database_version', value = 1"
    DoEvents
  End If
  
  If database_version < 2 Then
    Set rs = mysql.Execute("show index from masters where Key_name <> 'PRIMARY'")
    DoEvents
    rs.MoveFirst
    While rs.EOF = False
      mysql.Execute "Alter table masters drop index " & rs.Fields("Key_name").value
      DoEvents
      rs.MoveNext
    Wend
    rs.CloseRecordset
    Set rs = Nothing
    mysql.Execute "Alter table masters add index (`TITLE3`), add index (`TITLE5`), add index (`TITLE6`), add index (`TITLE7`), add index (`SINGER3`), add index (`SINGER5`), add index (`SINGER6`), add index (`SINGER7`), add index (`TYPE`), add index (`HITS`), add index (`NEW`), add index (`POPULER`)"
    DoEvents
    mysql.Execute "DROP PROCEDURE IF EXISTS `spAbrSingerK`"
    DoEvents
    mysql.Execute "DROP PROCEDURE IF EXISTS `spAbrTitleK`"
    DoEvents
    mysql.Execute "DROP PROCEDURE IF EXISTS `spSingerK`"
    DoEvents
    mysql.Execute "DROP PROCEDURE IF EXISTS `spTitleK`"
    DoEvents
    mysql.Execute "CREATE DEFINER=`karaoke`@`%` PROCEDURE `spAbrSingerK`(IN _CARI char(50), IN _KATEGORI INTEGER, IN _START INTEGER, IN _LIMIT INTEGER) BEGIN PREPARE STMT FROM 'SELECT TITLE, SINGER, IDMUSIC, PATH, ANALOG, VOL, TITLE2, SINGER2, CODE, TITLE3, SINGER3, TITLE4, SINGER4 FROM masters where (SINGER6 like ? OR SINGER7 like ?) AND TYPE = ? order by SINGER3 ASC, TITLE3 ASC LIMIT ?,?'; SET @CARI = _CARI; SET @CARI2 = _CARI; SET @KATEGORI = _KATEGORI; SET @START = _START; SET @LIMIT = _LIMIT; EXECUTE STMT USING @CARI, @CARI2, @KATEGORI, @START, @LIMIT; END"
    DoEvents
    mysql.Execute "CREATE DEFINER=`karaoke`@`%` PROCEDURE `spAbrTitleK`(IN _CARI char(50), IN _KATEGORI INTEGER, IN _START INTEGER, IN _LIMIT INTEGER) BEGIN PREPARE STMT FROM 'SELECT TITLE, SINGER, IDMUSIC, PATH, ANALOG, VOL, TITLE2, SINGER2, CODE, TITLE3, SINGER3, TITLE4, SINGER4 FROM masters where (TITLE6 like ? OR TITLE7 like ?) AND TYPE = ? order by TITLE3 ASC, SINGER3 ASC LIMIT ?,?'; SET @CARI = _CARI; SET @CARI2 = _CARI; SET @KATEGORI = _KATEGORI; SET @START = _START; SET @LIMIT = _LIMIT; EXECUTE STMT USING @CARI, @CARI2, @KATEGORI, @START, @LIMIT; END"
    DoEvents
    mysql.Execute "CREATE DEFINER=`karaoke`@`%` PROCEDURE `spSingerK`(IN _CARI char(50), IN _KATEGORI INTEGER, IN _START INTEGER, IN _LIMIT INTEGER) BEGIN PREPARE STMT FROM 'SELECT TITLE, SINGER, IDMUSIC, PATH, ANALOG, VOL, TITLE2, SINGER2, CODE, TITLE3, SINGER3, TITLE4, SINGER4 FROM masters where (SINGER3 like ? OR SINGER5 like ?) AND TYPE = ? order by SINGER3 ASC, TITLE3 ASC LIMIT ?,?'; SET @CARI = _CARI; SET @CARI2 = _CARI; SET @KATEGORI = _KATEGORI; SET @START = _START; SET @LIMIT = _LIMIT; EXECUTE STMT USING @CARI, @CARI2, @KATEGORI, @START, @LIMIT; END"
    DoEvents
    mysql.Execute "CREATE DEFINER=`karaoke`@`%` PROCEDURE `spTitleK`( IN _CARI char(50), IN _KATEGORI INTEGER, IN _START INTEGER, IN _LIMIT INTEGER ) BEGIN PREPARE STMT FROM  'SELECT TITLE, SINGER, IDMUSIC, PATH, ANALOG, VOL, TITLE2, SINGER2, CODE, TITLE3, SINGER3, TITLE4, SINGER4 FROM masters where (TITLE3 like ? OR TITLE5 like ?) AND TYPE = ? order by TITLE3 ASC, SINGER3 ASC LIMIT ?,?'; SET @CARI = _CARI; SET @CARI2 = _CARI; SET @KATEGORI = _KATEGORI; SET @START = _START; SET @LIMIT = _LIMIT; EXECUTE STMT USING @CARI, @CARI2, @KATEGORI, @START, @LIMIT; END"
    DoEvents
    mysql.Execute "update setting set value = 2 where name = 'database_version'"
    DoEvents
  End If
  
  If database_version < 3 Then
    Set rs = mysql.Execute("SHOW COLUMNS FROM room WHERE Field IN ('remoteCode')")
    DoEvents
    recordCount = rs.recordCount
    rs.CloseRecordset
    Set rs = Nothing
    If recordCount <> 1 Then
      mysql.Execute "ALTER TABLE room ADD COLUMN remoteCode INT(10) UNSIGNED, ADD INDEX idxRemoteCode (remoteCode)"
      DoEvents
    End If
    mysql.Execute "UPDATE room SET IPADD = '192.168.137.1' where IPADD is null or IPADD = ''"
    DoEvents
    mysql.Execute "UPDATE room SET remoteCode = 100000 + IDROOM WHERE (remoteCode < 100000) or (remoteCode is null)"
    DoEvents
    mysql.Execute "alter table setting engine=myisam"
    DoEvents
    mysql.Execute "Alter table masters add index `TITLE3_SINGER3` (`TITLE3`, `SINGER3`), add index `SINGER3_TITLE3` (`SINGER3`, `TITLE3`)"
    DoEvents
    mysql.Execute "Alter table masters drop column `DISC`, drop column `WORD`, drop column `POP`, drop column `ROCK`, drop column `JAZZ`, drop column `COUNTRY`, drop column `REGGAE`, drop column `RNB`, drop column `CHACHA`"
    DoEvents
    mysql.Execute "update setting set value = 3 where name = 'database_version'"
    DoEvents
  End If
  
  If database_version < 4 Then
    mysql.Execute "Alter table `karaoke`.`masters` add column `DISC` varchar(255) NULL after `POPULER`, add column `WORD` smallint(11) UNSIGNED NULL after `DISC`, add column `POP` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `WORD`, add column `ROCK` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `POP`, add column `JAZZ` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `ROCK`, add column `COUNTRY` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `JAZZ`, add column `REGGAE` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `COUNTRY`, add column `RNB` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `REGGAE`, add column `CHACHA` tinyint(3) UNSIGNED DEFAULT 0 NOT NULL after `RNB`, change `TITLE6` `TITLE6` varchar(255) DEFAULT '' NULL  after `CHACHA`, change `TITLE7` `TITLE7` varchar(255) DEFAULT '' NULL  after `TITLE6`, change `SINGER6` `SINGER6` varchar(255) DEFAULT '' NULL  after `TITLE7`, change `SINGER7` `SINGER7` varchar(255) DEFAULT '' NULL  after `SINGER6`"
    DoEvents
    mysql.Execute "update setting set value = 4 where name = 'database_version'"
    DoEvents
  End If

  If mysql.Error.Number <> 0 Then
    LogError Name, "updateStructure2"
  End If

End Sub

Private Function getAbbreviation(text As String)

  On Error Resume Next

  Dim result As String
  Dim textArray() As String
  Dim a As Long

  result = ""

  textArray = Split(text, " ")
  For a = LBound(textArray) To UBound(textArray)
    If Len(textArray(a)) > 0 Then
      result = result & Left(textArray(a), 1)
    End If
  Next

  getAbbreviation = result

End Function

