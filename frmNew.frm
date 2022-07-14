VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNew 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmNew"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstTop 
      Height          =   3375
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   7280168
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmNew.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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



Private Sub Form_Load()
    On Error Resume Next
    vpbfrmNew = True
    Dim lokasi As String
    lokasi = App.Path
    

        Skin1.LoadSkin lokasi + "\skin\sknnew.skn"
        Skin1.ApplySkinByName hWnd, "sknnew"
'        lstTop.Picture = LoadPicture(lokasi + "\Picture\normalscreen\nations.jpg")

    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE + SWP_NOSIZE
    Me.Move (0 - Screen.Width)
    frmRoom.Enabled = False
    frmRoom.hotnon
    isiListTop
End Sub

Sub isiListTop()
    On Error Resume Next
    Dim chm As ColumnHeader
    Dim LV As ListItem
    
    lstTop.ColumnHeaders.Clear
    lstTop.ListItems.Clear
    
    Set chm = lstTop.ColumnHeaders.add(, , , 10)
    Set chm = lstTop.ColumnHeaders.add(, , , 1500)
    lstTop.ColumnHeaders(2).Alignment = lvwColumnCenter
    
    Set LV = lstTop.ListItems.add(, , "10")
        LV.SubItems(1) = "10"
    Set LV = lstTop.ListItems.add(, , "25")
        LV.SubItems(1) = "25"
    Set LV = lstTop.ListItems.add(, , "50")
        LV.SubItems(1) = "50"
    Set LV = lstTop.ListItems.add(, , "100")
        LV.SubItems(1) = "100"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    vpbfrmNew = False
    frmRoom.hot
End Sub

Sub prcOK()
    On Error Resume Next
    vpbHits = False
    frmRoom.Enabled = True
    frmRoom.prcNew (lstTop.selectedItem.text)
    Unload Me
End Sub

Private Sub lstTop_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        prcOK
    End If
End Sub
