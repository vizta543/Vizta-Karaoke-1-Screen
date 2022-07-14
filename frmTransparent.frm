VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTransparent 
   BorderStyle     =   0  'None
   Caption         =   "frmTransparent"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmTransparent.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmTransparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lokasi As String

Public Sub LoadSkin(PilihSkin As Integer)
    On Error Resume Next
    lokasi = App.Path
    Select Case PilihSkin
    Case 0
        If frmUser.settingScreenResolution = "S-SD" Then
            Skin1.LoadSkin lokasi + "\skin\main.skn"
            Skin1.ApplySkinByName hWnd, "MainForm"
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            Skin1.LoadSkin lokasi + "\skin\main_hd.skn"
            Skin1.ApplySkinByName hWnd, "MainForm"
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            Skin1.LoadSkin lokasi + "\skin\main_fullhd.skn"
            Skin1.ApplySkinByName hWnd, "MainForm"
        End If
    Case 5
        If frmUser.settingScreenResolution = "S-SD" Then
            Skin1.LoadSkin lokasi + "\skin\movie.skn"
            Skin1.ApplySkinByName hWnd, "MovieForm"
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            Skin1.LoadSkin lokasi + "\skin\movie_hd.skn"
            Skin1.ApplySkinByName hWnd, "MovieForm"
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            Skin1.LoadSkin lokasi + "\skin\movie_fullhd.skn"
            Skin1.ApplySkinByName hWnd, "MovieForm"
        End If
    Case 7
        If frmUser.settingScreenResolution = "S-SD" Then
            Skin1.LoadSkin lokasi + "\skin\skntv.skn"
            Skin1.ApplySkinByName hWnd, "skntv"
        ElseIf frmUser.settingScreenResolution = "S-HD" Then
            Skin1.LoadSkin lokasi + "\skin\skntv_hd.skn"
            Skin1.ApplySkinByName hWnd, "skntv"
        ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
            Skin1.LoadSkin lokasi + "\skin\skntv_fullhd.skn"
            Skin1.ApplySkinByName hWnd, "skntv"
        End If
    End Select
End Sub

Public Sub GantiSkin(PilihSkin As Integer)
    On Error Resume Next
    If frmUser.settingScreenResolution = "S-SD" Then
        Select Case PilihSkin
            Case 0
                Skin1.ApplySkinByName hWnd, "MainForm"
            Case 1
'                Me.Top = 1500
                Skin1.ApplySkinByName hWnd, "MinimizePolos"
            Case 2 'HIT
                Skin1.ApplySkinByName hWnd, "MainHit"
            Case 3 'NEW
                Skin1.ApplySkinByName hWnd, "MainNew"
            Case 4 'POPULAR
                Skin1.ApplySkinByName hWnd, "MainPopuler"
            Case 5 'PLAYLIST
                Skin1.ApplySkinByName hWnd, "MainPlaylist"
            Case 6
                Skin1.ApplySkinByName hWnd, "volume"
            Case 7
                Skin1.ApplySkinByName hWnd, "MainFormKey"
            Case 8
                Skin1.ApplySkinByName hWnd, "MainFormTempo"
            Case 9
                Skin1.ApplySkinByName hWnd, "MinimizeKey"
            Case 10
                Skin1.ApplySkinByName hWnd, "MinimizeTempo"
            Case 11
                Skin1.ApplySkinByName hWnd, "MinimizeNation"
            Case 12
                Skin1.ApplySkinByName hWnd, "MinimizeNationMovie"
        End Select
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        Select Case PilihSkin
            Case 0
                Skin1.ApplySkinByName hWnd, "MainForm"
            Case 1
                Skin1.ApplySkinByName hWnd, "MinimizePolos"
            Case 2 'HIT
                Skin1.ApplySkinByName hWnd, "MainHit"
            Case 3 'NEW
                Skin1.ApplySkinByName hWnd, "MainNew"
            Case 4 'POPULAR
                Skin1.ApplySkinByName hWnd, "MainPopuler"
            Case 5 'PLAYLIST
                Skin1.ApplySkinByName hWnd, "MainPlaylist"
            Case 6
                Skin1.ApplySkinByName hWnd, "volume"
            Case 7
                Skin1.ApplySkinByName hWnd, "MainFormKey"
            Case 8
                Skin1.ApplySkinByName hWnd, "MainFormTempo"
            Case 9
                Skin1.ApplySkinByName hWnd, "MinimizeKey"
            Case 10
                Skin1.ApplySkinByName hWnd, "MinimizeTempo"
            Case 11
                Skin1.ApplySkinByName hWnd, "MinimizeNation"
            Case 12
                Skin1.ApplySkinByName hWnd, "MinimizeNationMovie"
        End Select
    ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
        Select Case PilihSkin
            Case 0
                Skin1.ApplySkinByName hWnd, "MainForm"
            Case 1
'                Me.Top = 2000
                Skin1.ApplySkinByName hWnd, "MinimizePolos"
            Case 2 'HIT
                Skin1.ApplySkinByName hWnd, "MainHit"
            Case 3 'NEW
                Skin1.ApplySkinByName hWnd, "MainNew"
            Case 4 'POPULAR
                Skin1.ApplySkinByName hWnd, "MainPopuler"
            Case 5 'PLAYLIST
                Skin1.ApplySkinByName hWnd, "MainPlaylist"
            Case 6
                Skin1.ApplySkinByName hWnd, "volume"
            Case 7
                Skin1.ApplySkinByName hWnd, "MainFormKey"
            Case 8
                Skin1.ApplySkinByName hWnd, "MainFormTempo"
            Case 9
                Skin1.ApplySkinByName hWnd, "MinimizeKey"
            Case 10
                Skin1.ApplySkinByName hWnd, "MinimizeTempo"
            Case 11
                Skin1.ApplySkinByName hWnd, "MinimizeNation"
            Case 12
                Skin1.ApplySkinByName hWnd, "MinimizeNationMovie"
        End Select
    End If
End Sub

Public Sub GantiSkinMovie(PilihSkin As Integer)
    On Error Resume Next
        Select Case PilihSkin
                Case 0
                    Skin1.ApplySkinByName hWnd, "MovieForm"
                Case 1
                    Skin1.ApplySkinByName hWnd, "MinimizePolos"
                Case 2 'HIT
                    Skin1.ApplySkinByName hWnd, "MainHit"
                Case 3 'NEW
                    Skin1.ApplySkinByName hWnd, "MainNew"
                Case 4 'POPULAR
                    Skin1.ApplySkinByName hWnd, "MainPopuler"
                Case 5 'PLAYLIST
                    Skin1.ApplySkinByName hWnd, "MainPlaylist"
                Case 6
                    Skin1.ApplySkinByName hWnd, "volume"
                Case 7
                    Skin1.ApplySkinByName hWnd, "MainFormKey"
                Case 8
                    Skin1.ApplySkinByName hWnd, "MainFormTempo"
                Case 9
                    Skin1.ApplySkinByName hWnd, "MinimizeKey"
                Case 10
                    Skin1.ApplySkinByName hWnd, "MinimizeTempo"
                Case 11
                    Skin1.ApplySkinByName hWnd, "MinimizeNation"
        End Select
End Sub

Public Sub GantiSkinTV(PilihSkin As Integer)
    On Error Resume Next
    Select Case PilihSkin
        Case 0
            Skin1.ApplySkinByName hWnd, "skntv"
        Case 1
            Skin1.ApplySkinByName hWnd, "MinimizePolos"
    End Select
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    
    lokasi = App.Path
    If frmUser.settingScreenResolution = "S-SD" Then
        Skin1.LoadSkin lokasi + "\skin\main.skn"
        Skin1.ApplySkinByName hWnd, "MainForm"
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        Skin1.LoadSkin lokasi + "\skin\main_hd.skn"
        Skin1.ApplySkinByName hWnd, "MainForm"
    ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
        Skin1.LoadSkin lokasi + "\skin\main_fullhd.skn"
        Skin1.ApplySkinByName hWnd, "MainForm"
    End If
End Sub



