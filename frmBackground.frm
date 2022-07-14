VERSION 5.00
Begin VB.Form frmBackground 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmBackground"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    Me.Left = 0
    Me.Top = 0
'    Me.Width = 15360
'    Me.Height = 11520

    'edited by Andi 28-01-2021
    If frmUser.settingScreenResolution = "S-SD" Then
        Me.Width = 15360
        Me.Height = 11520
    ElseIf frmUser.settingScreenResolution = "S-HD" Then
        Me.Width = 19200
        Me.Height = 10800
    ElseIf frmUser.settingScreenResolution = "S-FULLHD" Then
        Me.Width = 28800
        Me.Height = 16200
    End If
    'edited by Andi 28-01-2021
End Sub
