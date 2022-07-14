VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArtist 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   5220
   ClientTop       =   2175
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmArtist.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   4575
   End
   Begin MSComctlLib.ListView lstArtist 
      Height          =   4455
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Double Click Untuk Detailnya"
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H8000000D&
      Caption         =   "Label1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frmArtist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lblSearch.Tag = 1
Dim SQL As String
Dim I As Integer
Dim ch As ColumnHeader
Dim MyRs As MYSQL_RS
    SQL = " SELECT DISTINCTROW ARTIST FROM master_music"
    Set MyRs = MyConn.Execute(SQL)
     If MyConn.Error.Number <> 0 Then ShowError: Exit Sub
       
            'Memberi judul list data
            
            Set ch = lstArtist.ColumnHeaders.Add(, , "ARTIST", 4000)
            'Set ch = lstCountry.ColumnHeaders.Add(, , "ID", 1)
            
           
            lstArtist.GridLines = True
    Do Until MyRs.EOF
        'Isi list data
        For fld = 1 To MyRs.FieldCount
        Next fld
        Set LV = lstArtist.ListItems.Add(, , (MyRs.Fields((fld - 2)).Value))
        'LV.SubItems(1) = MyRs.Fields((fld - 2)).Value
        MyRs.MoveNext
    Loop
    lstArtist.Enabled = True
End Sub

Private Sub lstArtist_DblClick()
Me.Hide
frmRoom.lstMusic.ColumnHeaders.Clear
frmRoom.lstMusic.ListItems.Clear
Dim sqlm As String
Dim Im As Integer
Dim chm As ColumnHeader
Dim MyRsm As MYSQL_RS
    sqlm = " SELECT TITLE, ARTIST FROM master_music Where ARTIST = '" & lstArtist.SelectedItem & "';"
    Set MyRsm = MyConn.Execute(sqlm)
     If MyConn.Error.Number <> 0 Then ShowError: Exit Sub
       
            'Memberi judul list data
            
            Set chm = frmRoom.lstMusic.ColumnHeaders.Add(, , "Song Title / Judul Lagu ", 6000)
            Set chm = frmRoom.lstMusic.ColumnHeaders.Add(, , "Artist", 4000)
            
        
            frmRoom.lstMusic.GridLines = True
    Do Until MyRsm.EOF
        'Isi list data
        For fld = 1 To MyRsm.FieldCount
        Next fld
        Set LV = frmRoom.lstMusic.ListItems.Add(, , (MyRsm.Fields((fld - 3)).Value))
        LV.SubItems(1) = MyRsm.Fields((fld - 2)).Value
        MyRsm.MoveNext
    Loop
    frmRoom.lstMusic.Enabled = True
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Search_Listview lstArtist, txtSearch.text, Val(lblSearch.Tag)
End Sub



Function Search_Listview(LV As ListView, SearchText As String, Column As Long, Optional SelectRow As Boolean = True, Optional MakeVisible As Boolean = True) As Long
'---------------------------------------------------------------------------------------
' Procedure : Search_Listview
' DateTime  : 3/22/2004 13:25
' Author    : Robert Rowe
' Purpose   : Searches the specified column for the specified text
'             Returns the index of the found column or -1 if no match is found
'---------------------------------------------------------------------------------------

Dim SearchLength As Long
Dim CurrentRow As Long
Dim Result As Long

    SearchLength = Len(SearchText)
    If SearchLength = 0 Then
        Search_Listview = -1
        Exit Function
    End If
    
    Result = -1
    SearchText = UCase$(SearchText)
    If Column = 1 Then
        For CurrentRow = 1 To LV.ListItems.Count
            If UCase$(Left$(LV.ListItems(CurrentRow).text, SearchLength)) = SearchText Then
                Result = LV.ListItems(CurrentRow).Index
                Exit For
            End If
        Next CurrentRow
    Else
        For CurrentRow = 1 To LV.ListItems.Count
            If UCase$(Left$(LV.ListItems(CurrentRow).ListSubItems(Column - 1).text, SearchLength)) = SearchText Then
                Result = LV.ListItems(CurrentRow).Index
                Exit For
            End If
        Next CurrentRow
    End If
    
    If Result > -1 Then
        If SelectRow Then LV.ListItems(Result).Selected = True
        If MakeVisible Then LV.ListItems(Result).EnsureVisible
    End If
    Search_Listview = Result
End Function

