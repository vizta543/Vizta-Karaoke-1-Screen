VERSION 5.00
Begin VB.Form frmPromoImage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmPromoImage"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr 
      Interval        =   4000
      Left            =   7080
      Top             =   5520
   End
   Begin VB.Image img 
      Height          =   11520
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmPromoImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim promoImagePath As String
    
    promoImagePath = frmRoom.promoImageCollection(frmRoom.promoImageCollectionCurrent)
    
    If frmRoom.FileExists("\\" & vpbServerUtama & promoImagePath) Then
      promoImagePath = "\\" & vpbServerUtama & promoImagePath
    ElseIf frmRoom.FileExists("\\" & vpbServerBackup & promoImagePath) Then
      promoImagePath = "\\" & vpbServerBackup & promoImagePath
    Else
      promoImagePath = ""
    End If
    
    If promoImagePath <> "" Then
      img.Picture = LoadPicture(promoImagePath)
      'added by Andi 25-01-2021
      img.Width = Screen.Width
      img.Height = Screen.Height
      'added by Andi 25-01-2021
    End If
    
    frmRoom.promoImageCollectionCurrent = frmRoom.promoImageCollectionCurrent + 1
    If frmRoom.promoImageCollectionCurrent > frmRoom.promoImageCollection.Count Then
      frmRoom.promoImageCollectionCurrent = 1
    End If
  
    Form2.Visible = False
    
    If Err.Number <> 0 Then
      LogError Name, "Form_Load"
    End If

End Sub

Private Sub tmr_Timer()

  On Error Resume Next
  
  tmr.Enabled = False
  
  Unload Me
  
  Form2.Visible = True
End Sub
