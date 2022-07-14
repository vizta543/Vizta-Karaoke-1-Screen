Attribute VB_Name = "mMain"
Option Explicit

Private Function getFilePath() As String

  On Error Resume Next

  getFilePath = App.Path & "\log-" & Format$(Now, "yyyy-MM-dd") & ".txt"
  
End Function

Sub clearLog()
  
  On Error Resume Next
  
  Dim fs As New Scripting.FileSystemObject
  Dim fileName As String
  
  fileName = getFilePath()

  fs.deletefile fileName, True

End Sub

Sub LogText(text As String)

  On Error Resume Next
  
  Dim fs As New Scripting.FileSystemObject
  Dim fl As Scripting.TextStream
  Dim fileName As String
  
  fileName = getFilePath()
  
  Set fl = fs.OpenTextFile(fileName, 8, True)
  fl.WriteLine Format$(Now, "hh:mm:ss") & " " & App.title & " " & App.Major & "." & App.Minor & "." & App.Revision & ", " & text
  fl.Close
  Set fl = Nothing
  
  Set fs = Nothing
End Sub

Sub LogError(fileName As String, procedureName As String)
  LogText fileName & "." & procedureName & ", Number: " & Err.Number & ", LastDllError: " & Err.LastDllError & ", Description: " & Err.Description & "," & vbTab & "MySQL: " & MyConn.Error.Description & ", " & MyConnBackup.Error.Description
End Sub

