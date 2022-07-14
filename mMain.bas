Attribute VB_Name = "mMain"
Option Explicit
Public dbg As String  ' for Debug
Public dbg0 As String ' for Debug
Public dbg1 As String ' for Debug
Public dbg2 As String ' for Debug
Public dbg3 As String ' for Debug
Public dbg4 As String ' for Debug
Public FlashTitleID As Integer

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
Sub logDebug()
    Dim sFileText As String
    Dim iFileNo As Integer
    iFileNo = FreeFile
    'open the file for writing
    '  Open App.Path & "\Debug.txt" For Output As #iFileNo
    'please note, if this file already exists it will be overwritten!
    
    Open App.Path & "\Debug.txt" For Append As #iFileNo
    'please note, if this file already exists it will never be overwritten!
     
    'write some example text to the file
      Print #iFileNo, "----------------------------------------------"
    '  Print #iFileNo, dbg & Space(6) & "video width"
      Print #iFileNo, dbg0 & Space(6) & "blackbox0"
      Print #iFileNo, dbg1 & Space(6) & "blackbox1"
      Print #iFileNo, dbg2 & Space(6) & "blackbox2"
      Print #iFileNo, dbg3 & Space(6) & "blackbox3"
      Print #iFileNo, dbg4 & Space(6) & "blackbox4"
      Print #iFileNo, "----------------------------------------------"
      
     'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
End Sub


