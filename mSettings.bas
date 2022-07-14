Attribute VB_Name = "mSettings"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function getIniSetting(application As String, Key As String, default As String) As String

  On Error Resume Next
  
  Dim strBuffer As String
  Dim lLength As Long
  Dim BufferSize As Long
  
  BufferSize = 2048
  
  strBuffer = Space(BufferSize)

  lLength = GetPrivateProfileString(application, Key, default, strBuffer, BufferSize, App.Path & "\..\..\Setting-vod.ini")
  
  getIniSetting = Left(strBuffer, lLength)
End Function
