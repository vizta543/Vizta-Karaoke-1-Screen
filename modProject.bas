Attribute VB_Name = "modProject"
Option Explicit

Public Const brandName As String = "Vizta"

Public Const outletServerPort As Long = 32768

Public Const frmRoomRemoteCode As Boolean = True

Public Declare Function mbedtlsBlowfishDecrypt Lib "Vizta Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String, ByVal index0 As Long, ByVal index1 As Long) As Long
Public Declare Function opensslAESDecrypt Lib "Vizta Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String, ByVal index0 As Long, ByVal index1 As Long) As Long
Public Declare Function opensslSHA512Digest Lib "Vizta Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String) As Long

