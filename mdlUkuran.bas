Attribute VB_Name = "mdlUkuran"
Public MiniVideoWMPwidth As Integer
Public MiniVideoWMPheight As Integer
Public MiniVideoWMPtop As Integer
Public MiniVideoWMPleft As Integer

Public MiniTVcamerawidth As Integer
Public MiniTVcameraheight As Integer
Public MiniTVcameratop As Integer
Public MiniTVcameraleft As Integer

Public UkuranVideo As Integer  '1=max 2=min 3=main key/tempo/vol

Sub AktifUkuran()
    On Error Resume Next
    UkuranVideo = 1
    MiniVideoWMPwidth = 5480
    MiniVideoWMPheight = 4010
    MiniVideoWMPtop = 470
    MiniVideoWMPleft = 4970
    
    MiniTVcamerawidth = 5500
    MiniTVcameraheight = 4010
    MiniTVcameratop = 470
    MiniTVcameraleft = 4970
End Sub
