Attribute VB_Name = "mSubClass"
Option Explicit

'PARALEL PORT
Private Declare Function Inp Lib "inpout32.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32.dll" _
Alias "Out32" (ByVal PortAddress As Integer, ByVal value As Integer)

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
Public oldProc As Long

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    On Error Resume Next
    
    Dim lokasi As String
    Dim videoinputindex As Integer
    Dim i As Integer

    On Error Resume Next
    
    If (vpbfrmLoading) Then
        Exit Function
    End If
      
    If (frmRoom.vVideo = 0) Or (frmRoom.vVideo = 2) Then 'COMPUTER
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 1
                    If UkuranVideo = 1 Then
                        frmRoom.ScreenSaverAktif = 99
                        frmRoom.Minimal
                    Else
                        If (vpbfrmCamera) Then
                            videoinputindex = frmCamera.VideoCap1.VideoInputs.FindVideoInput("Video Tuner")
                            If frmCamera.VideoCap1.VideoInput = videoinputindex Then
                                Unload frmCamera
                            Else
                                frmRoom.ScreenSaverAktif = 0
                                frmRoom.Maksimal
                            End If
                        Else
                            frmRoom.ScreenSaverAktif = 0
                            frmRoom.Maksimal
                        End If
                    End If
                Case 2
                    frmRoom.ScreenSaverAktif = 99
                    frmRoom.PlayerMidi
                    UkuranVideo = 2
                Case 3
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.moviestart
                    Form2.Height = 0
                Case 4
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerTV
                Case 5
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.chatbuka
                Case 6
                    frmRoom.prcDVD
                Case 7
                    frmRoom.prcVCD
                Case 8
                    frmRoom.prcCD
                Case 9
                    frmRoom.prcHP
                Case 10
                    If frmRoom.cmdPause(1).Visible = True Then
                        frmRoom.cmdPause_Click 1
                    Else
                        If UkuranVideo <> 1 Then
                            frmRoom.ScreenSaverAktif = 5
                            frmRoom.Maksimal
                        Else
                            frmRoom.ScreenSaverAktif = 5
                            frmRoom.Maksimal
                            frmRoom.cmdPlay_Click
                            frmRoom.txtSearch.SelStart = 0
                            frmRoom.txtSearch.SelLength = Len(frmRoom.txtSearch.text)
                        End If
                    End If
                Case 11
                    If Not (vpbfrmVocal) Then
                        If vpbRemoteStatus = 2 Then
                            ScoreSetup = False
                        End If
                        If ScoreSetup = True Then
                            ScoreSetup = False
                            frmVocal.VocalAktif = 3
                            frmVocal.Show
                        Else
                            ScoreSetup = True
                            frmVocal.VocalAktif = 2
                            frmVocal.Show
                        End If
                        frmRoom.tmrMainVocal.Enabled = True
                    End If
                Case 13
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    prcSonglist
                Case 14
                    frmRoom.ScreenSaverAktif = 0
                    If UkuranVideo <> 1 Then
                        frmRoom.Maksimal
                        DoEvents
                    End If
                    frmRoom.prcHits
                Case 15
                    frmRoom.ScreenSaverAktif = 0
                    If UkuranVideo <> 1 Then
                        frmRoom.Maksimal
                        DoEvents
                    End If
                    vpbHits = False
                    frmRoom.prcNew (50)
                Case 16
                    frmRoom.ScreenSaverAktif = 0
                    If UkuranVideo <> 1 Then
                        frmRoom.Maksimal
                        DoEvents
                    End If
                    vpbHits = False
                    frmRoom.prcPopuler (50)
                Case 17
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    DoEvents
                    frmRoom.pencet_btnplaylist
                Case 18
                    frmRoom.prcLoginMember
                Case 19
                    frmRoom.prcAddMemberPlaylist
                Case 20
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Minimal
                    DoEvents
                    frmTransparent.GantiSkin (11)
                    frmCountry.Show
                    frmCountry.lstCountry.SetFocus
                Case 21
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Minimal
                    DoEvents
                    frmTransparent.GantiSkin (11)
                    frmCountry.Show
                    frmCountry.lstCountry.SetFocus
                Case 22
                    frmConfirmasi.vpengirim = 15
                    frmRoom.Minimal
                    frmConfirmasi.Show
                    frmConfirmasi.Move 4020, 2325
                    frmDice2.Show
                    frmDice2.Move 3765, 1875
                Case 23
                    frmRoom.prcSaran
                Case 26
                    If UkuranVideo = 1 Then
                        frmRoom.PlayLagu
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 30
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    prcTitleList
                Case 31
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    prcSingerList
                Case 33
                    frmabout.Show
                Case 34
                    frmHelp.Show
                Case 35
                    frmCall.vpengirim = 3
                    frmCall.Show
                    frmCall.Timer1.Enabled = True
                Case 38
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 39
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 40
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalKey
                    frmRoom.prcKeyUp
                Case 41
                     If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalKey
                    frmRoom.prckeyDown
                Case 42
                     If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalTempo
                    frmRoom.prcTempoUp
                Case 43
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalTempo
                    frmRoom.prcTempoDown
                Case 44
                    If vpbAmpliAuto = 1 Then
                        RemoteAmpli "ECHOUP", 0
                    ElseIf vpbAmpliAuto = 3 Then
                        Select Case vEffectAmpli
                            Case 2
                                vEffectAmpli = 1
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                            Case 3
                                vEffectAmpli = 2
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                            Case 4
                                vEffectAmpli = 3
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                        End Select
                        frmCall.vpengirim = 4
                        frmCall.Show
                        frmCall.Timer1.Interval = 5000
                        frmCall.Timer1.Enabled = True
                    ElseIf vpbAmpliAuto = 2 Then
                        frmConfirmasi.vpengirim = 12
                        frmRoom.Minimal
                        DoEvents
                        frmConfirmasi.Show
                        frmConfirmasi.Text1.SetFocus
                        frmConfirmasi.Top = 1000
                    End If
                Case 45
                    If vpbAmpliAuto = 1 Then
                        RemoteAmpli "ECHODN", 0
                    ElseIf vpbAmpliAuto = 3 Then
                        Select Case vEffectAmpli
                            Case 1
                                vEffectAmpli = 2
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                            Case 2
                                vEffectAmpli = 3
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                            Case 3
                                vEffectAmpli = 4
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                        End Select
                        
                        frmCall.vpengirim = 4
                        frmCall.Show
                        frmCall.Timer1.Interval = 5000
                        frmCall.Timer1.Enabled = True
                    ElseIf vpbAmpliAuto = 2 Then
                        frmConfirmasi.vpengirim = 12
                        frmRoom.Minimal
                        frmConfirmasi.Show
                        frmConfirmasi.Text1.SetFocus
                    End If
                Case 46
                    If vpbAmpliAuto = 1 Then
                      RemoteAmpli "MICUP", 0
                    ElseIf vpbAmpliAuto = 2 Then
                      frmConfirmasi.vpengirim = 14
                      frmRoom.Minimal
                      frmConfirmasi.Show
                      frmConfirmasi.Text1.SetFocus
                    End If
                Case 47
                    If vpbAmpliAuto = 1 Then
                      RemoteAmpli "MICDN", 0
                    ElseIf vpbAmpliAuto = 2 Then
                      frmConfirmasi.vpengirim = 14
                      frmRoom.Minimal
                      frmConfirmasi.Show
                      frmConfirmasi.Text1.SetFocus
                    End If
                Case 48
                    If vpbRemoteStatus = 1 Then
                        If frmRoom.txtSearch.text = "" Then
                            frmRoom.ScreenSaverAktif = 5
                        Else
                            frmRoom.ScreenSaverAktif = 5
                        End If
                        frmRoom.MinimalVolume
                        frmRoom.setvolumenaik
                    Else
                    
                        If frmRoom.txtSearch.text = "" Then
                            frmRoom.ScreenSaverAktif = 5
                        Else
                            frmRoom.ScreenSaverAktif = 5
                        End If
                        
                        If vpbAmpliAuto = 3 Then
                            vEffectAmpli = 1
                            For i = 1 To 5
                                RemoteAmpli "PRESET0", 1
                                Sleep 3
                            Next i
                            
                            frmCall.vpengirim = 4
                            frmCall.Show
                            frmCall.Timer1.Interval = 7000
                            frmCall.Timer1.Enabled = True
                        ElseIf vpbAmpliAuto = 2 Then
                            For i = 1 To 5
                                RemoteAmpli "Preset2", 1
                                Sleep 3
                            Next i
                            vEffectAmpli = 2
                        End If
                    End If
                Case 49
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 50
                    frmRoom.cmdSlow_Click
                Case 51
                    frmRoom.cmdFast_Click
                Case 52
                    frmRoom.prcRepeat
                Case 53
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    frmRoom.prcPageUpLstAll
                Case 54
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    frmRoom.prcPageDownLstAll
                Case 55
                    If UkuranVideo = 1 Then
                        frmRoom.PlayLagu
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 64
                    If UkuranVideo = 1 Then
                        frmRoom.PlayLagu
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 65
                    If Not (vpbfrmVocal) Then
                        If vpbBolehMainVocal = True Then
                            frmRoom.setvocal
                            If frmRoom.vVocalterus = False Then
                                frmVocal.VocalAktif = 0
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            Else
                                frmVocal.VocalAktif = 1
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            End If
                            frmRoom.vVocalterus = Not (frmRoom.vVocalterus)
                        End If
                    End If
                Case 66
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    If frmRoom.lstPlaylist.Visible = True And frmRoom.lstPlaylist.ListItems.Count <> 0 Then
                        frmRoom.cmdUp_Click
                    Else
                        frmRoom.PrioritasLagu
                    End If
                Case 67
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    If frmRoom.lstPlaylist.Visible = True Then
                        frmRoom.cmdDelete_Click
                    End If
                Case 68
                    If frmRoom.cmdPause(0).Visible = True Then
                        frmRoom.cmdPause_Click 0
                    Else
                        frmRoom.cmdPause_Click 1
                    End If
                Case 69
                    frmRoom.cmdStop_Click
                    frmRoom.ScreenSaverAktif = 5
                Case 70
                    vpbBlackBox = 2
                    frmUser.turnDiscoLampOn
                Case 71
                    vpbBlackBox = 0
                    frmUser.turnDiscoLampOff
                Case 73
                    If Not (vpbfrmVocal) Then
                        ScoreSetup = True
                        If ScoreSetup = True Then
                            ScoreSetup = False
                            frmVocal.VocalAktif = 3
                            frmVocal.Show
                        Else
                            ScoreSetup = True
                            frmVocal.VocalAktif = 2
                            frmVocal.Show
                        End If
                        frmRoom.tmrMainVocal.Enabled = True
                    End If
                Case 74
                    frmRoom.prcDeleteMemberPlaylist
                Case 75
                    frmRoom.prcClearMemberPlaylist
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf (frmRoom.vVideo = 1) Then 'MIDI
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 1
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerKomputer
                Case 3
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.moviestart
                    Form2.Height = 0
                Case 4
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerTV
                Case 5
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.chatbuka
                Case 23
                    frmRoom.prcSaran
                Case 33
                    frmabout.Show
                Case 34
                    frmHelp.Show
                Case 35
                    frmCall.vpengirim = 3
                    frmCall.Show
                    frmCall.Timer1.Enabled = True
'                Case 38
'                    If frmRoom.txtSearch.text = "" Then
'                        frmRoom.ScreenSaverAktif = 1
'                    Else
'                        frmRoom.ScreenSaverAktif = 6
'                    End If
'                    frmRoom.MinimalVolume
'                    frmRoom.setvolumenaik
'                Case 39
'                    If frmRoom.txtSearch.text = "" Then
'                        frmRoom.ScreenSaverAktif = 1
'                    Else
'                        frmRoom.ScreenSaverAktif = 6
'                    End If
'                    frmRoom.MinimalVolume
'                    frmRoom.setvolumeturun
                Case 44
                    If vpbAmpliAuto = 1 Then
                        RemoteAmpli "ECHOUP", 0
                    ElseIf vpbAmpliAuto = 3 Then
                        Select Case vEffectAmpli
                            Case 2
                                vEffectAmpli = 1
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                            Case 3
                                vEffectAmpli = 2
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                            Case 4
                                vEffectAmpli = 3
                                For i = 1 To 3
                                    RemoteAmpli "PRESETDN", 1
                                    Sleep 3
                                Next i
                        End Select
                        frmCall.vpengirim = 4
                        frmCall.Show
                        frmCall.Timer1.Interval = 5000
                        frmCall.Timer1.Enabled = True
                    ElseIf vpbAmpliAuto = 2 Then
                        frmConfirmasi.vpengirim = 12
                        frmConfirmasi.Show
                        frmConfirmasi.Text1.SetFocus
                    End If
                Case 45
                    If vpbAmpliAuto = 1 Then
                        RemoteAmpli "ECHODN", 0
                    ElseIf vpbAmpliAuto = 3 Then
                        Select Case vEffectAmpli
                            Case 1
                                vEffectAmpli = 2
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                            Case 2
                                vEffectAmpli = 3
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                            Case 3
                                vEffectAmpli = 4
                                For i = 1 To 3
                                    RemoteAmpli "PRESETUP", 1
                                    Sleep 3
                                Next i
                        End Select
                        
                        frmCall.vpengirim = 4
                        frmCall.Show
                        frmCall.Timer1.Interval = 5000
                        frmCall.Timer1.Enabled = True
                    ElseIf vpbAmpliAuto = 2 Then
                        frmConfirmasi.vpengirim = 12
                        frmConfirmasi.Show
                        frmConfirmasi.Text1.SetFocus
                    End If
                Case 46
                    RemoteAmpli "MICUP", 0
                Case 47
                    RemoteAmpli "MICDN", 0
'                Case 48
'                    If frmRoom.txtSearch.text = "" Then
'                        frmRoom.ScreenSaverAktif = 1
'                    Else
'                        frmRoom.ScreenSaverAktif = 6
'                    End If
'                    frmRoom.MinimalVolume
'                    frmRoom.setvolumenaik
'                Case 49
'                    If frmRoom.txtSearch.text = "" Then
'                        frmRoom.ScreenSaverAktif = 1
'                    Else
'                        frmRoom.ScreenSaverAktif = 6
'                    End If
'                    frmRoom.MinimalVolume
'                    frmRoom.setvolumeturun
'                Case 55
'                    SendKeys "{ENTER}"
'                Case 64
'                    SendKeys "{ENTER}"
                Case 70
                    vpbBlackBox = 2
                    frmUser.turnDiscoLampOn
                Case 71
                    vpbBlackBox = 0
                    frmUser.turnDiscoLampOff
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf (frmRoom.vVideo = 7) Then 'TV
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 1
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerKomputer
                Case 2
                    frmRoom.ScreenSaverAktif = 99
                    frmRoom.PlayerMidi
                    UkuranVideo = 2
                Case 3
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.moviestart
                    Form2.Height = 0
                Case 4
                    If UkuranVideo = 1 Then
                        frmRoom.ScreenSaverAktif = 99
                        frmRoom.Minimal
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 5
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.chatbuka
                Case 10
                    If frmRoom.cmdPause(1).Visible = True Then
                        frmRoom.cmdPause_Click 1
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                        frmRoom.cmdPlay_Click
                    End If
                Case 20
                    If UkuranVideo = 1 Then
                        frmRoom.lstTV_Click
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 23
                    frmRoom.prcSaran
                Case 26
                    If UkuranVideo = 1 Then
                        frmRoom.lstTV_Click
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 28
                    If UkuranVideo = 1 Then
                        frmRoom.lstTV_Click
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 33
                    frmabout.Show
                Case 34
                    frmHelp.Show
                Case 35
                    frmCall.vpengirim = 3
                    frmCall.Show
                    frmCall.Timer1.Enabled = True
                Case 38
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 39
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 48
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 49
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 50
                    frmRoom.cmdSlow_Click
                Case 51
                    frmRoom.cmdFast_Click
                Case 52
                    frmRoom.prcRepeat
                Case 55
                    SendKeys "{ENTER}"
                Case 64
                    SendKeys "{ENTER}"
                Case 65
                    If Not (vpbfrmVocal) Then
                        If vpbBolehMainVocal = True Then
                            frmRoom.setvocal
                            If frmRoom.vVocalterus = False Then
                                frmVocal.VocalAktif = 0
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            Else
                                frmVocal.VocalAktif = 1
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            End If
                            frmRoom.vVocalterus = Not (frmRoom.vVocalterus)
                        End If
                    End If
                Case 68
                    If frmRoom.cmdPause(0).Visible = True Then
                        frmRoom.cmdPause_Click 0
                    Else
                        frmRoom.cmdPause_Click 1
                    End If
                Case 69
                    frmRoom.cmdStop_Click
                Case 70
                    vpbBlackBox = 2
                    frmUser.turnDiscoLampOn
                Case 71
                    vpbBlackBox = 0
                    frmUser.turnDiscoLampOff
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 5 Then 'MOVIE
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 1
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerKomputer
                Case 2
                    frmRoom.ScreenSaverAktif = 99
                    frmRoom.PlayerMidi
                    UkuranVideo = 2
                Case 3
                    If UkuranVideo = 1 Then
                        frmRoom.ScreenSaverAktif = 99
                        frmRoom.Minimal
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 4
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerTV
                Case 5
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.chatbuka
                Case 6
                    frmRoom.prcDVD
                Case 7
                    frmRoom.prcVCD
                Case 8
                    frmRoom.prcCD
                Case 9
                    frmRoom.prcHP
                Case 10
                    If frmRoom.cmdPause(1).Visible = True Then
                        frmRoom.cmdPause_Click 1
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                        If frmRoom.lstMovie.ListItems.Count > 0 Then
                            frmRoom.PlayMovie
                        End If
                    End If
                Case 20
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    frmCategory.Show
                Case 21
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Minimal
                    DoEvents
                    frmTransparent.GantiSkin (11)
                    frmCountry.Show
                    frmCountry.lstCountry.SetFocus
                Case 23
                    frmRoom.prcSaran
                Case 26
                    If UkuranVideo = 1 Then
                        frmRoom.PlayMovie
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 30
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                    frmRoom.vpointer = 1
                    frmRoom.flsTitle.Movie = App.Path + "\picture\anim\titlemovie"
                    frmRoom.txtSearch_Change
                    frmRoom.txtSearch.SelStart = 0
                    frmRoom.txtSearch.SelLength = Len(frmRoom.txtSearch.text)
                Case 31
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    frmRoom.txtSearch.SetFocus
                    sMakeCaret frmRoom.txtSearch, frmRoom.caretLebar, frmRoom.caretTinggi
                    frmRoom.vpointer = 2
                    frmRoom.flsTitle.Movie = App.Path + "\picture\anim\artistmovie"
                    frmRoom.txtSearch_Change
                    frmRoom.txtSearch.SelStart = 0
                    frmRoom.txtSearch.SelLength = Len(frmRoom.txtSearch.text)
                Case 33
                    frmabout.Show
                Case 34
                    frmHelp.Show
                Case 35
                    frmCall.vpengirim = 3
                    frmCall.Show
                    frmCall.Timer1.Enabled = True
                Case 38
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 39
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 5
                    Else
                        frmRoom.ScreenSaverAktif = 5
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 48
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 49
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 50
                    frmRoom.cmdSlow_Click
                Case 51
                    frmRoom.cmdFast_Click
                Case 52
                    frmRoom.prcRepeat
                Case 55
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    If frmRoom.lstMovie.ListItems.Count > 0 Then
                        frmRoom.PlayMovie
                    End If
                Case 64
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.Maksimal
                    If frmRoom.lstMovie.ListItems.Count > 0 Then
                        frmRoom.PlayMovie
                    End If
                Case 68
                    If frmRoom.cmdPause(0).Visible = True Then
                        frmRoom.cmdPause_Click 0
                    Else
                        frmRoom.cmdPause_Click 1
                    End If
                Case 69
                    frmRoom.cmdStop_Click
                Case 70
                    vpbBlackBox = 2
                    frmUser.turnDiscoLampOn
                Case 71
                    vpbBlackBox = 0
                    frmUser.turnDiscoLampOff
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 3 Then 'CHAT
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 1
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerKomputer
                Case 2
                    frmRoom.ScreenSaverAktif = 99
                    frmRoom.PlayerMidi
                    UkuranVideo = 2
                Case 3
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.moviestart
                    Form2.Height = 0
                Case 4
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.PlayerTV
                Case 5
                    frmRoom.ScreenSaverAktif = 0
                    frmRoom.chatbuka
                Case 10
                    SendKeys "{ENTER}"
                Case 23
                    frmRoom.prcSaran
                Case 26
                    If UkuranVideo = 1 Then
                        frmRoom.chatSend
                    Else
                        frmRoom.ScreenSaverAktif = 0
                        frmRoom.Maksimal
                    End If
                Case 33
                    frmabout.Show
                Case 34
                    frmHelp.Show
                Case 35
                    frmCall.vpengirim = 3
                    frmCall.Show
                    frmCall.Timer1.Enabled = True
                Case 38
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 39
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 48
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumenaik
                Case 49
                    If frmRoom.txtSearch.text = "" Then
                        frmRoom.ScreenSaverAktif = 1
                    Else
                        frmRoom.ScreenSaverAktif = 6
                    End If
                    frmRoom.MinimalVolume
                    frmRoom.setvolumeturun
                Case 50
                    frmRoom.cmdSlow_Click
                Case 51
                    frmRoom.cmdFast_Click
                Case 52
                    frmRoom.prcRepeat
                Case 55
                    SendKeys "{ENTER}"
                Case 64
                    SendKeys "{ENTER}"
                Case 65
                    If Not (vpbfrmVocal) Then
                        If vpbBolehMainVocal = True Then
                            frmRoom.setvocal
                            If frmRoom.vVocalterus = False Then
                                frmVocal.VocalAktif = 0
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            Else
                                frmVocal.VocalAktif = 1
                                frmVocal.Show
                                frmRoom.tmrMainVocal.Enabled = True
                            End If
                            frmRoom.vVocalterus = Not (frmRoom.vVocalterus)
                        End If
                    End If
                Case 68
                    If frmRoom.cmdPause(0).Visible = True Then
                        frmRoom.cmdPause_Click 0
                    Else
                        frmRoom.cmdPause_Click 1
                    End If
                Case 69
                    frmRoom.cmdStop_Click
                Case 70
                    vpbBlackBox = 2
                    frmUser.turnDiscoLampOn
                Case 71
                    vpbBlackBox = 0
                    frmUser.turnDiscoLampOff
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 8 Then 'WELCOME
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmWelcome.StartKaraoke
                Case 26
                    frmWelcome.StartKaraoke
                Case 55
                    frmWelcome.StartKaraoke
                Case 64
                    frmWelcome.StartKaraoke
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 9 Then 'CALL & BILL
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmCall.prcOK
                Case 26
                    frmCall.prcOK
                Case 55
                    frmCall.prcOK
                Case 64
                    frmCall.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 10 Then 'SARAN
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmSaran.prcOK
                Case 26
                    frmSaran.prcOK
                Case 55
                    frmSaran.prcOK
                Case 64
                    frmSaran.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 11 Then 'ABOUT
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmabout.prcOK
                Case 26
                    frmabout.prcOK
                Case 55
                    frmabout.prcOK
                Case 64
                    frmabout.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 12 Then 'HELP
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmHelp.prcOK
                Case 26
                    frmHelp.prcOK
                Case 55
                    frmHelp.prcOK
                Case 64
                    frmHelp.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 13 Then 'Country
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmCountry.prcOK
                Case 26
                    frmCountry.prcOK
                Case 55
                    frmCountry.prcOK
                Case 64
                    frmCountry.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 14 Then 'Category
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmCategory.prcOK
                Case 26
                    frmCategory.prcOK
                Case 55
                    frmCategory.prcOK
                Case 64
                    frmCategory.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 15 Then 'Confirmasi
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmConfirmasi.prcKonfirm
                Case 26
                    frmConfirmasi.prcKonfirm
                Case 44
                    If frmConfirmasi.vpengirim = 12 Then
                        Select Case vEffectAmpli
                            Case 2
                                vEffectAmpli = 1
                            Case 3
                                vEffectAmpli = 2
                            Case 4
                                vEffectAmpli = 3
                        End Select
                        frmConfirmasi.flsAnim.SetVariable "vtulisan", vEffectAmpli
                    End If
                Case 45
                    If frmConfirmasi.vpengirim = 12 Then
                        Select Case vEffectAmpli
                            Case 1
                                vEffectAmpli = 2
                            Case 2
                                vEffectAmpli = 3
                            Case 3
                                vEffectAmpli = 4
                        End Select
                        frmConfirmasi.flsAnim.SetVariable "vtulisan", vEffectAmpli
                    End If
                    
                Case 46
                    frmConfirmasi.tmrAktif.Enabled = False
                    frmConfirmasi.tmrAktif.Enabled = True
                    If frmConfirmasi.vpengirim = 14 Then
                      Select Case vEffectAmpli
                        Case 1
                          vEffectAmpli = 9
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
                        Case 2
                          vEffectAmpli = 10
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
                        Case 3
                          vEffectAmpli = 11
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
                        Case 4
                          vEffectAmpli = 12
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 3
                        Case 5
                          vEffectAmpli = 1
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 6
                          vEffectAmpli = 2
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 7
                          vEffectAmpli = 3
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 8
                          vEffectAmpli = 4
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                      End Select
                    End If
                Case 47
                    frmConfirmasi.tmrAktif.Enabled = False
                    frmConfirmasi.tmrAktif.Enabled = True
                    If frmConfirmasi.vpengirim = 14 Then
                      Select Case vEffectAmpli
                        Case 1
                          vEffectAmpli = 5
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
                        Case 2
                          vEffectAmpli = 6
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
                        Case 3
                          vEffectAmpli = 7
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
                        Case 4
                          vEffectAmpli = 8
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 1
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                          vEffectAmpli = 1
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 10
                          vEffectAmpli = 2
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 11
                          vEffectAmpli = 3
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                        Case 12
                          vEffectAmpli = 4
                          frmConfirmasi.flsAnim.SetVariable "vtulisan", 2
                      End Select
                    End If
                    
                Case 55
                    frmConfirmasi.prcKonfirm
                Case 64
                    frmConfirmasi.prcKonfirm
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf frmRoom.vVideo = 16 Then 'tambahjam
        WndProc = 0
        If uMsg = WM_HOTKEY Then
            Select Case wParam
                Case 10
                    frmTambahJam.prcOK
                Case 26
                    frmTambahJam.prcOK
                Case 55
                    frmTambahJam.prcOK
                Case 64
                    frmTambahJam.prcOK
            End Select
        Else
            WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
        End If
    End If
    
End Function
