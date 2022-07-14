VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirectMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DirectX_Sound - visual basic example
' written by Pieter Philippaerts (Pieter@allapi.net)
' http://www.allapi.net/
'
' You may use this code within your programs, however
' you are strictly forbidden from to redistribute this
' code as source in any manner whatsoever.

Private odml   As DirectMusicLoader
Private odmp As DirectMusicPerformance
Private odms As DirectMusicSegment
Private odmss As DirectMusicSegmentState
Public Sub OpenFile(sFile As String)
    If IsPlaying Then 'if the music file of this class is currently playing, stop it
        'stop playing
        odmp.Stop odms, odmss, 0, 0
        'free resources
        Set odms = Nothing
        Set odmss = Nothing
    End If
    'load a segment from a file
    Set odms = odml.LoadSegment(sFile)
    'DirectMusic that the segment is based on a standard MIDI file, not one
    'authored specifically for DirectMusic. Calling this method ensures that
    'certain events are handled properly when the segment is played.
    odms.SetStandardMidiFile
    'download the collection associated with the segment so that the port can play the instruments
    odms.Download odmp
End Sub
Public Sub Play()
    'begin playback of a segment
    Set odmss = odmp.PlaySegment(odms, 0, 0)
End Sub
Public Sub StopMusic()
    'stop playback of one or more segments
    odmp.Stop odms, odmss, 0, 0
End Sub
Private Sub Class_Initialize()
    'create necesarry DirectMusic objects
    Set odml = odx.DirectMusicLoaderCreate()
    Set odmp = odx.DirectMusicPerformanceCreate()
    'initialize the performance and associate it with a DirectSound object
    'this method should be called only once, before any other methods are called on the performance
    odmp.Init Nothing, frmMain.hWnd
    'set the active port for the performance. This method must be called after the performance
    'is initialized and before any instruments are downloaded or any segment is played
    odmp.SetPort -1, 4
    'turn automatic downloading of instruments on
    odmp.SetMasterAutoDownload True
    'set the directory to be searched by the DirectMusicLoader.LoadBand,
    'DirectMusicLoader.LoadChordmap, DirectMusicLoader.LoadCollection,
    'DirectMusicLoader.LoadSegment, and DirectMusicLoader.LoadStyle methods
    'when a fully qualified path is not supplied
    odml.SetSearchDirectory App.Path
End Sub
Private Sub Class_Terminate()
    If IsPlaying Then odmp.Stop odms, odmss, 0, 0 'If the class is playing, stop it
    'close down the performance
    odmp.CloseDown
    Set odml = Nothing
    Set odmp = Nothing
    Set odms = Nothing
    Set odmss = Nothing
End Sub
Public Function IsPlaying() As Boolean
    If odmp Is Nothing Or odms Is Nothing Or odmss Is Nothing Then Exit Function
    IsPlaying = odmp.IsPlaying(odms, odmss)
End Function
'undefined
Public Property Get Volume() As Single
    Groove = odmp.GetMasterVolume()
End Property
Public Property Let Volume(New_Volume As Single)
    odmp.SetMasterVolume New_Volume
End Property
'0.25 <-> 2
Public Property Get Tempo() As Single
    Groove = odmp.GetMasterTempo()
End Property
Public Property Let Tempo(New_Tempo As Single)
    odmp.SetMasterTempo New_Tempo
End Property

