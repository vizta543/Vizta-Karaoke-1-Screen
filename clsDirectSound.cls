VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirectSound"
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

Dim bSound3D As Boolean
Dim BufferDesc As DSBUFFERDESC
Dim BufferDesc3D As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX
Private Const PI As Single = 3.141592
Private odsb As DirectSoundBuffer
Private odsb3d As DirectSound3DBuffer
Private odsl As DirectSound3DListener
Private Sub Class_Initialize()
    'initialize the WAVEFORMATEX structure
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
End Sub
Private Sub Class_Terminate()
    'free up resources
    Set odsb = Nothing
    Set odsb3d = Nothing
End Sub
Public Sub OpenWave(sFilename As String)
    Sound3D = False
    'create a DirectSoundBuffer object to hold a sequence of audio samples
    Set odsb = ods.CreateSoundBufferFromFile(sFilename, BufferDesc, waveFormat)
End Sub
Public Sub Open3DWave(sFilename As String)
    Sound3D = True
    'return a DirectSound3DListener object
    Set odsl = PrimarySound.GetDirectSound3DListener
    'create a DirectSoundBuffer object to hold a sequence of audio samples
    Set odsb = ods.CreateSoundBufferFromFile(sFilename, BufferDesc, waveFormat)
    BufferDesc3D.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    'retrieve the DirectSound3DBuffer object
    Set odsb3d = odsb.GetDirectSound3DBuffer
    'set the inside and outside angles of the sound projection cone for this sound buffer
    odsb3d.SetConeAngles DS3D_MINCONEANGLE, 100, DS3D_IMMEDIATE
    'set the current cone outside volume for this sound buffer
    odsb3d.SetConeOutsideVolume -400, DS3D_IMMEDIATE
    'set the listener's current position (in distance units)
    odsl.SetPosition 0, 0, 0, DS3D_IMMEDIATE
End Sub
Sub Play(Optional rLoop As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT, Optional xp As Single = 0, Optional yp As Single = 0, Optional zp As Single = 0, Optional angle As Long = 0)
    If odsb Is Nothing Then Exit Sub
    If Sound3D Then
        'set the orientation of the sound projection cone for this sound buffer
        odsb3d.SetConeOrientation 5 * Cos(PI * angle / 180), 0, 5 * Sin(PI * angle / 180), DS3D_IMMEDIATE
        'set the sound buffer's current position (in distance units)
        odsb3d.SetPosition 10 * (xp / 100), 10 * (yp / 100), 10 * (zp / 100), DS3D_IMMEDIATE
    End If
    'stop playing
    odsb.Stop
    'move the current play position
    odsb.SetCurrentPosition 0
    'play from the current position
    odsb.Play rLoop
End Sub
Public Sub StopSound()
    If odsb Is Nothing Then Exit Sub
    'stop playing
    odsb.Stop
End Sub
Public Sub Set3DPosition(Optional xp As Single = 0, Optional yp As Single = 0, Optional zp As Single = 0, Optional angle As Long = 0)
    If odsb Is Nothing Or Sound3D = False Then Exit Sub
    odsb3d.SetConeOrientation 5 * Cos(PI * angle / 180), 0, 5 * Sin(PI * angle / 180), DS3D_IMMEDIATE
    odsb3d.SetPosition 10 * (xp / 100), 10 * (yp / 100), 10 * (zp / 100), DS3D_IMMEDIATE
End Sub
Public Property Let Volume(ByVal NewVolume As Long) '-10,000 <-> 0
    If odsb Is Nothing Then Exit Property
    If NewVolume < DSBVOLUME_MIN Then NewVolume = DSBVOLUME_MIN
    If NewVolume > DSBVOLUME_MAX Then NewVolume = DSBVOLUME_MAX
    odsb.SetVolume NewVolume
End Property
Public Property Get Volume() As Long
    If odsb Is Nothing Then Exit Property
    Volume = odsb.GetVolume
End Property
Public Property Let Panning(ByVal NewPanning As Long) '-10,000 <-> 10,000
    If odsb Is Nothing Then Exit Property
    If NewPanning < DSBPAN_LEFT Then NewPanning = DSBPAN_LEFT
    If NewPanning > DSBPAN_RIGHT Then NewPanning = DSBPAN_RIGHT
    odsb.SetPan NewPanning
End Property
Public Property Get Panning() As Long
    If odsb Is Nothing Then Exit Property
    Panning = odsb.GetPan
End Property
Public Property Let Frequency(ByVal NewFrequency As Long) '100 <-> 100,000
    If odsb Is Nothing Then Exit Property
    If NewFrequency < DSBFREQUENCY_MIN Then NewFrequency = DSBFREQUENCY_MIN
    If NewFrequency > DSBFREQUENCY_MAX Then NewFrequency = DSBFREQUENCY_MAX
    odsb.SetFrequency NewFrequency
End Property
Public Property Get Frequency() As Long
    If odsb Is Nothing Then Exit Property
    Frequency = odsb.GetFrequency
End Property
Public Property Let Sound3D(New3D As Boolean)
    bSound3D = New3D
    If bSound3D Then
        waveFormat.nSize = LenB(WAVEFORMATEX)
        BufferDesc.lFlags = (DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME) Or DSBCAPS_STATIC
    Else
        waveFormat.nSize = 0
        BufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    End If
End Property
Public Property Get Sound3D() As Boolean
    Sound3D = bSound3D
End Property
