Attribute VB_Name = "mdlMain"
'DirectX_Sound - visual basic example
' written by Pieter Philippaerts (Pieter@allapi.net)
' http://www.allapi.net/
'
' You may use this code within your programs, however
' you are strictly forbidden from to redistribute this
'code as source in any manner whatsoever

'General declarations
Public odx As DirectX7
Public ods As DirectSound
Public PrimarySound As DirectSoundBuffer
Public WaveFile As clsDirectSound
Public MidiFile As clsDirectMusic
Public Wave3DFile As clsDirectSound
Sub Main()
    'Create a new DirectX7 object
    Set odx = New DirectX7
    'init the sound
    InitSound
    'Create two new  Direct Sound classes and one new DirectMusic class
    Set WaveFile = New clsDirectSound
    Set MidiFile = New clsDirectMusic
    Set Wave3DFile = New clsDirectSound
    'show the main form
    frmMain.Show
End Sub
Sub InitSound()
    Dim primDesc As DSBUFFERDESC, tFormat As WAVEFORMATEX
    Set ods = odx.DirectSoundCreate("")
    ods.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY
    primDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
    Set PrimarySound = ods.CreateSoundBuffer(primDesc, tFormat)
End Sub
Sub UnloadProgram()
    'clean up
    Set WaveFile = Nothing
    Set MidiFile = Nothing
    Set Wave3DFile = Nothing
    Set PrimarySound = Nothing
    Set ods = Nothing
    Set odx = Nothing
End Sub
