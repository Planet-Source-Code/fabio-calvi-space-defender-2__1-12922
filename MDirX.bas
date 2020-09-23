Attribute VB_Name = "MDirX"
'// sound use
Public DS As IDirectSound
Public Const MAX_SOUNDS = 4
Public DSB(4) As IDirectSoundBuffer
Public gbUseSound As Boolean
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)

Public Type WAVEFORMATEX  '// mimics PCMWAVEFORMAT structure - 16 bytes
        wFormatTag As Integer     '???
        nChannels As Integer      '1=mono, 2=stereo
        nSamplesPerSec As Long    '11025,22050,44100
        nAvgBytesPerSec As Long   'SamplesPerSec x 1 (8 bit) or x 2 (16 bit)
        nBlockAlign As Integer    '???
        wBitsPerSample As Integer '8 or 16
        cbSize As Integer
End Type
'
' Plays a sound
'
Sub Play_Sound(sWaveFile As String, SLoop As Boolean, SPlayStop As Boolean, SChannel As Integer)

    Dim lngFlag As Long
    Dim sFilename As String
    
    '// check global sound use variable //
    If Not gbUseSound Then Exit Sub
    
    If SLoop Then lngFlag = DSBPLAY_LOOPING
    
    sFilename = App.Path & "\" & sWaveFile
    If SPlayStop Then
        If Not (DSB(SChannel) Is Nothing) Then
            DSB(SChannel).Stop
               Set DSB(SChannel) = Nothing
        End If
        LoadWAVIntoDSB DS, sFilename, DSB(SChannel)
        DSB(SChannel).Play 0, 0, lngFlag
    Else
        If Not (DSB(SChannel) Is Nothing) Then
            DSB(SChannel).Stop
            Set DSB(SChannel) = Nothing
        End If
    End If
    Exit Sub

ErrorPlaySound:
    MsgBox Err.Description, , "Play_Sound ERROR"
End Sub

'
' Loads a Wave file into a direct sound buffer
'
Sub LoadWAVIntoDSB(Lds As IDirectSound, ByVal fName As String, Ldsb As IDirectSoundBuffer)
    
    On Error GoTo ErrorLoadWAV
    Dim hWave As Long
    Dim pcmwave As WAVEFORMATEX
    Dim lngSize As Long
    Dim lngPosition As Long
    Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
    Dim aByte() As Byte
    
    ReDim aByte(1 To FileLen(fName))
    hWave = FreeFile
    Open fName For Binary As hWave
    Get hWave, , aByte
    Close hWave
    lngPosition = 1
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
        lngPosition = lngPosition + 1
    Wend
    CopyMemory VarPtr(pcmwave), VarPtr(aByte(lngPosition + 8)), Len(pcmwave)
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
        lngPosition = lngPosition + 1
    Wend
    CopyMemory VarPtr(lngSize), VarPtr(aByte(lngPosition + 4)), Len(lngSize)
    Dim dsbd As DSBUFFERDESC
    With dsbd
        .dwSize = Len(dsbd)
        .dwflags = DSBCAPS_CTRLDEFAULT
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    Lds.CreateSoundBuffer dsbd, Ldsb, Nothing
    Ldsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
    CopyMemory ptr1, VarPtr(aByte(lngPosition + 4 + 4)), lng1
    If lng2 <> 0 Then
        CopyMemory ptr2, VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
    End If
    Exit Sub
    
ErrorLoadWAV:
    MsgBox Err.Description, , "LoadWAVIntoDSB ERROR"
End Sub
Public Sub InitDirectX()
    
    gbUseSound = True
    
    If gbUseSound Then
        'setup sound engine
        DirectSoundCreate ByVal 0&, DS, Nothing
        Call DS.SetCooperativeLevel(Form1.hwnd, DSSCL_NORMAL)
    End If
    
End Sub
Public Sub Terminate()

    Dim i As Integer
    MDirX.Play_Sound "null.wav", False, True, 0
    MDirX.Play_Sound "null.wav", False, True, 1
    MDirX.Play_Sound "null.wav", False, True, 2
    MDirX.Play_Sound "null.wav", False, True, 3
    MDirX.Play_Sound "null.wav", False, True, 4
    For i = 0 To MAX_SOUNDS
        Set DSB(i) = Nothing
    Next i
    Set DS = Nothing
    
End Sub


