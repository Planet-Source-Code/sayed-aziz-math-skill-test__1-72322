Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = &H1
Global Const SND_MEMORY = &H4
Global Const SND_LOOP = &H8
Public WavClaps As String, WavMov As String, WavClaps1 As String
Public Function NoiseGet(ByVal FileName) As String

Dim buffer As String
Dim f As Integer
Dim SoundBuffer As String

    On Error GoTo NoiseGet_Error

    buffer = Space$(1024)
    SoundBuffer = ""
    f = FreeFile
    Open FileName For Binary As f
    Do While Not EOF(f)
        Get #f, , buffer
        SoundBuffer = SoundBuffer & buffer
    Loop
    Close f
    NoiseGet = Trim$(SoundBuffer)
    
Exit Function

NoiseGet_Error:
    
    SoundBuffer = ""
    Exit Function

End Function
Public Sub NoisePlay(SoundBuffer As String, ByVal PlayMode As Integer)

    Dim retcode As Integer
    
    If SoundBuffer = "" Then Exit Sub

    sndPlaySound vbNullString, sndAsync
    retcode = sndPlaySound(ByVal SoundBuffer, PlayMode Or SND_LOOP)
    
End Sub


