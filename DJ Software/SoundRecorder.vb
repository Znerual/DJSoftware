
Public Class SoundRecorder
    Private Declare Function mciExecute Lib "winmm.dll" (ByVal IpstrCommand As String) As Long
    Private Declare Function mciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpStrCommand As String, ByVal LpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCall As Integer) As Integer

    Private RS As String
    Private CB As Long
    Private bIsRecording As Boolean = False

    Public Sub StartRecording()
        RS = Space$(128)
        mciSendString("open new type waveaudio alias capture", RS, 128, CB)
        mciSendString("record capture", RS, 128, CB)
        bIsRecording = True
    End Sub
    Public Function isRecording() As Boolean
        Return bIsRecording
    End Function
    Public Function StopRecordingReturnFilename(ByVal PathToSaveDirectory As String) As String
        If Not My.Computer.FileSystem.DirectoryExists(PathToSaveDirectory) Then
            Throw New Exception("Falscher Pfad angegeben\nSoundRecorder.vb\nStopRecording")
        End If
        If bIsRecording = False Then
            Throw New Exception("Kann die Aufnahme nicht stoppen weil sie nie gestartet wurde\nSoundRecorder.vb\nStopRecording")
        End If
        bIsRecording = False
        RS = Space$(128)
        Dim FullPathToFile As String = PathToSaveDirectory + "\" + My.Computer.Clock.LocalTime.ToString("HH-MM-SS-ss") + ".wav"
        mciSendString("stop capture", RS, 128, CB)
        mciSendString("save capture " + FullPathToFile, RS, 128, CB)
        mciSendString("close capture", RS, 128, CB)
        Return FullPathToFile
    End Function
End Class
