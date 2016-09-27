'This is a wrapper for the AxWindowsMediaPlayer with the basic funcinalities

Public Class SoundPlayer
    Private player As AxWMPLib.AxWindowsMediaPlayer
    Public Sub New(ByVal MediaPlayer As AxWMPLib.AxWindowsMediaPlayer)
        Me.player = MediaPlayer
    End Sub
    Public Sub setVolume(ByVal volume As Integer)
        player.settings.volume = volume
    End Sub
    Public Function getVolume() As Integer
        Return player.settings.volume
    End Function
    Public Sub setRate(ByVal rate As Double)
        player.settings.rate = rate
    End Sub
    Public Function getRate() As Double
        Return player.settings.rate
    End Function
    Public Sub setPath(ByVal path As String)
        player.URL = path
    End Sub
    Public Function getPath() As String
        Return player.URL
    End Function
    Public Sub StartPlaying()
        player.Ctlcontrols.play()
    End Sub
    Public Sub StopPlaying()
        player.Ctlcontrols.stop()
    End Sub
End Class
