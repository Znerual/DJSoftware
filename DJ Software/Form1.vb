

Public Class Form1
    Dim Pfade As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Dim autostart As String = Application.ExecutablePath
    Dim autostart1 As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "/" & IO.Path.GetFileName(autostart)
    Dim soundRecorder As New SoundRecorder()
    Dim soundPlayer1 As SoundPlayer
    Dim soundPlayer2 As SoundPlayer
    Public Function returnAutostartPath() As String
        Return autostart1
    End Function
    Private Sub TrackBar5_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar5.Scroll
        TrackBar1.Value = 10 - TrackBar5.Value
        TrackBar4.Value = TrackBar5.Value
        soundPlayer1.setVolume(TrackBar1.Value * 10)
        soundPlayer2.setVolume(TrackBar4.Value * 10)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog2.ShowDialog()
       
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        soundPlayer1.setPath(Pfade(ListBox1.Text))
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        soundPlayer2.setPath(Pfade(ListBox2.Text))
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        soundPlayer1.StopPlaying()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        soundPlayer2.StopPlaying()
    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        soundPlayer1.setVolume(TrackBar1.Value * 10)
    End Sub
    Private Sub TrackBar4_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll
        soundPlayer2.setVolume(TrackBar4.Value * 10)
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        setRate(soundPlayer1, TrackBar2.Value)
    End Sub

    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        setRate(soundPlayer2, TrackBar3.Value)
    End Sub
    Private Sub setRate(ByRef MediaPlayer As SoundPlayer, ByVal Value As Integer)
        Select Case Value
            Case 10
                MediaPlayer.setRate(2.5)
            Case 9
                MediaPlayer.setRate(2)
            Case 8
                MediaPlayer.setRate(1.7)
            Case 7
                MediaPlayer.setRate(1.5)
            Case 6
                MediaPlayer.setRate(1.2)
            Case 5
                MediaPlayer.setRate(1)
            Case 4
                MediaPlayer.setRate(0.8)
            Case 3
                MediaPlayer.setRate(0.5)
            Case 2
                MediaPlayer.setRate(0.3)
            Case 1
                MediaPlayer.setRate(0.2)
            Case 0
                MediaPlayer.setRate(0.1)
        End Select

    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        soundPlayer1.setPath(Pfade(ListBox1.Text))
        soundPlayer1.StartPlaying()
        soundPlayer2.setPath(Pfade(ListBox2.Text))
        soundPlayer2.StartPlaying()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        soundPlayer1.StopPlaying()
        soundPlayer2.StopPlaying()
    End Sub


    Private Sub TrackBar6_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar6.Scroll
        TrackBar2.Value = TrackBar6.Value
        TrackBar3.Value = TrackBar6.Value
        setRate(soundPlayer1, TrackBar6.Value)
        setRate(soundPlayer2, TrackBar6.Value)


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        soundPlayer1 = New SoundPlayer(AxWindowsMediaPlayer1)
        soundPlayer2 = New SoundPlayer(AxWindowsMediaPlayer2)
        updateTemplateSelection()

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Button9.Enabled = False
        Button10.Enabled = True
        Label19.Text = "Recording..."
        Label19.Visible = True

        soundRecorder.StartRecording()

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        FolderBrowserDialog1.ShowDialog()

        Button9.Enabled = True
        Button10.Enabled = False
        Button11.Enabled = True

        Label19.Text = "Recording Stopped"
        Label19.Visible = False

        Dim savedFilePath As String = soundRecorder.StopRecordingReturnFilename(FolderBrowserDialog1.SelectedPath)
        MsgBox("File Was Successfully Recorded : " + savedFilePath)
        My.Computer.Audio.Stop()

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Label19.Text = "Abspielen"
        Label19.Visible = True
        My.Computer.Audio.Play("C:\Users\Laurenz\Recording.wav", AudioPlayMode.Background)
    End Sub

    Private Sub LinksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinksToolStripMenuItem.Click
        soundPlayer1.setPath("")
    End Sub

    Private Sub RechtsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RechtsToolStripMenuItem.Click
        soundPlayer2.setPath("")
    End Sub

    Private Sub BeideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BeideToolStripMenuItem.Click
        soundPlayer1.setPath("")
        soundPlayer2.setPath("")
    End Sub

    Private Sub LinksToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinksToolStripMenuItem1.Click
        soundPlayer1.setPath(Pfade(ListBox1.Text))
    End Sub

    Private Sub RechtsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RechtsToolStripMenuItem1.Click
        soundPlayer2.setPath(Pfade(ListBox2.Text))
    End Sub

    Private Sub BeideToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BeideToolStripMenuItem1.Click
        soundPlayer1.setPath(Pfade(ListBox1.Text))
        soundPlayer2.setPath(Pfade(ListBox2.Text))
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        TrackBar1.Value = 0
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        TrackBar1.Value = 5
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        TrackBar1.Value = 10
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        TrackBar2.Value = 0
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        TrackBar2.Value = 5
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        TrackBar2.Value = 10
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        TrackBar4.Value = 0
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        TrackBar4.Value = 5
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        TrackBar4.Value = 10
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        TrackBar3.Value = 0
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        TrackBar3.Value = 5
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        TrackBar3.Value = 10
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        TrackBar5.Value = 0
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        TrackBar5.Value = 5
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        TrackBar5.Value = 10
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        TrackBar6.Value = 0
    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        TrackBar6.Value = 5
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        TrackBar6.Value = 10
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        ListBox1.SelectedItem = ListBox1.Text

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        ListBox2.SelectedItem = ListBox2.Text

    End Sub

    Private Sub TrackBar7_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar7.Scroll
        TrackBar1.Value = TrackBar7.Value
        TrackBar4.Value = TrackBar7.Value
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            IO.File.Copy(autostart, autostart1)
            MsgBox("Autostart erfolgreich erstellt")
        Catch ex As Exception
            MsgBox("Leider ist die Erstellung des autoruns fehlgeschlagen")
        End Try
    End Sub


    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        For Each track As String In OpenFileDialog1.FileNames
            Dim trackZerlegt As String() = track.Split("\")
            Dim TrackName As String = trackZerlegt(trackZerlegt.Length - 1)
            Dim TrackPath As String = track
            If Not Pfade.ContainsKey(TrackName) Then
                Pfade.Add(TrackName, TrackPath)
            End If
            ListBox1.SelectedIndex = ListBox1.Items.Add(TrackName)

        Next
    End Sub

    Private Sub OpenFileDialog2_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        For Each track As String In OpenFileDialog2.FileNames
            Dim trackZerlegt As String() = track.Split("\")
            Dim TrackName As String = trackZerlegt(trackZerlegt.Length - 1)
            Dim TrackPath As String = track
            If Not Pfade.ContainsKey(TrackName) Then
                Pfade.Add(TrackName, TrackPath)
            End If
            ListBox2.SelectedIndex = ListBox2.Items.Add(TrackName)
        Next
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If TextBox1.Text <> "" Then
            If My.Settings.pfad <> "" Then
                FolderBrowserDialog2.SelectedPath = My.Settings.pfad
            End If

            FolderBrowserDialog2.ShowDialog()
            If FolderBrowserDialog2.SelectedPath <> "" Then
                My.Settings.pfad = FolderBrowserDialog2.SelectedPath
                My.Settings.Save()
                Dim filePath As String = FolderBrowserDialog2.SelectedPath + "\" + TextBox1.Text + ".tld"
                Dim xml_writer As New XMLWriter(filePath)
                xml_writer.startBlock("list1")
                listboxToXml(ListBox1, xml_writer)
                xml_writer.endBlock()
                xml_writer.startBlock("list2")
                listboxToXml(ListBox2, xml_writer)
                xml_writer.endBlock()
                xml_writer.close()
                updateTemplateSelection()
            End If
        Else
            MsgBox("Bitte geben sie einen Namen ein")
        End If


    End Sub
    Public Sub listboxToXml(ByRef listbox As ListBox, ByRef xml_writer As XMLWriter)
        For i As Integer = 0 To listbox.Items.Count - 1
            xml_writer.addAttribute(i.ToString(), listbox.Items(i).ToString())
        Next
    End Sub
    Public Sub updateTemplateSelection()
        ComboBox1.Items.Clear()
        If My.Settings.pfad <> "" Then
            For Each datei In My.Computer.FileSystem.GetFiles(My.Settings.pfad)
                Dim dateiName As String = datei.Split("\").Last
                If validFile(dateiName) Then
                    ComboBox1.Items.Add(getFileNameWithoutEnd(dateiName))
                End If
            Next
        End If
    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        If ComboBox1.SelectedItem.ToString <> "" Then
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            Dim path As String = My.Settings.pfad + "\" + ComboBox1.SelectedItem.ToString() + ".tld"
            Dim xml_reader As New XMLReader(path)
            For Each item As XMLKnoten In xml_reader.readAll()
                MsgBox(item.ElementName)
            Next
            fillListboxWithXMLData(ListBox1, xml_reader)
            fillListboxWithXMLData(ListBox2, xml_reader)
            xml_reader.close()
        End If
    End Sub
    Public Sub fillListboxWithXMLData(ByRef listbox As ListBox, ByRef xml_reader As XMLReader)
        If xml_reader.hasMoreToRead() Then
            Dim knoten As XMLKnoten = xml_reader.read()

            For Each titel As String In knoten.Attributes.Values
                listbox.Items.Add(titel)
            Next
        End If



    End Sub
End Class
