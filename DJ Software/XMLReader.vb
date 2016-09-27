Public Class XMLReader
    Private m_reader As Xml.XmlTextReader
    Private m_path As String

    Public Sub New(ByVal path As String)
        Me.m_path = path
        m_reader = New Xml.XmlTextReader(m_path)
        Me.read()
        Me.read()
    End Sub
    Public Function hasMoreToRead() As Boolean
        Return Not m_reader.EOF
    End Function
    Public Function read() As XMLKnoten

        If m_reader.Read() Then
            Dim result As New XMLKnoten()

            result.ElementName = m_reader.Name
            If m_reader.AttributeCount > 0 Then
                While m_reader.MoveToNextAttribute
                    result.Attributes.Add(m_reader.Name, m_reader.Value)
                End While
            End If
            Return result
        End If

        Return Nothing
    End Function
    Public Function readAll() As List(Of XMLKnoten)
        Dim result As List(Of XMLKnoten) = New List(Of XMLKnoten)
        Do While hasMoreToRead()
            result.Add(read())
        Loop
        Return result
    End Function
    Public Sub close()
        m_reader.Close()
    End Sub
End Class
