Public Class XMLWriter
    Private enc As New System.Text.UTF8Encoding
    Private m_writer As Xml.XmlTextWriter
    Private m_path As String

    Public Sub New(ByVal path As String)
        Me.m_path = path
        m_writer = New Xml.XmlTextWriter(m_path, enc)
        m_writer.Formatting = Xml.Formatting.Indented
        m_writer.Indentation = 4
        m_writer.WriteStartDocument()
        m_writer.WriteStartElement("start")

    End Sub
    Public Sub startBlock(ByVal topic As String)
        m_writer.WriteStartElement(topic)
    End Sub
    Public Sub endBlock()
        m_writer.WriteEndElement()
    End Sub
    Public Sub addAttribute(ByVal attribut As String, ByVal context As String)
        m_writer.WriteAttributeString(attribut, context)
    End Sub
    Public Sub close()
        m_writer.WriteEndElement()
        m_writer.WriteEndDocument()
        m_writer.Close()
    End Sub
End Class
