Public Module FileOperations
    Public Function validFile(ByVal fileName As String) As Boolean
        Return fileName.EndsWith(".tld")
    End Function
    Public Function getFileNameWithoutEnd(ByVal fileName As String) As String
        Return fileName.Substring(0, fileName.Length - 4)
    End Function
End Module
