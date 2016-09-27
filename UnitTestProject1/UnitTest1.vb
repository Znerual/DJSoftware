Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestFileOperationsValidFile1()
        Assert.AreEqual(DJ_Software.validFile("aldh.tld"), True)
    End Sub
    <TestMethod()> Public Sub TestFileOperationsValidFile2()
        Assert.AreEqual(Of String)(DJ_Software.validFile("aldh.ted"), False, "Konnte nicht aldh zurückliefern")
        Assert.AreEqual(DJ_Software.getFileNameWithoutEnd("adlp.tld"), "adlp")
    End Sub
    <TestMethod()> Public Sub TestFileOperationsShorten()
        Assert.AreEqual(DJ_Software.getFileNameWithoutEnd("adlp.tld"), "adlp")
    End Sub
End Class