Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OMS.Business.Services

<TestClass>
Public Class SuzukiPreImportTests

    <TestMethod>
    Public Sub PeekInfoTypeCode_ReturnsCorrectCodeFor0600()
        Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CSV_Example", "0600-260717.csv")
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CSV_Example", "0600-260717.csv")
        End If
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine("Z:\Projects\ASTI\OrderManagementSystem-master\CSV_Example\0600-260717.csv")
        End If
        If Not File.Exists(samplePath) Then
            Assert.Inconclusive("Sample CSV 0600-260717.csv not found for unit test.")
            Return
        End If

        Dim code = SuzukiCsvParser.PeekInfoTypeCode(samplePath)
        Assert.AreEqual("0600", code)
    End Sub

    <TestMethod>
    Public Sub PeekInfoTypeCode_ReturnsCorrectCodeFor0602()
        Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CSV_Example", "0602-260701.csv")
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CSV_Example", "0602-260701.csv")
        End If
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine("Z:\Projects\ASTI\OrderManagementSystem-master\CSV_Example\0602-260701.csv")
        End If
        If Not File.Exists(samplePath) Then
            Assert.Inconclusive("Sample CSV 0602-260701.csv not found for unit test.")
            Return
        End If

        Dim code = SuzukiCsvParser.PeekInfoTypeCode(samplePath)
        Assert.AreEqual("0602", code)
    End Sub

    <TestMethod>
    Public Sub Parse0600And0630_ParsesSampleFileCorrectly()
        Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CSV_Example", "0600-260717.csv")
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CSV_Example", "0600-260717.csv")
        End If
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine("Z:\Projects\ASTI\OrderManagementSystem-master\CSV_Example\0600-260717.csv")
        End If
        If Not File.Exists(samplePath) Then
            Assert.Inconclusive("Sample CSV 0600-260717.csv not found for unit test.")
            Return
        End If

        Dim rows = SuzukiCsvParser.Parse0600And0630(samplePath)
        Assert.IsTrue(rows.Count > 0, "Parsed rows count should be greater than 0")
        Assert.AreEqual("0600", rows(0).InfoTypeCode)
        Assert.AreEqual("SUZUKI", rows(0).ClientCode)
        Assert.AreEqual("09407-17403-000", rows(0).CustomerItemNo)
    End Sub

    <TestMethod>
    Public Sub Parse0602_ParsesSampleFileCorrectly()
        Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CSV_Example", "0602-260701.csv")
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CSV_Example", "0602-260701.csv")
        End If
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine("Z:\Projects\ASTI\OrderManagementSystem-master\CSV_Example\0602-260701.csv")
        End If
        If Not File.Exists(samplePath) Then
            Assert.Inconclusive("Sample CSV 0602-260701.csv not found for unit test.")
            Return
        End If

        Dim rows = SuzukiCsvParser.Parse0602(samplePath)
        Assert.IsTrue(rows.Count > 0, "Parsed rows count should be greater than 0")
        Assert.AreEqual("0602", rows(0).InfoTypeCode)
        Assert.AreEqual("SUZUKI", rows(0).ClientCode)
        Assert.AreEqual("36610-41L20-000", rows(0).CustomerItemNo)
    End Sub

End Class
