Imports System.IO
Imports System.Text
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

    <TestMethod>
    Public Sub PeekInfoTypeCode_WithMangled3DigitCode_ReturnsEmptyAndDiagnostic()
        Dim tempFile = Path.GetTempFileName()
        Try
            ' Excel等で "0600" が 600 に変換されたパターン
            File.WriteAllText(tempFile, "600,A1,A2,SUZUKI,3754", Encoding.GetEncoding("shift-jis"))
            Dim diagMsg As String = ""
            Dim code = SuzukiCsvParser.PeekInfoTypeCode(tempFile, diagMsg)

            Assert.AreEqual(String.Empty, code, $"3桁コードは無効として空文字を返す必要があります. Actual code: [{code}], diag: [{diagMsg}]")
            Assert.IsTrue(diagMsg.Contains("600"), $"診断メッセージに値 '600' が含まれる必要があります. Actual diag: [{diagMsg}]")
            Assert.IsTrue(diagMsg.Contains("3桁") OrElse diagMsg.Contains("欠落"), $"診断メッセージに桁数や欠落に関する説明が含まれる必要があります. Actual diag: [{diagMsg}]")
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub PeekInfoTypeCode_WithHeaderLine_ReturnsCodeFromSecondLine()
        Dim tempFile = Path.GetTempFileName()
        Try
            ' 1行目が日本語ヘッダ、2行目がデータ行のパターン
            Dim sb As New Text.StringBuilder()
            sb.AppendLine("""情報区分コード"",""書面タイトル区分"",""支払方法等文言区分""")
            sb.AppendLine("""0600"",""A1"",""A2""")
            File.WriteAllText(tempFile, sb.ToString(), Encoding.GetEncoding("shift-jis"))

            Dim diagMsg As String = ""
            Dim code = SuzukiCsvParser.PeekInfoTypeCode(tempFile, diagMsg)

            Assert.AreEqual("0600", code, $"ヘッダ行をスキップして2行目の0600を取得する必要があります. Actual code: [{code}], diag: [{diagMsg}]")
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub PeekInfoTypeCode_WithUtf8Bom_ReturnsCode()
        Dim tempFile = Path.GetTempFileName()
        Try
            ' UTF-8 with BOM のパターン
            Dim content = """0600"",""A1"",""A2"""
            File.WriteAllText(tempFile, content, New UTF8Encoding(True))

            Dim diagMsg As String = ""
            Dim code = SuzukiCsvParser.PeekInfoTypeCode(tempFile, diagMsg)

            Assert.AreEqual("0600", code, $"UTF-8 BOM付でも正常に0600を取得する必要があります. Actual code: [{code}], diag: [{diagMsg}]")
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub Parse0600And0630_WithHeaderAndUtf8Bom_ParsesCorrectly()
        Dim tempFile = Path.GetTempFileName()
        Try
            Dim sb As New StringBuilder()
            ' ヘッダ行
            sb.AppendLine("""情報区分コード"",""書面タイトル区分"",""支払方法等文言区分"",""発注者コード"",""受注者コード"",""受注者事業所コード"",""発行日"",""発行時刻"",""対象基準日区分"",""対象基準日"",""部品番号"",""部品番号識別－１"",""部品名称"",""重要保安部品コード"",""荷姿コード"",""収容数"",""納入指示区分"",""仕入先コード"",""仕入先工場コード"",""仕入先出荷場所"",""納入先コード"",""納入先工場コード"",""手配担当"",""購買担当"",""完成工場"",""初物区分"",""先行時間区分"",""先行時間"",""背番号"",""納入サイクル"",""管理区分"",""予備"",""注文データ区分"",""納入日区分"",""生産月度区分"",""納入指示日"",""注文数""")
            ' 正常データ行 (37列)
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""5"",""0"",""20260713"",""100""")
            File.WriteAllText(tempFile, sb.ToString(), New UTF8Encoding(True))

            Dim rows = SuzukiCsvParser.Parse0600And0630(tempFile)
            Assert.AreEqual(1, rows.Count, "ヘッダをスキップしてデータ行1件が正しくパースされる必要があります")
            Assert.AreEqual("0600", rows(0).InfoTypeCode)
            Assert.AreEqual("09407-17403-000", rows(0).CustomerItemNo)
            Assert.AreEqual(100L, rows(0).OrderQty)
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

End Class
