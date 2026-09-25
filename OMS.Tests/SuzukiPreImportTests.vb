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
    Public Sub Parse0602_ParsesFullItemNameAndOrderNo_260716()
        Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CSV_Example", "03_CSV変換後", "0602-260716.csv")
        If Not File.Exists(samplePath) Then
            samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CSV_Example", "03_CSV変換後", "0602-260716.csv")
        End If
        If Not File.Exists(samplePath) Then
            Return
        End If

        Dim rows = SuzukiCsvParser.Parse0602(samplePath)
        Assert.IsTrue(rows.Count > 0, "Parsed rows count should be greater than 0")
        Dim rowWithWire = rows.FirstOrDefault(Function(r) r.CustomerItemName.Contains("WIRE BATTERY"))
        Assert.IsNotNull(rowWithWire, "Row with WIRE BATTERY should exist")
        Assert.AreEqual("WIRE BATTERY PLUS LEAD", rowWithWire.CustomerItemName)
        Assert.AreEqual("SH01TC95017", rowWithWire.CustomerOrderNo)
        Assert.AreEqual("", rowWithWire.CustomerOrderNoProcessNo)
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

    <TestMethod>
    Public Sub ParseSpirits6624_ParsesCustomerItemNoAtColumn7Correctly()
        Dim tempFile = Path.GetTempFileName()
        Try
            Dim sb As New StringBuilder()
            ' ヘッダ行 (39列)
            sb.AppendLine("""情報区分コード"",""書面タイトル区分"",""支払方法等文言区分"",""発注者コード"",""受注者コード"",""受注者事業所コード"",""発行日"",""部品番号"",""部品番号識別－１"",""部品番号識別－２"",""部品名称"",""荷姿コード"",""収容数"",""仕入先コード"",""仕入先工場コード"",""仕入先出荷場所"",""納入先コード"",""納入先工場コード"",""納入場所"",""直納・直送区分"",""発行番号"",""発行番号識別－１"",""発行番号識別－２"",""納入指示区分"",""注文データ区分"",""納入日区分"",""数量区分"",""納入指示日"",""納入指示時刻"",""出荷指示日"",""納入指示数"",""手配担当"",""発注理由"",""初物区分"",""発行注釈"",""供給ライン"",""供給工程"",""工事番号"",""包装区分""")
            ' 6624 データ行: 列7が "36611-96J01-000", 列9が "0"
            sb.AppendLine("""6624"",""H2"",""H2"",""SUZUKI"",""3754"","""",""20260716"",""36611-96J01-000"",""00"",""0"",""WIRE,TRIM SENSOR"",""PB"",""00050"",""3754"",""A00"","""",""P300"","""",""P300"",""H"",""D31108G"","""","""",""5"",""4"",""5"",""1"",""20261012"",""0000"","""",""000010"",""ﾃｷN2  1"",""74"","""",""2A99-000JP"","".A-14"",""ED70A4 98"","""",""""")
            File.WriteAllText(tempFile, sb.ToString(), Encoding.GetEncoding("shift-jis"))

            Dim rows = SuzukiCsvParser.ParseSpirits6624(tempFile)
            Assert.AreEqual(1, rows.Count, "6624データ行が1件パースされる必要があります")
            Assert.AreEqual("6624", rows(0).InfoTypeCode)
            Assert.AreEqual("36611-96J01-000", rows(0).CustomerItemNo, "6624の部品番号は列7 (36611-96J01-000) である必要があります")
            Assert.AreEqual("00", rows(0).CustomerItemNoProcessNo1)
            Assert.AreEqual("0", rows(0).CustomerItemNoProcessNo2)
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub GetCustomerItemNoColumnIndex_ReturnsExpectedIndexForEachInfoCode()
        Assert.AreEqual(10, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0501"))
        Assert.AreEqual(10, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0502"))
        Assert.AreEqual(10, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0600"))
        Assert.AreEqual(10, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0630"))
        Assert.AreEqual(8, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0602"))
        Assert.AreEqual(8, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0650"))
        Assert.AreEqual(8, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0651"))
        Assert.AreEqual(6, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0740"))
        Assert.AreEqual(9, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0813"))
        Assert.AreEqual(9, SuzukiPreImportService.GetCustomerItemNoColumnIndex("0814"))
        Assert.AreEqual(9, SuzukiPreImportService.GetCustomerItemNoColumnIndex("6604"))
        Assert.AreEqual(9, SuzukiPreImportService.GetCustomerItemNoColumnIndex("6634"))
        Assert.AreEqual(7, SuzukiPreImportService.GetCustomerItemNoColumnIndex("6624"), "6624 must be index 7")
        Assert.AreEqual(11, SuzukiPreImportService.GetCustomerItemNoColumnIndex("663N"), "663N must be index 11")
        Assert.AreEqual(11, SuzukiPreImportService.GetCustomerItemNoColumnIndex("663S"), "663S must be index 11")
        Assert.AreEqual(11, SuzukiPreImportService.GetCustomerItemNoColumnIndex("664T"), "664T must be index 11")
    End Sub

    <TestMethod>
    Public Sub Parse0600_10RowsMultiDeliverySchedule_AllParsedWithActiveFlagY()
        Dim tempFile = Path.GetTempFileName()
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("""情報区分コード"",""書面タイトル区分"",""支払方法等文言区分"",""発注者コード"",""受注者コード"",""受注者事業所コード"",""発行日"",""発行時刻"",""対象基準日区分"",""対象基準日"",""部品番号"",""部品番号識別－１"",""部品名称"",""重要保安部品コード"",""荷姿コード"",""収容数"",""納入指示区分"",""仕入先コード"",""仕入先工場コード"",""仕入先出荷場所"",""納入先コード"",""納入先工場コード"",""手配担当"",""購買担当"",""完成工場"",""初物区分"",""先行時間区分"",""先行時間"",""背番号"",""納入サイクル"",""管理区分"",""予備"",""注文データ区分"",""納入日区分"",""生産月度区分"",""納入指示日"",""注文数""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""5"",""H"",""20260824"",""0000094""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""1"",""J"",""202609"",""0000000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""2"",""5"",""G"",""20260713"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""2"",""5"",""G"",""20260714"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""2"",""5"",""G"",""20260715"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""2"",""5"",""G"",""20260716"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""2"",""5"",""G"",""20260717"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""1"",""5"",""G"",""20260720"",""0003000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""1"",""5"",""G"",""20260721"",""0002000""")
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""80"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""310"",""ﾅｶ"",""ｼ"",""B"","""",""D"",""002"","""","""","""","""",""1"",""5"",""G"",""20260722"",""0002000""")
            File.WriteAllText(tempFile, sb.ToString(), New UTF8Encoding(True))

            Dim rows = SuzukiCsvParser.Parse0600And0630(tempFile)
            Assert.AreEqual(10, rows.Count, "0600の10行すべてが正常にパースされる必要があります")
            For i As Integer = 0 To rows.Count - 1
                Assert.AreEqual("0600", rows(i).InfoTypeCode)
                Assert.AreEqual("09407-17403-000", rows(i).CustomerItemNo)
                Assert.AreEqual("Y", rows(i).ActiveFlag, $"行 {i + 1} の ACTIVE_FLAG は初期 'Y' である必要があります")
            Next
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub Parse0600_DeliveryDate_6Digit_8Digit_And_InvalidLength_Handling()
        Dim tempFile = Path.GetTempFileName()
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("""情報区分コード"",""書面タイトル区分"",""支払方法等文言区分"",""発注者コード"",""受注者コード"",""受注者事業所コード"",""発行日"",""発行時刻"",""対象基準日区分"",""対象基準日"",""部品番号"",""部品番号識別－１"",""部品名称"",""重要保安部品コード"",""荷姿コード"",""収容数"",""納入指示区分"",""仕入先コード"",""仕入先工場コード"",""仕入先出荷場所"",""納入先コード"",""納入先工場コード"",""手配担当"",""購買担当"",""完成工場"",""初物区分"",""先行時間区分"",""先行時間"",""背番号"",""納入サイクル"",""管理区分"",""予備"",""注文データ区分"",""納入日区分"",""生産月度区分"",""納入指示日"",""注文数""")
            ' Row 1: 8桁 (20260824) -> 2026/08/24
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""5"",""H"",""20260824"",""0000094""")
            ' Row 2: 6桁 (202609) with no DB -> Nothing (when calRepo is nothing)
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""1"",""J"",""202609"",""0000000""")
            ' Row 3: 7桁 (2026091) -> Nothing
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""1"",""J"",""2026091"",""0000000""")
            ' Row 4: 5桁 (20260) -> Nothing
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""1"",""J"",""20260"",""0000000""")
            ' Row 5: 空文字 -> Nothing
            sb.AppendLine("""0600"",""A1"",""A2"",""SUZUKI"",""3754"","""",""20260702"",""2028"",""4"",""20260713"",""09407-17403-000"",""01"",""CLAMP"","""",""S542"",""01000"",""4NN62"",""3754"",""A00"","""",""3115"",""300"",""ﾅｶ"",""ｼ"",""L"","""",""D"",""002"","""","""","""","""",""1"",""1"",""J"","""",""0000000""")
            File.WriteAllText(tempFile, sb.ToString(), New UTF8Encoding(True))

            Dim rows = SuzukiCsvParser.Parse0600And0630(tempFile)
            Assert.AreEqual(5, rows.Count)
            Assert.IsTrue(rows(0).DeliveryDate.HasValue, "8桁日付はHasValueがTrueである必要があります")
            Assert.AreEqual(New Date(2026, 8, 24), rows(0).DeliveryDate.Value)
            Assert.IsFalse(rows(1).DeliveryDate.HasValue, "calRepo未指定時の6桁日付はNothingである必要があります")
            Assert.IsFalse(rows(2).DeliveryDate.HasValue, "7桁日付はNothingである必要があります")
            Assert.IsFalse(rows(3).DeliveryDate.HasValue, "5桁日付はNothingである必要があります")
            Assert.IsFalse(rows(4).DeliveryDate.HasValue, "空文字日付はNothingである必要があります")
        Finally
            If File.Exists(tempFile) Then File.Delete(tempFile)
        End Try
    End Sub

    <TestMethod>
    Public Sub GetPreviousMonthSymbol_ReturnsCorrectSymbolForEachMonth()
        ' 1月 -> 前年12月 (N)
        Assert.AreEqual("N", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 1, 15)))
        ' 2月 -> 1月 (A)
        Assert.AreEqual("A", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 2, 10)))
        ' 3月 -> 2月 (B)
        Assert.AreEqual("B", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 3, 1)))
        ' 4月 -> 3月 (C)
        Assert.AreEqual("C", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 4, 30)))
        ' 5月 -> 4月 (D)
        Assert.AreEqual("D", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 5, 20)))
        ' 6月 -> 5月 (E)
        Assert.AreEqual("E", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 6, 15)))
        ' 7月 -> 6月 (F)
        Assert.AreEqual("F", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 7, 7)))
        ' 8月 -> 7月 (G)
        Assert.AreEqual("G", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 8, 25)))
        ' 9月 -> 8月 (H)
        Assert.AreEqual("H", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 9, 25)))
        ' 10月 -> 9月 (J) ※Iはスキップ
        Assert.AreEqual("J", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 10, 1)))
        ' 11月 -> 10月 (K)
        Assert.AreEqual("K", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 11, 15)))
        ' 12月 -> 11月 (M) ※Lはスキップ
        Assert.AreEqual("M", SuzukiPreImportService.GetPreviousMonthSymbol(New DateTime(2026, 12, 31)))
    End Sub

    <TestMethod>
    Public Sub SupportedInfoCodes_ContainsValidCodesAndExcludes0813()
        Dim expectedCodes = {"0501", "0502", "0600", "0602", "0630", "0650", "0651", "0740", "0814", "6604", "6624", "6634", "663N", "663S", "664T"}
        For Each c In expectedCodes
            Assert.IsTrue(SuzukiPreImportService.SupportedInfoCodes.Contains(c), $"情報区分コード {c} は SupportedInfoCodes に含まれている必要があります")
        Next

        ' 2026/09/18 仕様変更により除外された0813
        Assert.IsFalse(SuzukiPreImportService.SupportedInfoCodes.Contains("0813"), "情報区分コード 0813 は SupportedInfoCodes に含まれていてはなりません")

        ' 未対応コード
        Assert.IsFalse(SuzukiPreImportService.SupportedInfoCodes.Contains("9999"), "未定義のコードは含まれていてはなりません")
        Assert.IsFalse(SuzukiPreImportService.SupportedInfoCodes.Contains(""), "空文字は含まれていてはなりません")
    End Sub

End Class
