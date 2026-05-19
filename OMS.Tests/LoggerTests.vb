Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports OMS.Common

<TestClass>
Public Class LoggerTests

    Private _tmpRoot As String

    <TestInitialize>
    Public Sub Init()
        ' テストごとに一時フォルダを作って衝突を避ける
        _tmpRoot = Path.Combine(Path.GetTempPath(), "OMS_LoggerTests_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_tmpRoot)
    End Sub

    <TestCleanup>
    Public Sub Cleanup()
        Try
            If Directory.Exists(_tmpRoot) Then
                Directory.Delete(_tmpRoot, True)
            End If
        Catch
            ' 他プロセスで掴まれて削除できない場合などは無視
        End Try
        Console.SetOut(New StreamWriter(Console.OpenStandardOutput()) With {.AutoFlush = True})
    End Sub

    '==============================
    ' コンストラクタ
    '==============================
    <TestMethod>
    Public Sub Ctor_EmptyPath_Throws()
        Dim ex = Assert.ThrowsException(Of ArgumentException)(
        Sub()
            Dim _ignore = New OMS.Common.Logger("   ")
        End Sub)
        StringAssert.Contains(ex.Message, "log file path is empty.")
    End Sub

    <TestMethod>
    Public Sub Ctor_ValidPath_Ok()
        Dim logPath = Path.Combine(_tmpRoot, "app.log")
        Dim logger = New Logger(logPath)
        Assert.IsNotNull(logger)
    End Sub

    '==============================
    ' Write(message) : 既定 INFO
    '==============================
    <TestMethod>
    Public Sub Write_Info_WritesOneLine()
        Dim logPath = Path.Combine(_tmpRoot, "info.log")
        Dim logger = New Logger(logPath)

        logger.Write("hello")

        Assert.IsTrue(File.Exists(logPath))
        Dim lines = File.ReadAllLines(logPath, Encoding.UTF8)
        Assert.AreEqual(1, lines.Length)

        ' 例: 2026-02-10 13:45:12 [INFO] hello
        StringAssert.Contains(lines(0), " [INFO] ")
        StringAssert.Contains(lines(0), "hello")
        ' 日付フォーマットチェック（yyyy-MM-dd HH:mm）
        StringAssert.Matches(lines(0), New Regex("^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[INFO\] "))
    End Sub

    '==============================
    ' Write(level, message)
    '==============================
    <TestMethod>
    Public Sub Write_Level_Uppercased()
        Dim logPath = Path.Combine(_tmpRoot, "level.log")
        Dim logger = New Logger(logPath)

        logger.Write("warn", "disk low")

        Dim line = File.ReadAllText(logPath, Encoding.UTF8).TrimEnd()
        StringAssert.Contains(line, " [WARN] ")
        StringAssert.Contains(line, "disk low")
    End Sub

    '==============================
    ' Write(level, message, ex)
    '==============================
    <TestMethod>
    Public Sub Write_WithException_AppendsTypeAndMessage()
        Dim logPath = Path.Combine(_tmpRoot, "ex.log")
        Dim logger = New Logger(logPath)

        Dim ex As New InvalidOperationException("oops")
        logger.Write("error", "failed to run", ex)

        Dim line = File.ReadAllText(logPath, Encoding.UTF8).TrimEnd()
        ' 例: [ERROR] failed to run | InvalidOperationException: oops
        StringAssert.Contains(line, " [ERROR] ")
        StringAssert.Contains(line, "failed to run")
        StringAssert.Contains(line, "InvalidOperationException: oops")
    End Sub

    '==============================
    ' ディレクトリ自動作成
    '==============================
    <TestMethod>
    Public Sub Write_CreatesDirectoryIfMissing()
        Dim newDir = Path.Combine(_tmpRoot, "sub1\sub2\sub3")
        Dim logPath = Path.Combine(newDir, "auto.log")
        Dim logger = New Logger(logPath)

        logger.Write("auto dir")

        Assert.IsTrue(File.Exists(logPath))
        Dim line = File.ReadAllText(logPath, Encoding.UTF8).TrimEnd()
        StringAssert.Contains(line, " [INFO] ")
        StringAssert.Contains(line, "auto dir")
    End Sub

    '==============================
    ' スレッドセーフ：多重書き込み
    '==============================
    <TestMethod>
    Public Sub Write_IsThreadSafe_NoLostOrCorruptedLines()
        Dim logPath = Path.Combine(_tmpRoot, "parallel.log")
        Dim logger = New Logger(logPath)

        Dim threads As Integer = 8
        Dim perThread As Integer = 200
        Dim events(threads - 1) As Thread

        For t = 0 To threads - 1
            Dim idx = t
            events(idx) = New Thread(
                Sub()
                    For i = 1 To perThread
                        logger.Write("t" & idx & "-" & i.ToString())
                    Next
                End Sub)
            events(idx).IsBackground = True
            events(idx).Start()
        Next

        For t = 0 To threads - 1
            events(t).Join()
        Next

        Dim lines = File.ReadAllLines(logPath, Encoding.UTF8)
        ' 期待行数 8 * 200 = 1600
        Assert.AreEqual(threads * perThread, lines.Length)

        For Each ln In lines
            StringAssert.Contains(ln, " [INFO] ")
        Next
    End Sub

    '==============================
    ' フォールバック：I/O 失敗でも例外を出さない
    ' （ディレクトリパスをファイルとして渡して失敗を誘発）
    '==============================
    <TestMethod>
    Public Sub Write_OnIoFailure_FallsBackToConsole()

        ' 1) あえて「ファイル名無し＝ディレクトリそのもの」を出力先にする
        '    → AppendAllText がディレクトリへ書こうとして失敗（環境を問わず I/O 例外）
        Dim onlyDir = Path.Combine(_tmpRoot, "onlydir")
        Directory.CreateDirectory(onlyDir)

        Dim badPath = onlyDir            ' ★ ファイル名なし（ディレクトリそのもの）
        Dim logger = New OMS.Common.Logger(badPath)

        ' 2) 例外が出ない（= フォールバックで吸収）ことを検証
        Dim threw As Boolean = False
        Try
            logger.Write("it goes to console")
        Catch
            threw = True
        End Try
        Assert.IsFalse(threw, "I/O失敗時に例外がスローされました。フォールバックで吸収されていません。")

        ' 3) 実ファイルが生成/更新されていないことを検証
        '    → ディレクトリ直下にファイルが作成されていないことを確認
        Dim files = Directory.GetFiles(onlyDir)
        Assert.AreEqual(0, files.Length, "フォールバック時にファイルが作成されています。")
    End Sub

End Class