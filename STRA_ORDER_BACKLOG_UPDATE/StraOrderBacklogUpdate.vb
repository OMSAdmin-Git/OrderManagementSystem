Imports System.IO
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess
Imports Oracle.ManagedDataAccess.Client
Imports System.Windows.Forms
Imports System.Xml.Linq

Module StraOrderBacklogUpdate

    ' 受注残 内示受注残数削除 定時処理

    ' 毎日 21:00 にタスクスケジューラーより起動されます。
    ' STRAMMICの受注残情報Viewから情報を取得、受注テーブルへ反映更新する。
    ' 受注テーブルの受注残数が０以上かつSTRAMMIC受注残情報が取得できない受注データは、
    ' STRAMMICの出荷が完了(受注残数０)と判断し、受注テーブルの受注残数を[0](ゼロ)に更新する。
    ' ※STRAMMIC側で出荷完了した場合、受注残情報VIEWにデータが表示されなくなる仕様

    'アセンブリ名:STRA_ORDER_BACKLOG_UPDATE
    'ルート名前空間:STRA_ORDER_BACKLOG_UPDATE
    'フレームワーク: .NET Framework 4.8
    '種類:コンソールアプリケーション
    '環境
    '参照 アセンブリ
    'System
    'System.Core
    'System.Data
    'System.Data.DataSetExtensions
    'System.Deployment
    'System.Net.Http
    'System.Numeric
    'System.Windows.Forms
    'System.Xml
    'System.Xml.Linq
    '
    '参照 プロジェクト
    'OMS.Common
    'OMS.Data
    '
    ' スケジューラー設定

    Private Const LogFilename As String = "StraOrderBacklog.txt"

    Function Main() As Integer
        Dim rt As Integer = 0
        Dim errors As List(Of String) = New List(Of String)()
        ' 設定
        Dim xmlDoc As XDocument = XDocument.Load("Config.xml")
        Dim userId = ""
        Dim logPath = ""
        Dim _logger As Logger = Nothing
        Try
            For Each StraOrderBacklogUpdate In xmlDoc.Descendants("StraOrderBacklogUpdate")
                userId = StraOrderBacklogUpdate.Element("UserId")?.Value
                logPath = StraOrderBacklogUpdate.Element("LogPath")?.Value
            Next
            If (userId = "" Or logPath = "") Then
                Throw New Exception("Configration  Error")
            End If

            ' INITIAL
            Dim wkUpdatedAt = DateTime.Now              ' WK_UPDATED_AT(更新日時)
            Dim wkUpdatedUserId = "BkOfOdr"             ' WK_UPDATED_USER_ID(更新ユーザーID)
            Dim wkUpdatedPgId = "Backlog of Orders"     ' WK_UPDATED_PG_ID(更新プログラムID)

            ' Log
            _logger = New Logger(Path.Combine(logPath, LogFilename))
            EnsureDirectory(logPath)
            _logger.Write($"StraOrderBacklogUpdate.exe Start: {wkUpdatedAt.ToString()}")

            Dim repo = New StraOrderBackLog(userId)

            ' UPDATE(受注残更新)
            errors.Add(repo.UpdateOrderBack(wkUpdatedAt, wkUpdatedUserId, wkUpdatedPgId))

            ' UPDATE(受注残ゼロ更新)
            errors.Add(repo.UpdateOrderBackZero(wkUpdatedAt, wkUpdatedUserId, wkUpdatedPgId))

            If (0 = errors.FindAll(Function(x) x <> "").Count) Then
                _logger.Write($"StraOrderBacklogUpdate.exe Complete: {DateTime.Now.ToString()}")
            Else
                Throw New Exception("Update Error")
            End If

        Catch ex As Exception
            errors.Add(ex.Message)
            Dim ms As String = String.Join(vbCrLf, errors.Where(Function(x) x <> ""))
            _logger?.Write($"StraOrderBacklogUpdate.exe Error: {ms}")
            MessageBox.Show(ms, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rt = -1
        End Try
        Return rt

    End Function

End Module
