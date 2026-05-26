Imports System.IO
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess
Imports Oracle.ManagedDataAccess.Client
Imports System.Windows.Forms

Module StraOrderBacklogUpdate

    ' STRAMMICの受注残情報Viewから情報を取得、受注テーブルへ反映更新する。
    ' 受注テーブルの受注残数が０以上かつSTRAMMIC受注残情報が取得できない受注データは、
    ' STRAMMICの出荷が完了(受注残数０)と判断し、受注テーブルの受注残数を[0](ゼロ)に更新する。
    ' ※STRAMMIC側で出荷完了した場合、受注残情報VIEWにデータが表示されなくなる使用

    Sub Main()

        ' INITIAL
        Dim wkUpdatedAt = DateTime.Now              ' WK_UPDATED_AT(更新日時)
        Dim wkUpdatedUserId = "BkOfOdr"             ' WK_UPDATED_USER_ID(更新ユーザーID)
        Dim wkUpdatedPgId = "Backlog of Orders"     ' WK_UPDATED_PG_ID(更新プログラムID)

        ' 天方環境
        Dim UserId = "User Id=OMSDB;Password=Amagata001;Data Source=//192.168.10.15:1521/OMSDB;"
        'アスカ環境
        'Dim UserId = "User Id=OMSDB;Password=Amagata001;Data Source=//192.168.100.126:1521/orcl;"
        '本番環境
        'Dim UserId = "User Id=OMSDB;Password=Amagata001;Data Source=//192.168.70.225:1521/OMSDB;"
        '検証環境
        'Dim UserId = "User Id=OMSTS;Password=Amagata001;Data Source=//192.168.70.225:1521/OMSDB;"

        ' Log
        Dim logPath = "C:\ASTI\StraOrderBacklogUpdate"
        Dim _logger As Logger = New Logger(logPath)
        EnsureDirectory(logPath)
        _logger.Write($"StraOrderBacklogUpdate.exe Start: {wkUpdatedAt.ToString()}")

        Dim errors As List(Of String) = New List(Of String)()
        Dim repo = New StraOrderBackLog(UserId)

        Try
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
            Dim ms As String = String.Join(vbCrLf, errors.Where(Function(x) x <> ""))
            _logger.Write($"StraOrderBacklogUpdate.exe Error: {ms}")
            MessageBox.Show(ms, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Module
