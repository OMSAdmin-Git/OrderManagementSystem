Imports System.IO
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess
Imports Oracle.ManagedDataAccess.Client

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
        Dim UserId = "User Id=OMSDB;Password=Amagata001;Data Source=//192.168.100.126:1521/orcl;"
        ' Log
        Dim logPath = "C:\ASTI\StraOrderBacklogUpdate"
        Dim _logger As Logger = New Logger(logPath)
        EnsureDirectory(logPath)
        _logger.Write($"StraOrderBacklogUpdate.exe Start: {wkUpdatedAt.ToString()}")

        ' UPDATE(受注残更新)






#If False Then
        Dim logPath = "C:\Temp"
        Dim UserId = "User Id=OMSDB;Password=Amagata001;Data Source=//192.168.100.126:1521/orcl;"
        Dim _logger As Logger = New Logger(logPath)
        Try
            _logger.Write($"StraOrderBacklogUpdate.exe Get OrderBack: {DateTime.Now.ToString()}")

            Dim repo = New OrderRepository(UserId)
            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(UserId)
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()


            Dim dt = repo.GetOrders(conn, tran, OrderRepository.OrdersTable.Orders, status:="DUE_SET")

            Dim cnt = dt.Rows.Count


            Dim dr = repo.GetOrders(conn, tran, OrderRepository.OrdersTable.Orders, status:="DUE_SET")

            Dim dc = repo.ToClass(dr)
            Dim dct = dc.Count

        Catch ex As Exception
            _logger.Write($"StraOrderBacklogUpdate.exe error: {ex.Message}")

        End Try

#End If


    End Sub

End Module
