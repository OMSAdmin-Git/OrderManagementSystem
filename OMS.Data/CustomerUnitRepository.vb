Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    Public Class CustomerUnitRepository

#Region "フィールド・コンストラクタ"

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

#End Region

#Region "一覧取得"

        ' 注文工場／担当者マスタ一覧取得
        Public Function GetCustomerUnitList(
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  customer_unit_id   AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name AS ""CustomerUnitName"",")
            sb.AppendLine("  active_flag        AS ""ActiveFlag"",")
            sb.AppendLine("  created_at         AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id    AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id      AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at         AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id    AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id      AS ""UpdatedPgId""")
            sb.AppendLine("FROM customer_unit_mst ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' LIKE 検索
            Dim pCustomerUnitName As String = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

            If pCustomerUnitName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_unit_name) LIKE UPPER(:p_cuname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cuname", OracleDbType.Varchar2) With {.Value = pCustomerUnitName})
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY customer_unit_id ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt

        End Function

#End Region

#Region "単一取得"

        Public Function GetCustomerUnit(customerUnitId As Long) As DataTable
            Dim dt As New DataTable()

            Dim sql As String = "
                SELECT 
                    customer_unit_id    AS ""CustomerUnitId"",
                    customer_unit_name  AS ""CustomerUnitName"",
                    active_flag         AS ""ActiveFlag"",
                    created_at          AS ""CreatedAt"",
                    created_user_id     AS ""CreatedUserId"",
                    created_pg_id       AS ""CreatedPgId"",
                    updated_at          AS ""UpdatedAt"",
                    updated_user_id     As ""UpdatedUserId"",
                    updated_pg_id       AS ""UpdatedPgId""
                FROM customer_unit_mst
                WHERE customer_unit_id = :p_id
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerUnitId
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

#End Region

#Region "重複チェック"
        ''' <summary>
        ''' 注文工場／担当者名の重複をチェックします（大文字小文字を無視、必要に応じて自IDを除外）。
        ''' </summary>
        ''' <param name="customerUnitName">チェック対象の名称（CUSTOMER_UNIT_NAME）</param>
        ''' <param name="excludeCustomerUnitId">更新時など自レコードを除外したい場合のID（新規時は0）</param>
        ''' <returns>重複が存在すれば True、存在しなければ False</returns>
        Public Function ExistsCustomerUnit(
            customerUnitName As String,
            Optional excludeCustomerUnitId As Long = 0
        ) As Boolean

            ' NULL/空はそもそも重複チェックの対象外（呼び出し側で必須チェックするのが基本）
            If String.IsNullOrWhiteSpace(customerUnitName) Then
                Return False
            End If

            Dim sql As New StringBuilder()
            sql.AppendLine("SELECT 1 ")
            sql.AppendLine("FROM customer_unit_mst ")
            sql.AppendLine("WHERE UPPER(customer_unit_name) = UPPER(:p_name) ")

            If excludeCustomerUnitId > 0 Then
                sql.AppendLine("    AND CUSTOMER_UNIT_ID <> :p_exclude ")
            End If

            sql.AppendLine("FETCH FIRST 1 ROWS ONLY ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql.ToString(), conn)
                    cmd.BindByName = True

                    ' 注文工場／担当者名
                    cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName

                    If excludeCustomerUnitId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeCustomerUnitId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function
#End Region

#Region "INSERT / Update（通常版）"

        ''' <summary>
        ''' 注文工場／担当者（CUSTOMER_UNIT_MST）を登録します。
        ''' </summary>
        ''' <param name="customerUnitId">
        ''' 画面から受け取ったID。<br/>
        ''' 現行SQLでは <c>CUSTOMER_UNIT_SEQ.NEXTVAL</c> による採番を使用しており、VALUES の2番目の <c>:p_id</c> は
        ''' 実質的に <c>customer_unit_name</c> の位置に対応します（カラム順に注意）。<br/>
        ''' ※ 採番を全てシーケンスで行う設計の場合、引数そのものを廃止し、SQLの <c>:p_id</c> も削除するのが自然です。
        ''' </param>
        ''' <param name="customerUnitName">名称（CUSTOMER_UNIT_NAME）</param>
        ''' <param name="activeFlag">有効フラグ（'Y' or 'N'）</param>
        ''' <param name="loginUserId">作成/更新ユーザーID</param>
        ''' <param name="programId">作成/更新プログラムID</param>
        ''' <returns>採番された <c>customer_unit_id</c> を返します。</returns>
        ''' <remarks>
        ''' ・監査項目（created_*/updated_*）は本メソッドで自動設定します。<br/>
        ''' ・戻り値は <c>RETURNING customer_unit_id INTO :p_newid</c> により取得します。
        ''' </remarks>
        Public Function InsertCustomerUnit(
            customerUnitId As Long,
            customerUnitName As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
            INSERT INTO customer_unit_mst (
                customer_unit_id,
                customer_unit_name,
                active_flag,
                created_at,
                created_user_id,
                created_pg_id,
                updated_at,
                updated_user_id,
                updated_pg_id
            ) VALUES (
                CUSTOMER_UNIT_SEQ.NEXTVAL,
                :p_name,
                :p_active,
                SYSDATE,
                :p_user,
                :p_pg,
                SYSDATE,
                :p_user,
                :p_pg
            )
            RETURNING customer_unit_id INTO :p_newid
        "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName

                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        Dim pNewId = New OracleParameter(":p_newid", OracleDbType.Int64)
                        pNewId.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(pNewId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()
                        Return CLng(pNewId.Value)
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' 注文工場／担当者（CUSTOMER_UNIT_MST）を登録します（IDはDB側で採番、NULL許容項目のみを対象）。
        ''' </summary>
        ''' <param name="customerUnitName">名称（CUSTOMER_UNIT_NAME）</param>
        ''' <param name="activeFlag">有効フラグ（'Y' or 'N'）</param>
        ''' <param name="loginUserId">作成/更新ユーザーID</param>
        ''' <param name="programId">作成/更新プログラムID</param>
        ''' <returns>採番された <c>customer_unit_id</c> を返します。</returns>
        ''' <remarks>
        ''' ・本メソッドは列指定を絞り、DB側のデフォルト/トリガーに依存しやすい形です。<br/>
        ''' ・戻り値は <c>RETURNING customer_unit_id INTO :p_newid</c> により取得します。
        ''' </remarks>
        Public Function InsertCustomerUnitNullable(
            customerUnitName As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO customer_unit_mst (
                    customer_unit_name,
                    active_flag,
                    created_user_id,
                    created_pg_id,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    :p_name,
                    :p_active,
                    :p_user,
                    :p_pg,
                    :p_user,
                    :p_pg
                )
                RETURNING customer_unit_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True

                        cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName

                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        Dim pNewId As New OracleParameter(":p_newid", OracleDbType.Int64)
                        pNewId.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(pNewId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

                        If pNewId.Value Is Nothing OrElse pNewId.Value Is DBNull.Value Then
                            Throw New ApplicationException("採番IDの取得に失敗しました。")
                        End If

                        If TypeOf pNewId.Value Is OracleDecimal Then
                            Dim od As OracleDecimal = DirectCast(pNewId.Value, OracleDecimal)
                            Return od.ToInt64()
                        Else
                            Return Convert.ToInt64(pNewId.Value.ToString())
                        End If
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' 注文工場／担当者（CUSTOMER_UNIT_MST）を更新します（排他制御：<c>updated_at</c> 一致時のみ）。
        ''' </summary>
        ''' <param name="customerUnitId">主キー（CUSTOMER_UNIT_ID）</param>
        ''' <param name="customerUnitName">名称（CUSTOMER_UNIT_NAME）</param>
        ''' <param name="activeFlag">有効フラグ（'Y' or 'N'）</param>
        ''' <param name="loginUserId">更新ユーザーID</param>
        ''' <param name="programId">更新プログラムID</param>
        ''' <param name="originalUpdatedAt">
        ''' 画面表示時点の <c>updated_at</c>。<br/>
        ''' 本値とDB上の <c>updated_at</c> が一致した場合のみ更新を許可します。
        ''' </param>
        ''' <returns>
        ''' 更新件数（通常 1）。<br/>
        ''' 0 の場合は排他エラー（他セッションにより更新済み）を示します。
        ''' </returns>
        ''' <remarks>
        ''' ・<c>updated_at</c> が DATE か TIMESTAMP かに応じて、パラメータ型（<c>OracleDbType.Date</c> / <c>OracleDbType.TimeStamp</c>）を合わせてください。<br/>
        ''' ・画面側で <c>originalUpdatedAt</c> を Hidden に保持して渡す運用を想定します。
        ''' </remarks>
        Public Function UpdateCustomerUnit(
            customerUnitId As Long,
            customerUnitName As String,
            activeFlag As String,
            loginUserId As String,
            programId As String,
            originalUpdatedAt As DateTime
        ) As Integer

            Dim sql As String = "
            UPDATE customer_unit_mst
               SET customer_unit_name   = :p_name,
                   active_flag          = :p_active,
                   updated_at           = SYSDATE,
                   updated_user_id      = :p_user,
                   updated_pg_id        = :p_pg
             WHERE customer_unit_id     = :p_id
               AND updated_at           = :p_origupd
        "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName

                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerUnitId
                        cmd.Parameters.Add(":p_origupd", OracleDbType.Date).Value = originalUpdatedAt

                        Dim affected = cmd.ExecuteNonQuery()

                        tran.Commit()
                        Return affected
                    End Using
                End Using
            End Using
        End Function

#End Region

#Region "更新（排他：updated_at一致時のみ）"
        Public Function UpdateCustomerUnitWithConcurrency(
            customerUnitId As Long,
            customerUnitName As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Integer

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()

                    ' 現在の updated_at を取得
                    Dim currentUpdatedAt As DateTime
                    Dim exists As Boolean = False
                    Using cmdChk As New OracleCommand("
                        SELECT updated_at
                        FROM customer_unit_mst
                        WHERE customer_unit_id = :p_id
                    ", conn)
                        cmdChk.Transaction = tran
                        cmdChk.BindByName = True
                        cmdChk.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerUnitId

                        Using rdr = cmdChk.ExecuteReader()
                            If rdr.Read() Then
                                If Not rdr.IsDBNull(0) Then currentUpdatedAt = rdr.GetDateTime(0)
                                exists = True
                            End If
                        End Using
                    End Using

                    If Not exists Then
                        tran.Rollback()
                        Return -1
                    End If

                    ' 排他 UPDATE
                    Dim sqlUpd As String = "
                        UPDATE customer_unit_mst
                           SET customer_unit_name   = :p_name,
                               active_flag          = :p_active,
                               updated_at           = SYSDATE,
                               updated_user_id      = :p_user,
                               updated_pg_id        = :p_pg
                         WHERE customer_unit_id     = :p_id
                           AND updated_at           = :p_currupd
                    "

                    Dim affected As Integer = 0

                    Using cmdUpd As New OracleCommand(sqlUpd, conn)
                        cmdUpd.Transaction = tran
                        cmdUpd.BindByName = True
                        cmdUpd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName

                        cmdUpd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmdUpd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmdUpd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmdUpd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerUnitId
                        cmdUpd.Parameters.Add(":p_currupd", OracleDbType.Date).Value = currentUpdatedAt

                        affected = cmdUpd.ExecuteNonQuery()
                    End Using

                    If affected = 0 Then
                        tran.Rollback()
                        Return 0
                    End If

                    tran.Commit()
                    Return affected

                End Using
            End Using
        End Function
#End Region

    End Class

End Namespace