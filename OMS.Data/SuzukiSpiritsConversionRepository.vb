Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data

    Public Class SuzukiSpiritsConversionRepository
#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

#Region "一覧取得"
        ' スズキSPIRITS納入ホーム変換マスタ一覧取得
        Public Function GetSuzukiSpiritsConversionList(
            Optional ByVal deliveryCodePlan As String = Nothing,
            Optional ByVal deliveryCodeOrder As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  conversion_id          AS ""ConversionId"",")
            sb.AppendLine("  delivery_code_plan     AS ""deliveryCodePlan"",")
            sb.AppendLine("  delivery_code_order    AS ""deliveryCodeOrder"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM suzuki_spirits_conversion_list_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pDeliveryCodePlan As String = Utils.BuildLikePattern(deliveryCodePlan, LikeMode.Contains)
            Dim pDeliveryCodeOrder As String = Utils.BuildLikePattern(deliveryCodeOrder, LikeMode.Contains)
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

            If pDeliveryCodePlan IsNot Nothing Then
                sb.AppendLine("AND UPPER(delivery_code_plan) LIKE UPPER(:p_dcode_plan) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_dcode_plan", OracleDbType.Varchar2) With {.Value = pDeliveryCodePlan})
            End If

            If pDeliveryCodeOrder IsNot Nothing Then
                sb.AppendLine("AND UPPER(delivery_code_order) LIKE UPPER(:p_dcode_order) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_dcode_order", OracleDbType.Varchar2) With {.Value = pDeliveryCodeOrder})
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY delivery_code_order, delivery_code_plan ")

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

        Public Function GetSuzukiSpiritsConversion(ConversionId As Long) As DataTable
            Dim dt As New DataTable()

            Dim sql As String = "
                SELECT 
                    conversion_id       AS ""ConversionId"",
                    delivery_code_plan AS ""deliveryCodePlan"",
                    delivery_code_order AS ""deliveryCodeOrder"",
                    active_flag         AS ""ActiveFlag"",
                    created_at          AS ""CreatedAt"",
                    created_user_id     AS ""CreatedUserId"",
                    created_pg_id       AS ""CreatedPgId"",
                    updated_at          AS ""UpdatedAt"",
                    updated_user_id     As ""UpdatedUserId"",
                    updated_pg_id       AS ""UpdatedPgId""
                FROM suzuki_spirits_conversion_list_view
                WHERE conversion_id = :p_id
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = ConversionId
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

        Public Function ExistsSuzukiSpiritsConversion(
            deliveryCodeOrder As String,
            Optional excludeConversionId As Long = 0) As Boolean

            Dim sql As String = "
                SELECT 1
                FROM suzuki_spirits_conversion_list_view
                WHERE delivery_code_order = :p_dcode_order
            "

            If excludeConversionId > 0 Then
                sql &= "    AND conversion_id <> :p_exclude "
            End If

            sql &= " FETCH FIRST 1 ROWS ONLY "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_dcode_order", OracleDbType.Varchar2).Value = deliveryCodeOrder

                    ' 自レコード除外（更新時）
                    If excludeConversionId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeConversionId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function

#End Region

#Region "新規追加"

        Public Function InsertSuzukiSpiritsConversionNullable(
            deliveryCodePlan As String,
            deliveryCodeOrder As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO suzuki_spirits_conversion_mst (
                    delivery_code_plan,
                    delivery_code_order,
                    active_flag,
                    created_user_id,
                    created_pg_id,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    :p_dcode_plan,
                    :p_dcode_order,
                    :p_active,
                    :p_user,
                    :p_pg,
                    :p_user,
                    :p_pg
                )
                RETURNING conversion_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True

                        cmd.Parameters.Add(":p_dcode_plan", OracleDbType.Varchar2).Value = deliveryCodePlan
                        cmd.Parameters.Add(":p_dcode_order", OracleDbType.Varchar2).Value = deliveryCodeOrder

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

#End Region

#Region "更新（排他：updated_at一致時のみ）"

        Public Function UpdateSuzukiSpiritsConversionNullable(
            conversionId As Long,
            deliveryCodePlan As String,
            deliveryCodeOrder As String,
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
                        FROM suzuki_spirits_conversion_mst
                        WHERE conversion_id = :p_id
                    ", conn)
                        cmdChk.Transaction = tran
                        cmdChk.BindByName = True
                        cmdChk.Parameters.Add(":p_id", OracleDbType.Int64).Value = conversionId

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
                        UPDATE suzuki_spirits_conversion_mst
                        SET delivery_code_plan = :p_delivery_plan,
                            delivery_code_order = :p_delivery_order,
                            active_flag         = :p_active,
                            updated_at          = SYSDATE,
                            updated_user_id     = :p_user,
                            updated_pg_id       = :p_pg
                        WHERE conversion_id       = :p_id
                            AND updated_at      = :p_currupd
                    "

                    Dim affected As Integer = 0

                    Using cmdUpd As New OracleCommand(sqlUpd, conn)
                        cmdUpd.Transaction = tran
                        cmdUpd.BindByName = True

                        cmdUpd.Parameters.Add(":p_delivery_plan", OracleDbType.Varchar2).Value = deliveryCodePlan
                        cmdUpd.Parameters.Add(":p_delivery_order", OracleDbType.Varchar2).Value = deliveryCodeOrder

                        cmdUpd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmdUpd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmdUpd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmdUpd.Parameters.Add(":p_id", OracleDbType.Int64).Value = conversionId
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
