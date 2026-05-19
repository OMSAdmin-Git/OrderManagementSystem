Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    Public Class InfoTypeRepository

#Region "フィールド・コンストラクタ"

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

#End Region

#Region "一覧取得"

        ''' <summary>
        ''' 情報区分マスタ一覧取得
        ''' </summary>
        Public Function GetInfoTypeList(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal prodMgmtUserId As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  info_type_id           AS ""InfoTypeId"",")
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  customer_info_type     AS ""CustomerInfoType"",")
            sb.AppendLine("  info_type              AS ""InfoType"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM info_type_list_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pCustomerCode As String = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName As String = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter As String = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName As String = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId As String = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

            If pCustomerCode IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_code) LIKE UPPER(:p_ccode) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_ccode", OracleDbType.Varchar2) With {.Value = pCustomerCode})
            End If

            If pCustomerName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_name) LIKE UPPER(:p_cname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cname", OracleDbType.Varchar2) With {.Value = pCustomerName})
            End If

            If pProfitCenter IsNot Nothing Then
                sb.AppendLine("AND UPPER(profit_center) LIKE UPPER(:p_pc) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_pc", OracleDbType.Varchar2) With {.Value = pProfitCenter})
            End If

            If pCustomerUnitName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_unit_name) LIKE UPPER(:p_cuname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cuname", OracleDbType.Varchar2) With {.Value = pCustomerUnitName})
            End If

            If pProdMgmtUserId IsNot Nothing Then
                Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
                If Not isAdmin Then
                    'sb.AppendLine("AND UPPER(prod_mgmt_user_id) = UPPER(:p_user) ")
                    sb.AppendLine("AND UPPER(prod_mgmt_user_id) LIKE UPPER(:p_user) ")
                    prm.Add(New OracleParameter(":p_user", OracleDbType.Varchar2) With {.Value = pProdMgmtUserId})
                End If
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY customer_code, profit_center, customer_unit_id, info_type")

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
        Public Function GetInfoType(infoTypeId As Long) As DataTable
            Dim dt As New DataTable()

            Dim sql As String = "
                SELECT 
                    info_type_id        AS ""InfoTypeId"",
                    customer_setting_id AS ""CustomerSettingId"",
                    customer_code       AS ""CustomerCode"",
                    customer_name       AS ""CustomerName"",
                    profit_center       AS ""ProfitCenter"",
                    customer_unit_id    AS ""CustomerUnitId"",
                    customer_unit_name  AS ""CustomerUnitName"",
                    customer_info_type  AS ""CustomerInfoType"",
                    info_type           AS ""InfoType"",
                    active_flag         AS ""ActiveFlag"",
                    created_at          AS ""CreatedAt"",
                    created_user_id     AS ""CreatedUserId"",
                    created_pg_id       AS ""CreatedPgId"",
                    updated_at          AS ""UpdatedAt"",
                    updated_user_id     As ""UpdatedUserId"",
                    updated_pg_id       AS ""UpdatedPgId""
                FROM info_type_list_view
                WHERE info_type_id = :p_id
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = infoTypeId
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

        Public Function ExistsInfoType(
            customerSettingId As Long,
            customerInfoType As String,
            Optional excludeInfoTypeId As Long = 0
        ) As Boolean

            Dim sql As String = "
                SELECT 1
                FROM info_type_mst
                WHERE customer_setting_id = :p_custid
                    AND customer_info_type = :p_type
            "

            If excludeInfoTypeId > 0 Then
                sql &= "    AND info_type_id <> :p_exclude "
            End If

            sql &= " FETCH FIRST 1 ROWS ONLY "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                    cmd.Parameters.Add(":p_type", OracleDbType.Varchar2).Value = customerInfoType

                    ' 自レコード除外（更新時）
                    If excludeInfoTypeId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeInfoTypeId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function
#End Region

#Region "INSERT / Update（通常版）"

        Public Function InsertInfoType(
            infoTypeId As Long,
            customerSettingId As Long,
            customerInfoType As String,
            infoType As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
            INSERT INTO info_type_mst (
                info_type_id,
                customer_setting_id,
                customer_info_type,
                info_type,
                active_flag,
                created_at,
                created_user_id,
                created_pg_id,
                updated_at,
                updated_user_id,
                updated_pg_id
            ) VALUES (
                INFO_TYPE_SEQ.NEXTVAL,
                :p_custid,
                :p_ctype,
                :p_type,
                :p_active,
                SYSDATE,
                :p_user,
                :p_pg,
                SYSDATE,
                :p_user,
                :p_pg
            )
            RETURNING info_type_id INTO :p_newid
        "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_ctype", OracleDbType.Varchar2).Value = customerInfoType
                        cmd.Parameters.Add(":p_type", OracleDbType.Varchar2).Value = infoType

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

        Public Function InsertInfoTypeNullable(
            customerSettingId As Long,
            customerInfoType As String,
            infoType As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO info_type_mst (
                    customer_setting_id,
                    customer_info_type,
                    info_type,
                    active_flag,
                    created_user_id,
                    created_pg_id,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    :p_custid,
                    :p_ctype,
                    :p_type,
                    :p_active,
                    :p_user,
                    :p_pg,
                    :p_user,
                    :p_pg
                )
                RETURNING info_type_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True

                        cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_ctype", OracleDbType.Varchar2).Value = customerInfoType
                        cmd.Parameters.Add(":p_type", OracleDbType.Varchar2).Value = infoType

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

        Public Function UpdateInfoType(
            infoTypeId As Long,
            customerInfoType As String,
            infoType As String,
            activeFlag As String,
            loginUserId As String,
            programId As String,
            originalUpdatedAt As DateTime
        ) As Integer

            Dim sql As String = "
                UPDATE info_type_mst
                SET customer_info_type  = :p_ctype,
                    info_type           = :p_type,
                    active_flag         = :p_active,
                    updated_at          = SYSDATE,
                    updated_user_id     = :p_user,
                    updated_pg_id       = :p_pg
                WHERE info_type_id      = :p_id
                    AND updated_at      = :p_origupd
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_ctype", OracleDbType.Varchar2).Value = customerInfoType
                        cmd.Parameters.Add(":p_type", OracleDbType.Varchar2).Value = infoType

                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = infoTypeId
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

        Public Function UpdateInfoTypeWithConcurrency(
            infoTypeId As Long,
            customerInfoType As String,
            infoType As String,
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
                        FROM info_type_mst
                        WHERE info_type_id = :p_id
                    ", conn)
                        cmdChk.Transaction = tran
                        cmdChk.BindByName = True
                        cmdChk.Parameters.Add(":p_id", OracleDbType.Int64).Value = infoTypeId

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
                        UPDATE info_type_mst
                        SET customer_info_type  = :p_ctype,
                            info_type           = :p_type,
                            active_flag         = :p_active,
                            updated_at          = SYSDATE,
                            updated_user_id     = :p_user,
                            updated_pg_id       = :p_pg
                        WHERE info_type_id      = :p_id
                            AND updated_at      = :p_currupd
                    "

                    Dim affected As Integer = 0

                    Using cmdUpd As New OracleCommand(sqlUpd, conn)
                        cmdUpd.Transaction = tran
                        cmdUpd.BindByName = True

                        cmdUpd.Parameters.Add(":p_ctype", OracleDbType.Varchar2).Value = customerInfoType
                        cmdUpd.Parameters.Add(":p_type", OracleDbType.Varchar2).Value = infoType

                        cmdUpd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmdUpd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmdUpd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmdUpd.Parameters.Add(":p_id", OracleDbType.Int64).Value = infoTypeId
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
