Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    Public Class CustomerRepository

#Region "フィールド・コンストラクタ"

        Private ReadOnly _connectionString As String
        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

#End Region

#Region "一覧取得"

        ' 取引先設定マスタ一覧取得
        Public Function GetCustomerList(
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
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  user_name              AS ""UserName"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM customer_list_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' LIKE 検索
            Dim pCustomerCode = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)
            Dim pActiveFlag = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

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

            If pActiveFlag IsNot Nothing Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY customer_code, profit_center, customer_unit_id")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

        ' 取引先設定マスタ一覧取得
        Public Function GetCustomerImpRuleList(
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
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  NVL(TRIM(reconcile_flag), 'Y') AS ""ReconcileFlag"",")
            sb.AppendLine("  NVL(TRIM(fcst_reconcile_flag), 'Y') AS ""FcstReconcileFlag"",")
            sb.AppendLine("  NVL(TRIM(reconcile_type), 1) AS ""ReconcileType"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  user_name              AS ""UserName"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM (
                                SELECT
                                    v.*,
                                    ROW_NUMBER() OVER (
                                        PARTITION BY customer_setting_id
                                        ORDER BY updated_at DESC NULLS LAST
                                    ) AS rn
                                FROM customer_imp_rule_list_view v
                            )
                            ")
            sb.AppendLine("WHERE rn = 1 ")

            Dim prm As New List(Of OracleParameter)()

            ' LIKE 検索
            Dim pCustomerCode = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)
            Dim pActiveFlag = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

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

            If pActiveFlag IsNot Nothing Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY customer_code, profit_center, customer_unit_id")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

#End Region

#Region "単一取得"

        Public Function GetCustomerSetting(customerSettingId As Long) As DataTable
            Dim dt As New DataTable()

            Dim sql As String = "
                SELECT 
                    customer_setting_id AS ""CustomerSettingId"",
                    customer_code       AS ""CustomerCode"",
                    customer_name       AS ""CustomerName"",
                    profit_center       AS ""ProfitCenter"",
                    customer_unit_id    AS ""CustomerUnitId"",
                    customer_unit_name  AS ""CustomerUnitName"",
                    prod_mgmt_user_id   AS ""ProdMgmtUserId"",
                    user_name           AS ""UserName"",
                    active_flag         AS ""ActiveFlag"",
                    created_at          AS ""CreatedAt"",
                    created_user_id     AS ""CreatedUserId"",
                    created_pg_id       AS ""CreatedPgId"",
                    updated_at          AS ""UpdatedAt"",
                    updated_user_id     As ""UpdatedUserId"",
                    updated_pg_id       AS ""UpdatedPgId""
                FROM customer_list_view
                WHERE customer_setting_id = :p_id
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

        Public Function FindCustomerSettingIdSimple(
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long?
        ) As Long?

            Dim pc As Object = If(profitCenter Is Nothing, DBNull.Value, CType(profitCenter, Object))
            Dim cuid As Object = If(customerUnitId.HasValue, CType(customerUnitId.Value, Object), DBNull.Value)

            Dim sql As String = "
                SELECT MIN(customer_setting_id)
                FROM customer_setting_mst
                WHERE UPPER(customer_code) = UPPER(:p_ccode)
                    AND NVL(UPPER(profit_center), CHR(0)) = NVL(UPPER(:p_pc), CHR(0))
                    AND NVL(customer_unit_id, -1) = NVL(:p_cuid, -1)
            "
            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode

                    Dim pPc As New OracleParameter(":p_pc", OracleDbType.Varchar2) With {.Value = pc}
                    cmd.Parameters.Add(pPc)

                    Dim pCuid As New OracleParameter(":p_cuid", OracleDbType.Int64) With {.Value = cuid}
                    cmd.Parameters.Add(pCuid)

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then Return Nothing
                    Return Convert.ToInt64(obj)
                End Using
            End Using
        End Function

#End Region

#Region "候補リスト取得"

        ' 取引先コード一覧取得
        Public Function GetCustomerCodes(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT DISTINCT customer_code ")
            sb.AppendLine("FROM customer_setting_mst ")
            sb.AppendLine("WHERE active_flag = 'Y' ")
            Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
            If Not isAdmin Then
                sb.AppendLine("AND prod_mgmt_user_id = :p_user ")
            End If
            sb.AppendLine("ORDER BY customer_code ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If Not isAdmin Then
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = prodMgmtUserId
                    End If
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

        Public Function GetStraCustomerCodes(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT DISTINCT m.fsectcd ")
            sb.AppendLine("FROM sectm m ")
            sb.AppendLine("LEFT JOIN sectd d ON m.fsectcd = d.fsectcd ")
            sb.AppendLine("WHERE d.fsecttyp = 'CU' ")
            sb.AppendLine("ORDER BY m.fsectcd ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

        Public Function GetCustomerNames(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT DISTINCT customer_name ")
            sb.AppendLine("FROM customer_list_view ")
            sb.AppendLine("WHERE active_flag = 'Y' ")
            Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
            If Not isAdmin Then
                sb.AppendLine("AND prod_mgmt_user_id = :p_user ")
            End If
            sb.AppendLine("ORDER BY customer_name ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If Not isAdmin Then
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = prodMgmtUserId
                    End If
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

        ' PC一覧取得
        Public Function GetProfitCenters(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT DISTINCT profit_center ")
            sb.AppendLine("FROM customer_setting_mst ")
            sb.AppendLine("WHERE active_flag = 'Y' ")
            Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
            If Not isAdmin Then
                sb.AppendLine("AND prod_mgmt_user_id = :p_user ")
            End If
            sb.AppendLine("ORDER BY profit_center")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If Not isAdmin Then
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = prodMgmtUserId
                    End If
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

        'STRAMMIC側　PC一覧取得
        Public Function GetStraProfitCenters(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sb As New System.Text.StringBuilder()
            'sb.AppendLine("SELECT DISTINCT fusrstr1 ")
            'sb.AppendLine("FROM usrdeffldf ")
            'sb.AppendLine("ORDER BY fusrstr1")

            sb.AppendLine("SELECT DISTINCT SECTM.FSECTCD ")
            sb.AppendLine("FROM SECTM inner join SECTD on SECTM.FSECTCD = SECTD.FSECTCD and SECTD.FSECTTYP = 'CH' ")
            sb.AppendLine("WHERE 1 = 1 ")
            sb.AppendLine("ORDER BY SECTM.FSECTCD")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

        ' 注文工場／担当者一覧取得
        Public Function GetCustomerUnitNames(prodMgmtUserId As String) As List(Of String)
            Dim result As New List(Of String)()

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT DISTINCT customer_unit_name ")
            sb.AppendLine("FROM customer_list_view ")
            sb.AppendLine("WHERE active_flag = 'Y' ")
            Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
            If Not isAdmin Then
                sb.AppendLine("AND prod_mgmt_user_id = :p_user ")
            End If

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    If Not isAdmin Then
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = prodMgmtUserId
                    End If
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader.IsDBNull(0) Then Continue While
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            Return result

        End Function

#End Region

#Region "マスタチェック"

        ' SECTM.取引先コードの存在チェック
        Public Function ExistsCustomerCode(customerCode As String) As Boolean
            Const sql As String = "
                SELECT 1
                FROM sectm m
                LEFT JOIN sectd d ON m.fsectcd = d.fsectcd
                WHERE m.fsectcd = :code
                    AND d.fsecttyp = 'CU'
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.Parameters.Add(":code", OracleDbType.Varchar2).Value = customerCode
                    conn.Open()
                    Return cmd.ExecuteScalar() IsNot Nothing
                End Using
            End Using
        End Function

        ' USRDEFFLDF.PCの存在チェック
        ' SECTM.PCの存在チェック
        Public Function ExistsProfitCenter(profitCenter As String) As Boolean
            'Const sql As String = "
            '    SELECT 1
            '    FROM usrdeffldf
            '    WHERE fusrstr1 = :pc
            '"
            Const sql As String = "
                SELECT 1
                FROM SECTM inner join SECTD on SECTM.FSECTCD = SECTD.FSECTCD and SECTD.FSECTTYP = 'CH'
                WHERE 1 = 1 and SECTM.FSECTCD = :pc
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.Parameters.Add(":pc", OracleDbType.Varchar2).Value = profitCenter
                    conn.Open()
                    Return cmd.ExecuteScalar() IsNot Nothing
                End Using
            End Using
        End Function

#End Region

#Region "重複チェック"
        ''' <summary>
        ''' 取引先設定（CUSTOMER_SETTING_MST）の重複をチェックします。
        ''' </summary>
        ''' <param name="customerCode">取引先コード（CUSTOMER_CODE）。大文字小文字は無視して比較します。</param>
        ''' <param name="profitCenter">PC（PROFIT_CENTER）。NULL 許容。値が NULL の場合は DB 側の NULL と一致したときのみマッチします。</param>
        ''' <param name="customerUnitId">注文工場／担当者ID（CUSTOMER_UNIT_ID）。NULL 許容。値が NULL の場合は DB 側の NULL と一致したときのみマッチします。</param>
        ''' <param name="excludeCustomerSettingId">更新時など、自レコードを重複判定から除外したい場合の ID。新規時は 0。</param>
        ''' <returns>
        ''' 重複が存在する場合は <c>True</c>、存在しない場合は <c>False</c> を返します。
        ''' </returns>
        ''' <remarks>
        ''' 一意性キーの想定は「<c>customer_code</c> + <c>profit_center</c> + <c>customer_unit_id</c>」の組です。<br/>
        ''' ・<c>customer_code</c> と <c>profit_center</c> は大文字小文字を無視して比較します。<br/>
        ''' ・<c>profit_center</c> と <c>customer_unit_id</c> は、引数が <c>NULL</c> の場合、DB 側の同列が <c>NULL</c> の行のみ一致と判定します。<br/>
        ''' ・パフォーマンスのため、先頭 1 行で打ち切ります（<c>FETCH FIRST 1 ROWS ONLY</c>）。
        ''' </remarks>
        Public Function ExistsCustomerSetting(
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long?,
            Optional excludeCustomerSettingId As Long = 0
        ) As Boolean

            Dim sql As String = "
                SELECT 1
                FROM customer_setting_mst
                WHERE UPPER(customer_code) = UPPER(:p_ccode)
                    AND (UPPER(profit_center) = UPPER(:p_pc)
                        OR (profit_center IS NULL AND :p_pc IS NULL))
                    AND (customer_unit_id = :p_cuid
                        OR (customer_unit_id IS NULL AND :p_cuid IS NULL))
            "

            If excludeCustomerSettingId > 0 Then
                sql &= "   AND customer_setting_id <> :p_exclude "
            End If

            sql &= " FETCH FIRST 1 ROWS ONLY "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True

                    ' 必須：customer_code（空でも比較文字列として渡す）
                    cmd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode

                    ' profit_center（NULL 許容）
                    Dim pcParam As New OracleParameter(":p_pc", OracleDbType.Varchar2)
                    pcParam.Value = If(String.IsNullOrWhiteSpace(profitCenter), CType(DBNull.Value, Object), profitCenter)
                    cmd.Parameters.Add(pcParam)

                    ' customer_unit_id（NULL 許容）
                    Dim cuParam As New OracleParameter(":p_cuid", OracleDbType.Int64)
                    cuParam.Value = If(customerUnitId.HasValue, CType(customerUnitId.Value, Object), DBNull.Value)
                    cmd.Parameters.Add(cuParam)

                    ' 自レコード除外（更新時）
                    If excludeCustomerSettingId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeCustomerSettingId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function
#End Region

#Region "INSERT / Update（通常版）"

        Public Function InsertCustomerSetting(
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long,
            prodMgmtUserId As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO customer_setting_mst (
                    customer_setting_id,
                    customer_code,
                    profit_center,
                    customer_unit_id,
                    prod_mgmt_user_id,
                    active_flag,
                    created_at,
                    created_user_id,
                    created_pg_id,
                    updated_at,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    CUSTOMER_SETTING_SEQ.NEXTVAL,
                    :p_ccode,
                    :p_pc,
                    :p_cuid,
                    :p_puser,
                    :p_active,
                    SYSDATE,
                    :p_user,
                    :p_pg,
                    SYSDATE,
                    :p_user,
                    :p_pg
                )
                RETURNING customer_setting_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode
                        cmd.Parameters.Add(":p_pc", OracleDbType.Varchar2).Value = profitCenter
                        cmd.Parameters.Add(":p_cuid", OracleDbType.Int64).Value = customerUnitId
                        cmd.Parameters.Add(":p_puser", OracleDbType.Varchar2).Value = prodMgmtUserId

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

        ' INSERT

        Public Function InsertCustomerSettingNullable(
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long?,
            prodMgmtUserId As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO customer_setting_mst (
                    customer_code,
                    profit_center,
                    customer_unit_id,
                    prod_mgmt_user_id,
                    active_flag,
                    created_user_id,
                    created_pg_id,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    :p_ccode,
                    :p_pc,
                    :p_cuid,
                    :p_puser,
                    :p_active,
                    :p_user,
                    :p_pg,
                    :p_user,
                    :p_pg
                )
                RETURNING customer_setting_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode

                        Dim pPc As New OracleParameter(":p_pc", OracleDbType.Varchar2)
                        pPc.Value = If(String.IsNullOrWhiteSpace(profitCenter), CType(DBNull.Value, Object), profitCenter)
                        cmd.Parameters.Add(pPc)

                        Dim pCuid As New OracleParameter(":p_cuid", OracleDbType.Int64)
                        pCuid.Value = If(customerUnitId.HasValue, CType(customerUnitId.Value, Object), DBNull.Value)
                        cmd.Parameters.Add(pCuid)

                        cmd.Parameters.Add(":p_puser", OracleDbType.Varchar2).Value = prodMgmtUserId

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

        Public Function UpdateCustomerSetting(
            customerSettingId As Long,
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long,
            prodMgmtUserId As String,
            activeFlag As String,
            loginUserId As String,
            programId As String,
            originalUpdatedAt As DateTime
        ) As Integer

            Dim sql As String = "
                UPDATE customer_setting_mst
                   SET customer_code        = :p_ccode,
                       profit_center        = :p_pc,
                       customer_unit_id     = :p_cuid,
                       prod_mgmt_user_id    = :p_puser,
                       active_flag          = :p_active,
                       updated_at           = SYSDATE,
                       updated_user_id      = :p_user,
                       updated_pg_id        = :p_pg
                 WHERE customer_setting_id  = :p_id
                   AND updated_at           = :p_origupd
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode
                        cmd.Parameters.Add(":p_pc", OracleDbType.Varchar2).Value = profitCenter
                        cmd.Parameters.Add(":p_cuid", OracleDbType.Int64).Value = customerUnitId
                        cmd.Parameters.Add(":p_puser", OracleDbType.Varchar2).Value = prodMgmtUserId

                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId
                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_origupd", OracleDbType.Date).Value = originalUpdatedAt

                        Dim affected = cmd.ExecuteNonQuery()

                        tran.Commit()
                        Return affected
                    End Using
                End Using
            End Using
        End Function

#End Region

#Region "名称→ID解決"

        Public Function GetCustomerUnitIdByName(customerUnitName As String) As Long
            If String.IsNullOrWhiteSpace(customerUnitName) Then
                Throw New ApplicationException("注文工場／担当者名 が空です。")
            End If

            Dim sql As String = "
                SELECT customer_unit_id
                FROM CUSTOMER_UNIT_MST
                WHERE customer_unit_name = :p_name
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = customerUnitName
                    conn.Open()
                    Using rdr = cmd.ExecuteReader()
                        Dim foundId As Long = 0
                        Dim count As Integer = 0
                        While rdr.Read()
                            If rdr.IsDBNull(0) Then Continue While
                            Dim od As OracleDecimal = rdr.GetOracleDecimal(0)
                            foundId = od.ToInt64()
                            count += 1
                            If count > 1 Then
                                Throw New ApplicationException("同名の『注文工場／担当者名』が複数見つかりました。ID を特定できません。")
                            End If
                        End While
                        If count = 0 Then
                            Throw New ApplicationException("指定の『注文工場／担当者名』が見つかりません。")
                        End If
                        Return foundId
                    End Using
                End Using
            End Using
        End Function

        Public Function GetProdMgmtUserIdByName(userName As String) As String
            If String.IsNullOrWhiteSpace(userName) Then
                Throw New ApplicationException("生産管理担当者名 が空です。")
            End If

            Dim sql As String = "
                SELECT fusrid
                FROM USRF
                WHERE fusrname = :p_name
                ORDER BY fusrid ASC
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_name", OracleDbType.Varchar2).Value = userName
                    conn.Open()
                    Using rdr = cmd.ExecuteReader()
                        Dim foundId As String = Nothing
                        Dim count As Integer = 0
                        While rdr.Read()
                            If rdr.IsDBNull(0) Then Continue While
                            foundId = rdr.GetString(0)
                            count += 1
                            If count > 1 Then
                                Throw New ApplicationException("同名の『生産管理担当者名』が複数見つかりました。ID を特定できません。")
                            End If
                        End While
                        If count = 0 Then
                            Throw New ApplicationException("指定の『生産管理担当者名』が見つかりません。")
                        End If
                        Return foundId
                    End Using
                End Using
            End Using
        End Function

#End Region

#Region "更新（排他：updated_at一致時のみ）"

        Public Function UpdateCustomerSettingWithConcurrency(
            customerSettingId As Long,
            customerCode As String,
            profitCenter As String,
            customerUnitId As Long?,
            prodMgmtUserId As String,
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
                        FROM customer_setting_mst
                        WHERE customer_setting_id = :p_id
                    ", conn)
                        cmdChk.Transaction = tran
                        cmdChk.BindByName = True
                        cmdChk.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId

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
                        UPDATE customer_setting_mst
                           SET customer_code        = :p_ccode,
                               profit_center        = :p_pc,
                               customer_unit_id     = :p_cuid,
                               prod_mgmt_user_id    = :p_puser,
                               active_flag          = :p_active,
                               updated_at           = SYSDATE,
                               updated_user_id      = :p_user,
                               updated_pg_id        = :p_pg
                         WHERE customer_setting_id  = :p_id
                           AND updated_at           = :p_currupd
                    "

                    Dim affected As Integer = 0

                    Using cmdUpd As New OracleCommand(sqlUpd, conn)
                        cmdUpd.Transaction = tran
                        cmdUpd.BindByName = True
                        cmdUpd.Parameters.Add(":p_ccode", OracleDbType.Varchar2).Value = customerCode

                        Dim pPc As New OracleParameter(":p_pc", OracleDbType.Varchar2)
                        pPc.Value = If(String.IsNullOrWhiteSpace(profitCenter), DBNull.Value, profitCenter)
                        cmdUpd.Parameters.Add(pPc)

                        Dim pCuid As New OracleParameter(":p_cuid", OracleDbType.Int64)
                        pCuid.Value = If(customerUnitId.HasValue, customerUnitId.Value, DBNull.Value)
                        cmdUpd.Parameters.Add(pCuid)

                        cmdUpd.Parameters.Add(":p_puser", OracleDbType.Varchar2).Value = prodMgmtUserId

                        cmdUpd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmdUpd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmdUpd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId

                        cmdUpd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId
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