Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Text
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    Public Class MappingRepository

#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

#Region "一覧取得"
        ' マッピングマスタ一覧取得
        Public Function GetMappingList(
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
            sb.AppendLine("  profile_id             AS ""ProfileId"",")
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  folder_type            AS ""FolderType"",")
            sb.AppendLine("  version                AS ""Version"",")
            sb.AppendLine("  header_row_index       AS ""HeaderRowIndex"",")
            sb.AppendLine("  data_start_row_index   AS ""DataStartRowIndex"",")
            sb.AppendLine("  default_sheet_name     AS ""DefaultSheetName"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM mapping_profile_list_view ")
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

            sb.AppendLine("ORDER BY customer_code, profit_center, customer_unit_id ")

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

        ' マッピング明細マスタ一覧取得
        Public Function GetFieldMappingList(
            Optional ByVal profileId As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  field_mapping_id       AS ""FieldMappingId"",")
            sb.AppendLine("  profile_id             AS ""ProfileId"",")
            sb.AppendLine("  format_type            AS ""FormatType"",")
            sb.AppendLine("  target_field           AS ""TargetField"",")
            sb.AppendLine("  source_column_index    AS ""SourceColumnIndex"",")
            sb.AppendLine("  source_header_name     AS ""SourceHeaderName"",")
            sb.AppendLine("  source_cell_address    AS ""SourceCellAddress"",")
            sb.AppendLine("  source_sheet_name      AS ""SourceSheetName"",")
            sb.AppendLine("  row_selector_type      AS ""RowSelectorType"",")
            sb.AppendLine("  row_selector_value     AS ""RowSelectorValue"",")
            sb.AppendLine("  data_type              AS ""DataType"",")
            sb.AppendLine("  format_pattern         AS ""FormatPattern"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM field_mapping_list_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()
            Dim pProfileId As String = If(String.IsNullOrWhiteSpace(profileId), Nothing, profileId.Trim())
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

            If pProfileId IsNot Nothing Then
                sb.AppendLine("AND profile_id = :p_profile ")
                prm.Add(New OracleParameter(":p_profile", OracleDbType.Varchar2) With {.Value = pProfileId})
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY format_type, field_mapping_id ")

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
        ' ProfileId (NUMBER(10,0)) で一意に取得
        Public Function GetMappingProfile(profileId As Long) As DataTable
            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  profile_id             AS ""ProfileId"",")
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  folder_type            AS ""FolderType"",")
            sb.AppendLine("  version                AS ""Version"",")
            sb.AppendLine("  header_row_index       AS ""HeaderRowIndex"",")
            sb.AppendLine("  data_start_row_index   AS ""DataStartRowIndex"",")
            sb.AppendLine("  default_sheet_name     AS ""DefaultSheetName"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM mapping_profile_list_view ")
            sb.AppendLine("WHERE profile_id = :p_profile ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(New OracleParameter(":p_profile", OracleDbType.Int64) With {.Value = profileId})
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' 取引先設定IDからマッピングプロファイルを取得（なければ空リスト）
        ''' </summary>

        Public Function GetMappingInfo(customerSettingId As Long, FolderType As Integer) As List(Of MappingInfo)
            Dim list As New List(Of MappingInfo)()

            Using conn As New OracleConnection(_connectionString)
                conn.Open()

                Dim sql As String =
                "Select m.customer_setting_id " &
                ",m.file_id " &
                ",m.folder_type " &
                ",nvl(m.header_row_index,-1) as header_row_index " &
                ",nvl(m.data_start_row_index,-1) as data_start_row_index " &
                ",m.default_sheet_name " &
                ",s2.format_type " &
                ",s2.file_type " &
                ",s2.delimiter " &
                ",s2.enclosure " &
                ",s2.header_flag " &
                ",s2.charset " &
                ",s1.profile_id " &
                ",s1.target_field " &
                ",nvl(s1.source_column_index,-1) as source_column_index " &
                ",s1.source_header_name " &
                ",s1.source_sheet_name " &
                ",s1.source_cell_address " &
                ",s1.row_selector_type " &
                ",s1.row_selector_value " &
                ",s1.data_type " &
                ",s1.format_pattern " &
                "From mapping_profile_mst m " &
                "inner Join field_mapping_mst s1 ON m.profile_id = s1.profile_id " &
                "inner Join file_mst s2 ON m.file_id = s2.file_id and s1.format_type = s2.format_type " &
                "where m.active_flag = 'Y' and s1.active_flag = 'Y' and s2.active_flag = 'Y' " &
                "And m.customer_setting_id = :p_customer_setting_id " &
                "And m.folder_type = :p_folder_type " &
                "order by s1.source_column_index"

                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                    cmd.Parameters.Add(":p_folder_type", OracleDbType.Int64).Value = FolderType

                    Using reader As OracleDataReader = cmd.ExecuteReader()

                        'While reader.Read()
                        '    list.Add(New MappingInfo With {
                        '        .CustomerSettingId = Convert.ToInt64(reader("customer_setting_id")),
                        '        .FileId = Convert.ToInt64(reader("file_id")),
                        '        .FolderType = Convert.ToInt32(reader("folder_type")),
                        '        .HeaderRowIndex = Convert.ToInt32(reader("header_row_index")),
                        '        .DataStartRowIndex = Convert.ToInt32(reader("data_start_row_index")),
                        '        .default_sheet_name = reader("default_sheet_name").ToString,
                        '        .profile_id = Convert.ToInt32(reader("profile_id")),
                        '        .format_type = reader("format_type").ToString,
                        '        .file_type = reader("file_type").ToString,
                        '        .delimiter = reader("delimiter").ToString,
                        '        .enclosure = reader("enclosure").ToString,
                        '        .header_flag = reader("header_flag").ToString,
                        '        .line_ending = reader("line_ending").ToString,
                        '        .charset = reader("charset").ToString,
                        '        .target_field = reader("target_field").ToString,
                        '        .source_column_index = Convert.ToInt32(reader("source_column_index")),
                        '        .source_header_name = reader("source_header_name").ToString,
                        '        .source_sheet_name = reader("source_sheet_name").ToString,
                        '        .source_cell_address = reader("source_cell_address").ToString,
                        '        .row_selector_type = reader("row_selector_type").ToString,
                        '        .row_selector_value = reader("row_selector_value").ToString,
                        '        .data_type = reader("data_type").ToString,
                        '        .format_pattern = reader("format_pattern").ToString,
                        '        .trim_flag = reader("trim_flag").ToString,
                        '        .null_if = reader("null_if").ToString,
                        '        .default_value = reader("default_value").ToString
                        '    })
                        'End While
                        While reader.Read()
                            list.Add(New MappingInfo With {
                                .customerSettingId = Convert.ToInt64(reader("customer_setting_id")),
                                .FileId = Convert.ToInt64(reader("file_id")),
                                .FolderType = Convert.ToInt32(reader("folder_type")),
                                .HeaderRowIndex = Convert.ToInt32(reader("header_row_index")),
                                .DataStartRowIndex = Convert.ToInt32(reader("data_start_row_index")),
                                .default_sheet_name = reader("default_sheet_name").ToString,
                                .profile_id = Convert.ToInt32(reader("profile_id")),
                                .format_type = reader("format_type").ToString,
                                .file_type = reader("file_type").ToString,
                                .delimiter = reader("delimiter").ToString,
                                .enclosure = reader("enclosure").ToString,
                                .header_flag = reader("header_flag").ToString,
                                .charset = reader("charset").ToString,
                                .target_field = reader("target_field").ToString,
                                .source_column_index = Convert.ToInt32(reader("source_column_index")),
                                .source_header_name = reader("source_header_name").ToString,
                                .source_sheet_name = reader("source_sheet_name").ToString,
                                .source_cell_address = reader("source_cell_address").ToString,
                                .row_selector_type = reader("row_selector_type").ToString,
                                .row_selector_value = reader("row_selector_value").ToString,
                                .data_type = reader("data_type").ToString,
                                .format_pattern = reader("format_pattern").ToString
                            })
                        End While
                    End Using
                End Using
            End Using

            Return list
        End Function
#End Region

#Region "重複チェック"

        Public Function ExistsMappingProfile(
            customerSettingId As Long,
            folderType As Integer,
            Optional excludeProfileId As Long = 0
        ) As Boolean

            Dim sql As String = "
                SELECT 1
                FROM mapping_profile_list_view
                WHERE customer_setting_id = :p_custid
                    AND folder_type = :p_type
            "

            If excludeProfileId > 0 Then
                sql &= "    AND profile_id <> :p_exclude "
            End If

            sql &= " FETCH FIRST 1 ROWS ONLY "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                    cmd.Parameters.Add(":p_type", OracleDbType.Int32).Value = folderType

                    ' 自レコード除外（更新時）
                    If excludeProfileId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeProfileId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function

#End Region

#Region "INSERT / UPDATE / DELETE（マッピング詳細）"

        ' MAPPING_PROFILE_MST INSERT: 1件追加し、新しい主キーを返す
        Public Function InsertMappingProfile(
            fileId As String,
            customerSettingId As String,
            folderType As String,
            version As String,
            headerRowIndex As String,
            dataStartRowindex As String,
            defaultSheetName As String,
            userId As String,
            pgId As String
            ) As Long

            Dim sql As String = "INSERT INTO mapping_profile_mst (
                            file_id, 
                            customer_setting_id, 
                            folder_type,
                            version, 
                            header_row_index,
                            data_start_row_index,
                            default_sheet_name,
                            created_user_id, 
                            created_pg_id,
                            updated_user_id, 
                            updated_pg_id
                        ) VALUES (
                            :file_id, 
                            :customer_setting_id, 
                            :folder_type,
                            :version, 
                            :header_row_index, 
                            :data_start_row_index, 
                            :default_sheet_name,
                            :created_user_id, 
                            :created_pg_id,
                            :updated_user_id, 
                            :updated_pg_id
                        )
                        RETURNING profile_id INTO :p_newid
                    "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        ' --- パラメータ ---
                        AddInt32(cmd, ":file_id", Convert.ToInt32(fileId))
                        AddInt32(cmd, ":customer_setting_id", Convert.ToInt32(customerSettingId))
                        AddInt32(cmd, ":folder_type", Convert.ToInt32(folderType))
                        AddInt32OrNull(cmd, ":version", version)
                        AddInt32OrNull(cmd, ":header_row_index", headerRowIndex)
                        AddInt32OrNull(cmd, ":data_start_row_index", dataStartRowindex)
                        AddVarcharOrNull(cmd, ":default_sheet_name", defaultSheetName)
                        AddVarchar(cmd, ":created_user_id", userId)
                        AddVarchar(cmd, ":created_pg_id", pgId)
                        AddVarchar(cmd, ":updated_user_id", userId)
                        AddVarchar(cmd, ":updated_pg_id", pgId)

                        Dim pNewId = New OracleParameter(":p_newid", OracleDbType.Int64)
                        pNewId.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(pNewId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

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

        ' MAPPING_PROFILE_MST UPDATE: 主キーで1件更新
        Public Function UpdateMappingProfile(
            profileId As String,
            version As String,
            headerRowIndex As String,
            dataStartRowindex As String,
            defaultSheetName As String,
            activeFlag As String,
            userId As String,
            pgId As String
            ) As Long

            Dim sql As String = "UPDATE  mapping_profile_mst
                                SET version             =:version, 
                                header_row_index        =:header_row_index,
                                data_start_row_index    =:data_start_row_index,
                                default_sheet_name      =:default_sheet_name,
                                active_flag             =:active_flag,
                                updated_at              = SYSDATE,
                                updated_user_id         =:updated_user_id,
                                updated_pg_id           =:updated_pg_id
                                WHERE profile_id = :profile_id
                                "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        ' --- パラメータ ---
                        AddInt32OrNull(cmd, ":version", version)
                        AddInt32OrNull(cmd, ":header_row_index", headerRowIndex)
                        AddInt32OrNull(cmd, ":data_start_row_index", dataStartRowindex)
                        AddVarcharOrNull(cmd, ":default_sheet_name", defaultSheetName)
                        AddVarchar(cmd, ":active_flag", activeFlag)
                        AddVarchar(cmd, ":updated_user_id", userId)
                        AddVarchar(cmd, ":updated_pg_id", pgId)
                        AddInt32(cmd, ":profile_id", profileId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

                        Return profileId

                    End Using
                End Using
            End Using
        End Function

        ' FIELD_MAPPING_MST INSERT: 1件追加し、新しい主キーを返す
        Public Function InsertFieldMapping(
            profileId As String,
            formatType As String,
            targetField As String,
            sourceColumnIndex As String,
            sourceHeaderName As String,
            sourceSheetName As String,
            sourceCellAddress As String,
            rowSelectorType As String,
            rowSelectorValue As String,
            dataType As String,
            formatPattern As String,
            userId As String,
            pgId As String
        ) As Decimal

            Using cn As New OracleConnection(_connectionString)
                cn.Open()
                Using cmd As New OracleCommand() With {.Connection = cn, .BindByName = True}
                    cmd.CommandText = "
                        INSERT INTO field_mapping_mst (
                            profile_id, 
                            format_type, 
                            target_field,
                            source_column_index, 
                            source_header_name, 
                            source_cell_address, 
                            source_sheet_name,
                            row_selector_type, 
                            row_selector_value, 
                            data_type, 
                            format_pattern,
                            created_user_id, 
                            created_pg_id,
                            updated_user_id, 
                            updated_pg_id
                        ) VALUES (
                            :profile_id, 
                            :format_type, 
                            :target_field,
                            :source_column_index, 
                            :source_header_name, 
                            :source_cell_address, 
                            :source_sheet_name,
                            :row_selector_type, 
                            :row_selector_value, 
                            :data_type, 
                            :format_pattern,
                            :created_user_id, 
                            :created_pg_id,
                            :updated_user_id, 
                            :updated_pg_id
                        )
                        RETURNING field_mapping_id INTO :out_id
                    "

                    ' --- パラメータ ---
                    AddInt32(cmd, ":profile_id", Convert.ToInt32(profileId))
                    AddVarchar(cmd, ":format_type", formatType)
                    AddVarchar(cmd, ":target_field", targetField)
                    AddVarchar(cmd, ":data_type", dataType)
                    AddInt32OrNull(cmd, ":source_column_index", sourceColumnIndex)
                    AddVarcharOrNull(cmd, ":source_header_name", sourceHeaderName)
                    AddVarcharOrNull(cmd, ":source_sheet_name", sourceSheetName)
                    AddVarcharOrNull(cmd, ":source_cell_address", sourceCellAddress)
                    AddVarcharOrNull(cmd, ":row_selector_type", rowSelectorType)
                    AddVarcharOrNull(cmd, ":row_selector_value", rowSelectorValue)
                    AddVarcharOrNull(cmd, ":format_pattern", formatPattern)
                    AddVarchar(cmd, ":created_user_id", userId)
                    AddVarchar(cmd, ":created_pg_id", pgId)
                    AddVarchar(cmd, ":updated_user_id", userId)
                    AddVarchar(cmd, ":updated_pg_id", pgId)

                    Dim pNewId = New OracleParameter(":out_id", OracleDbType.Int64)
                    pNewId.Direction = ParameterDirection.Output
                    cmd.Parameters.Add(pNewId)

                    cmd.ExecuteNonQuery()

                    If TypeOf pNewId.Value Is OracleDecimal Then
                        Dim od As OracleDecimal = DirectCast(pNewId.Value, OracleDecimal)
                        Return od.ToInt64()
                    Else
                        Return Convert.ToInt64(pNewId.Value.ToString())
                    End If
                End Using
            End Using
        End Function

        ' FIELD_MAPPING_MST UPDATE: 主キーで1件更新
        Public Function UpdateFieldMapping(
            fieldMappingId As String,
            formatType As String,
            targetField As String,
            sourceColumnIndex As String,
            sourceHeaderName As String,
            sourceSheetName As String,
            sourceCellAddress As String,
            rowSelectorType As String,
            rowSelectorValue As String,
            dataType As String,
            formatPattern As String,
            activeFlag As String,
            userId As String,
            pgId As String
        ) As Integer

            Dim affected As Integer

            Using cn As New OracleConnection(_connectionString)
                cn.Open()
                Using cmd As New OracleCommand() With {.Connection = cn, .BindByName = True}
                    cmd.CommandText = "
                        UPDATE field_mapping_mst
                           SET format_type         = :format_type,
                               target_field        = :target_field,
                               source_column_index = :source_column_index,
                               source_header_name  = :source_header_name,
                               source_cell_address = :source_cell_address,
                               source_sheet_name   = :source_sheet_name,
                               row_selector_type   = :row_selector_type,
                               row_selector_value  = :row_selector_value,
                               data_type           = :data_type,
                               format_pattern      = :format_pattern,
                               active_flag         = :active_flag,
                               updated_at          = SYSDATE,
                               updated_user_id     = :updated_user_id,
                               updated_pg_id       = :updated_pg_id
                         WHERE field_mapping_id   = :field_mapping_id
                    "

                    ' --- パラメータ ---
                    AddVarchar(cmd, ":format_type", formatType)
                    AddVarchar(cmd, ":target_field", targetField)
                    AddIntOrNull(cmd, ":source_column_index", sourceColumnIndex)
                    AddVarcharOrNull(cmd, ":source_header_name", sourceHeaderName)
                    AddVarcharOrNull(cmd, ":source_cell_address", sourceCellAddress)
                    AddVarcharOrNull(cmd, ":source_sheet_name", sourceSheetName)
                    AddVarcharOrNull(cmd, ":row_selector_type", rowSelectorType)
                    AddVarcharOrNull(cmd, ":row_selector_value", rowSelectorValue)
                    AddVarchar(cmd, ":data_type", dataType)
                    AddVarcharOrNull(cmd, ":format_pattern", formatPattern)
                    AddChar(cmd, ":active_flag", activeFlag)
                    AddVarchar(cmd, ":updated_user_id", userId)
                    AddVarchar(cmd, ":updated_pg_id", pgId)
                    AddDecimal(cmd, ":field_mapping_id", fieldMappingId)

                    affected = cmd.ExecuteNonQuery()
                End Using
            End Using

            Return affected
        End Function

        ' FIELD_MAPPING_MST DELETE: 主キーで1件削除
        Public Function DeleteFieldMapping(fieldMappingId As String) As Integer
            Dim affected As Integer

            Using cn As New OracleConnection(_connectionString)
                cn.Open()
                Using cmd As New OracleCommand() With {.Connection = cn, .BindByName = True}
                    cmd.CommandText = "
                        DELETE FROM field_mapping_mst
                        WHERE field_mapping_id = :field_mapping_id
                    "
                    AddDecimal(cmd, ":field_mapping_id", fieldMappingId)
                    affected = cmd.ExecuteNonQuery()
                End Using
            End Using

            Return affected
        End Function
#End Region

    End Class
End Namespace
