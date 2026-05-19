Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    Public Class FileRepository

#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

#End Region

#Region "一覧取得"
        ' ファイルマスタ一覧取得
        Public Function GetFileList(
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
            sb.AppendLine("  file_id                AS ""FileId"",")
            sb.AppendLine("  customer_setting_id    AS ""CustomerSettingId"",")
            sb.AppendLine("  customer_code          AS ""CustomerCode"",")
            sb.AppendLine("  customer_name          AS ""CustomerName"",")
            sb.AppendLine("  profit_center          AS ""ProfitCenter"",")
            sb.AppendLine("  customer_unit_id       AS ""CustomerUnitId"",")
            sb.AppendLine("  customer_unit_name     AS ""CustomerUnitName"",")
            sb.AppendLine("  prod_mgmt_user_id      AS ""ProdMgmtUserId"",")
            sb.AppendLine("  folder_type            AS ""FolderType"",")
            sb.AppendLine("  format_type            AS ""FormatType"",")
            sb.AppendLine("  file_type              AS ""FileType"",")
            sb.AppendLine("  delimiter              AS ""Delimiter"",")
            sb.AppendLine("  enclosure              AS ""Enclosure"",")
            sb.AppendLine("  header_flag            AS ""HeaderFlag"",")
            sb.AppendLine("  charset                AS ""Charset"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId""")
            sb.AppendLine("FROM file_list_view ")
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

            sb.AppendLine("ORDER BY customer_code, profit_center, customer_unit_id, folder_type")

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

        Public Function GetFile(fileId As Long) As DataTable
            Dim dt As New DataTable()

            Dim sql As String = "
                SELECT 
                    file_id             AS ""FileId"",
                    customer_setting_id AS ""CustomerSettingId"",
                    customer_code       AS ""CustomerCode"",
                    customer_name       AS ""CustomerName"",
                    profit_center       AS ""ProfitCenter"",
                    customer_unit_id    AS ""CustomerUnitId"",
                    customer_unit_name  AS ""CustomerUnitName"",
                    folder_type         AS ""FolderType"",
                    format_type         AS ""FormatType"",
                    file_type           AS ""FileType"",
                    delimiter           AS ""Delimiter"",
                    enclosure           AS ""Enclosure"",
                    header_flag         AS ""HeaderFlag"",
                    active_flag         AS ""ActiveFlag"",
                    created_at          AS ""CreatedAt"",
                    created_user_id     AS ""CreatedUserId"",
                    created_pg_id       AS ""CreatedPgId"",
                    updated_at          AS ""UpdatedAt"",
                    updated_user_id     As ""UpdatedUserId"",
                    updated_pg_id       AS ""UpdatedPgId""
                FROM file_list_view
                WHERE file_id = :p_id
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = fileId
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

        'ファイルID取得
        Public Function FindFileIdSimple(
            CustomerSettingId As Long,
            FolderType As Long?
        ) As Long?

            Dim sql As String = "
                SELECT file_id
                FROM file_mst
                WHERE customer_setting_id = :p_custid
                    AND folder_type = :p_type
            "
            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = CustomerSettingId
                    cmd.Parameters.Add(":p_type", OracleDbType.Int32).Value = FolderType

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then Return Nothing
                    Return Convert.ToInt64(obj)
                End Using
            End Using
        End Function

#End Region

#Region "重複チェック"

        Public Function ExistsFile(
            customerSettingId As Long,
            folderType As Integer,
            Optional excludeFileId As Long = 0
        ) As Boolean

            Dim sql As String = "
                SELECT 1
                FROM file_list_view
                WHERE customer_setting_id = :p_custid
                    AND folder_type = :p_folder
            "

            If excludeFileId > 0 Then
                sql &= "    AND file_id <> :p_exclude "
            End If

            sql &= " FETCH FIRST 1 ROWS ONLY "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                    cmd.Parameters.Add(":p_folder", OracleDbType.Int32).Value = folderType

                    ' 自レコード除外（更新時）
                    If excludeFileId > 0 Then
                        cmd.Parameters.Add(":p_exclude", OracleDbType.Int64).Value = excludeFileId
                    End If

                    conn.Open()
                    Dim obj = cmd.ExecuteScalar()
                    Return (obj IsNot Nothing AndAlso obj IsNot DBNull.Value)
                End Using
            End Using
        End Function
#End Region

#Region "INSERT / Update（通常版）"

        Public Function InsertFIle(
            fileId As Long,
            customerSettingId As Long,
            folderType As Integer,
            formatType As String,
            fileType As String,
            delimiter As String,
            enclosure As String,
            headerFlag As String,
            charset As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
            INSERT INTO file_mst (
                file_id,
                customer_setting_id,
                folder_type,
                format_type,
                file_type,
                delimiter,
                enclosure,
                header_flag,
                charset,
                active_flag,
                created_at,
                created_user_id,
                created_pg_id,
                updated_at,
                updated_user_id,
                updated_pg_id
            ) VALUES (
                FILE_SEQ.NEXTVAL,
                :p_custid,
                :p_foldertype,
                :p_format,
                :p_filetype,
                :p_delimiter,
                :p_enclosure,
                :p_header,
                :p_charset,
                :p_active,
                SYSDATE,
                :p_user,
                :p_pg,
                SYSDATE,
                :p_user,
                :p_pg
            )
            RETURNING file_id INTO :p_newid
        "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_foldertype", OracleDbType.Int32).Value = folderType
                        cmd.Parameters.Add(":p_format", OracleDbType.Varchar2).Value = formatType
                        cmd.Parameters.Add(":p_filetype", OracleDbType.Varchar2).Value = fileType
                        cmd.Parameters.Add(":p_delimiter", OracleDbType.Varchar2).Value = delimiter
                        cmd.Parameters.Add(":p_enclosure", OracleDbType.Varchar2).Value = enclosure
                        cmd.Parameters.Add(":p_header", OracleDbType.Varchar2).Value = headerFlag
                        cmd.Parameters.Add(":p_charset", OracleDbType.Varchar2).Value = charset

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

        Public Function InsertFileNullable(
            customerSettingId As Long,
            folderType As Integer,
            formatType As String,
            fileType As String,
            delimiter As String,
            enclosure As String,
            headerFlag As String,
            charset As String,
            activeFlag As String,
            loginUserId As String,
            programId As String
        ) As Long

            Dim sql As String = "
                INSERT INTO file_mst (
                    customer_setting_id,
                    folder_type,
                    format_type,
                    file_type,
                    delimiter,
                    enclosure,
                    header_flag,
                    charset,
                    active_flag,
                    created_user_id,
                    created_pg_id,
                    updated_user_id,
                    updated_pg_id
                ) VALUES (
                    :p_custid,
                    :p_foldertype,
                    :p_format,
                    :p_filetype,
                    :p_delimiter,
                    :p_enclosure,
                    :p_header,
                    :p_charset,
                    :p_active,
                    :p_user,
                    :p_pg,
                    :p_user,
                    :p_pg
                )
                RETURNING file_id INTO :p_newid
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True

                        cmd.Parameters.Add(":p_custid", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_foldertype", OracleDbType.Int32).Value = folderType
                        cmd.Parameters.Add(":p_format", OracleDbType.Varchar2).Value = formatType
                        cmd.Parameters.Add(":p_filetype", OracleDbType.Varchar2).Value = fileType
                        cmd.Parameters.Add(":p_delimiter", OracleDbType.Varchar2).Value = delimiter
                        cmd.Parameters.Add(":p_enclosure", OracleDbType.Varchar2).Value = enclosure
                        cmd.Parameters.Add(":p_header", OracleDbType.Varchar2).Value = headerFlag
                        cmd.Parameters.Add(":p_charset", OracleDbType.Varchar2).Value = charset
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

        Public Function UpdateFile(
            fileId As Long,
            formatType As String,
            fileType As String,
            delimiter As String,
            enclosure As String,
            headerFlag As String,
            charset As String,
            activeFlag As String,
            loginUserId As String,
            programId As String,
            originalUpdatedAt As DateTime
        ) As Integer

            Dim sql As String = "
                UPDATE file_mst
                   SET format_type          = :p_format,
                       file_type            = :p_filetype,
                       delimiter            = :p_delimiter,
                       enclosure            = :p_enclosure,
                       header_flag          = :p_header,
                       charset              = :p_charset,
                       active_flag          = :p_active,
                       updated_at           = SYSDATE,
                       updated_user_id      = :p_user,
                       updated_pg_id        = :p_pg
                 WHERE file_id = :p_id
                   AND updated_at = :p_origupd
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.Parameters.Add(":p_format", OracleDbType.Varchar2).Value = formatType
                        cmd.Parameters.Add(":p_filetype", OracleDbType.Varchar2).Value = fileType
                        cmd.Parameters.Add(":p_delimiter", OracleDbType.Varchar2).Value = delimiter
                        cmd.Parameters.Add(":p_enclosure", OracleDbType.Varchar2).Value = enclosure
                        cmd.Parameters.Add(":p_header", OracleDbType.Varchar2).Value = headerFlag
                        cmd.Parameters.Add(":p_charset", OracleDbType.Varchar2).Value = charset
                        cmd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId
                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = fileId
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
        Public Function UpdateFileWithConcurrency(
            fileId As Long,
            formatType As String,
            fileType As String,
            delimiter As String,
            enclosure As String,
            headerFlag As String,
            charset As String,
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
                        FROM file_mst
                        WHERE file_id = :p_id
                    ", conn)
                        cmdChk.Transaction = tran
                        cmdChk.BindByName = True
                        cmdChk.Parameters.Add(":p_id", OracleDbType.Int64).Value = fileId

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
                        UPDATE file_mst
                        SET format_type         = :p_format,
                            file_type           = :p_filetype,
                            delimiter           = :p_delimiter,
                            enclosure           = :p_enclosure,
                            header_flag         = :p_header,
                            charset             = :p_charset,
                            active_flag         = :p_active,
                            updated_at          = SYSDATE,
                            updated_user_id     = :p_user,
                            updated_pg_id       = :p_pg
                        WHERE file_id           = :p_id
                            AND updated_at      = :p_currupd
                    "

                    Dim affected As Integer = 0

                    Using cmdUpd As New OracleCommand(sqlUpd, conn)
                        cmdUpd.Transaction = tran
                        cmdUpd.BindByName = True
                        cmdUpd.Parameters.Add(":p_format", OracleDbType.Varchar2).Value = formatType
                        cmdUpd.Parameters.Add(":p_filetype", OracleDbType.Varchar2).Value = fileType
                        cmdUpd.Parameters.Add(":p_delimiter", OracleDbType.Varchar2).Value = delimiter
                        cmdUpd.Parameters.Add(":p_enclosure", OracleDbType.Varchar2).Value = enclosure
                        cmdUpd.Parameters.Add(":p_header", OracleDbType.Varchar2).Value = headerFlag
                        cmdUpd.Parameters.Add(":p_charset", OracleDbType.Varchar2).Value = charset

                        cmdUpd.Parameters.Add(":p_active", OracleDbType.Char).Value = activeFlag
                        cmdUpd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = loginUserId
                        cmdUpd.Parameters.Add(":p_pg", OracleDbType.Varchar2).Value = programId
                        cmdUpd.Parameters.Add(":p_id", OracleDbType.Int64).Value = fileId
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
