
Imports Oracle.ManagedDataAccess.Client
Imports System.Data
Imports System.Text
Imports OMS.Common
Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace OMS.Data
    Public Class ImpFilesStageRepository

#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

#Region "一覧取得"
        ''' <summary>
        ''' filename で 取込ファイルワークテーブルを取得する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="fileName"></param>
        ''' <param name="folderPath"></param>
        ''' <returns></returns>
        Public Function GetImpFilesStageFilename(conn As OracleConnection, tran As OracleTransaction, Optional ByVal fileName As String = Nothing, Optional ByVal folderPath As String = Nothing) As ImpFilesStageRow

            Dim records = GetImpFilesStage(conn, tran, fileName:=fileName, folderPath:=folderPath)
            Dim dataRows = ToClass(records)

            If (dataRows.Count = 0) Then
                Return Nothing
            End If
            Return dataRows(0)
        End Function
        ''' <summary>
        ''' 一時ファイルテーブル一覧取得 生産計画用
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="impFileStageId"></param>
        ''' <param name="status"></param>
        ''' <param name="fileName"></param>
        ''' <param name="folderPath"></param>
        ''' <returns></returns>
        Public Function GetImpFilesStage(conn As OracleConnection, tran As OracleTransaction,
            Optional ByVal impFileStageId As Long? = Nothing,
            Optional ByVal status As String = Nothing,
            Optional ByVal fileName As String = Nothing,
            Optional ByVal folderPath As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT * ")
            sb.AppendLine("FROM imp_files_stage ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pStatus As String = Utils.BuildLikePattern(status, LikeMode.Contains)
            Dim pFilename As String = Utils.BuildLikePattern(fileName, LikeMode.Contains)
            Dim pFolderPath As String = Utils.BuildLikePattern(folderPath, LikeMode.Contains)
            Dim pImpFileStageId = impFileStageId

            If pStatus IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = pStatus})
            End If

            If pFilename IsNot Nothing Then
                sb.AppendLine("AND UPPER(file_name) LIKE UPPER(:p_filename) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_filename", OracleDbType.Varchar2) With {.Value = pFilename})
            End If

            If pFolderPath IsNot Nothing Then
                sb.AppendLine("AND UPPER(folder_path) LIKE UPPER(:p_folderPath) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_folderPath", OracleDbType.Varchar2) With {.Value = pFolderPath})
            End If

            If impFileStageId IsNot Nothing Then
                sb.AppendLine("AND imp_file_stage_id = :p_impFileStageId ")
                prm.Add(New OracleParameter(":p_impFileStageId", OracleDbType.Int64) With {.Value = pImpFileStageId})
            End If

            Using cmd As New OracleCommand(sb.ToString(), conn)
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
            Return dt

        End Function

        ' 一時ファイルテーブル一覧取得
        Public Function GetImpFilesStage(
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
            sb.AppendLine("  imp_file_stage_id     AS ""ImpFileStageId"", ")
            sb.AppendLine("  customer_setting_id   AS ""CustomerSettingId"", ")
            sb.AppendLine("  customer_code         AS ""CustomerCode"", ")
            sb.AppendLine("  customer_name         AS ""CustomerName"", ")
            sb.AppendLine("  profit_center         AS ""ProfitCenter"", ")
            sb.AppendLine("  customer_unit_id      AS ""CustomerUnitId"", ")
            sb.AppendLine("  customer_unit_name    AS ""CustomerUnitName"", ")
            sb.AppendLine("  folder_type           AS ""FolderType"", ")
            sb.AppendLine("  folder_path           AS ""FolderPath"", ")
            sb.AppendLine("  file_name             AS ""FileName"", ")
            sb.AppendLine("  staged_folder_path    AS ""StagedFolderPath"", ")
            sb.AppendLine("  staged_file_name      AS ""StagedFileName"", ")
            sb.AppendLine("  reconcile_flag        AS ""ReconcileFlag"", ")
            sb.AppendLine("  fcst_reconcile_flag   AS ""FcstReconcileFlag"", ")
            sb.AppendLine("  hand_flag             AS ""HandFlag"", ")
            sb.AppendLine("  created_at            AS ""CreatedAt"", ")
            sb.AppendLine("  created_user_id       AS ""CreatedUserId"", ")
            sb.AppendLine("  created_pg_id         AS ""CreatedPgId"", ")
            sb.AppendLine("  updated_at            AS ""UpdatedAt"", ")
            sb.AppendLine("  updated_user_id       AS ""UpdatedUserId"", ")
            sb.AppendLine("  updated_pg_id         AS ""UpdatedPgId"",")
            sb.AppendLine("  prod_mgmt_user_id     AS ""ProdMgmtUserId"" ")
            sb.AppendLine("FROM imp_files_stage_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pCustomerCode As String = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName As String = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter As String = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName As String = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId As String = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)

            If pCustomerCode IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_code) LIKE UPPER(:p_ccode) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_ccode", OracleDbType.Varchar2) With {.Value = pCustomerCode})
            End If

            If pCustomerName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_name) LIKE UPPER(:p_cname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cname", OracleDbType.Varchar2) With {.Value = pCustomerCode})
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

            'sb.AppendLine("ORDER BY created_at, customer_setting_id ")
            sb.AppendLine("ORDER BY customer_setting_id, folder_type, created_at")

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

#Region "INSERT"

        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As ImpFilesStageRow) As String

            Dim records = New List(Of ImpFilesStageRow)()
            records.Add(row)
            Return InsertRange(conn, tran, records)

        End Function

        ''' <summary>
        ''' OrderRow class リストをDBに追加する (元のコード同等呼び出し
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of ImpFilesStageRow))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, records.ToList())
                    tran.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' InsertRange
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns></returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of ImpFilesStageRow)) As String

            Dim errorMessage As String = ""
            If records Is Nothing Then Return "ImpFilesStageRow: no record data. "
            Try

                'conn.Open()
                Const sql As String =
                "INSERT INTO imp_files_stage (" &
                "  customer_setting_id, folder_type, folder_path, file_name, staged_folder_path, staged_file_name, " &
                "  reconcile_flag, fcst_reconcile_flag, hand_flag, status, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id" &
                ") VALUES (" &
                "  :p_csid, :p_ftype, :p_folder, :p_fname, :p_sfolder, :p_sfname, " &
                "  :p_recon, :p_fcst_recon, :p_hand, :p_status, " &
                "  :p_c_at, :p_c_uid, :p_c_pid, " &
                "  :p_u_at, :p_u_uid, :p_u_pid" &
                ")"

                Using cmd As New OracleCommand(sql, conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    For Each r In records
                        cmd.Parameters.Clear()

                        cmd.Parameters.Add(":p_csid", OracleDbType.Int32).Value = r.CustomerSettingId
                        cmd.Parameters.Add(":p_ftype", OracleDbType.Int32).Value = r.FolderType

                        ' 長さ対策（最大1000/255）
                        cmd.Parameters.Add(":p_folder", OracleDbType.Varchar2, 1000).Value =
                                SafeVarchar(r.FolderPath, 1000)
                        cmd.Parameters.Add(":p_fname", OracleDbType.Varchar2, 255).Value =
                                SafeVarchar(r.FileName, 255)
                        cmd.Parameters.Add(":p_sfolder", OracleDbType.Varchar2, 1000).Value =
                                SafeVarchar(r.StagedFolderPath, 1000)
                        cmd.Parameters.Add(":p_sfname", OracleDbType.Varchar2, 255).Value =
                                SafeVarchar(r.StagedFileName, 255)

                        ' フラグ（Y/N 文字）
                        cmd.Parameters.Add(":p_recon", OracleDbType.Char, 1).Value = NormalizeYN(r.ReconcileFlag)
                        cmd.Parameters.Add(":p_fcst_recon", OracleDbType.Char, 1).Value = NormalizeYN(r.FcstReconcileFlag)
                        cmd.Parameters.Add(":p_hand", OracleDbType.Char, 1).Value = NormalizeYN(r.HandFlag)

                        ' ステータス
                        cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = r.Status

                        ' 監査系
                        cmd.Parameters.Add(":p_c_at", OracleDbType.Date).Value = r.CreatedAt
                        cmd.Parameters.Add(":p_c_uid", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.CreatedUserId, 9)
                        cmd.Parameters.Add(":p_c_pid", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.CreatedPgId, 150)
                        cmd.Parameters.Add(":p_u_at", OracleDbType.Date).Value = r.UpdatedAt
                        cmd.Parameters.Add(":p_u_uid", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.UpdatedUserId, 9)
                        cmd.Parameters.Add(":p_u_pid", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.UpdatedPgId, 150)

                        cmd.ExecuteNonQuery()
                    Next

                    'tran.Commit()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage
        End Function




        'Public Sub InsertRange(records As IEnumerable(Of ImpFilesStageRow))
        '    If records Is Nothing Then Return

        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            Const sql As String =
        '                "INSERT INTO imp_files_stage (" &
        '                "  customer_setting_id, folder_type, folder_path, file_name, staged_folder_path, staged_file_name, " &
        '                "  reconcile_flag, fcst_reconcile_flag, hand_flag, status, " &
        '                "  created_at, created_user_id, created_pg_id, " &
        '                "  updated_at, updated_user_id, updated_pg_id" &
        '                ") VALUES (" &
        '                "  :p_csid, :p_ftype, :p_folder, :p_fname, :p_sfolder, :p_sfname, " &
        '                "  :p_recon, :p_fcst_recon, :p_hand, :p_status, " &
        '                "  :p_c_at, :p_c_uid, :p_c_pid, " &
        '                "  :p_u_at, :p_u_uid, :p_u_pid" &
        '                ")"

        '            Using cmd As New OracleCommand(sql, conn)
        '                cmd.Transaction = tran
        '                cmd.BindByName = True
        '                cmd.CommandType = CommandType.Text

        '                For Each r In records
        '                    cmd.Parameters.Clear()

        '                    cmd.Parameters.Add(":p_csid", OracleDbType.Int32).Value = r.CustomerSettingId
        '                    cmd.Parameters.Add(":p_ftype", OracleDbType.Int32).Value = r.FolderType

        '                    ' 長さ対策（最大1000/255）
        '                    cmd.Parameters.Add(":p_folder", OracleDbType.Varchar2, 1000).Value =
        '                        SafeVarchar(r.FolderPath, 1000)
        '                    cmd.Parameters.Add(":p_fname", OracleDbType.Varchar2, 255).Value =
        '                        SafeVarchar(r.FileName, 255)
        '                    cmd.Parameters.Add(":p_sfolder", OracleDbType.Varchar2, 1000).Value =
        '                        SafeVarchar(r.StagedFolderPath, 1000)
        '                    cmd.Parameters.Add(":p_sfname", OracleDbType.Varchar2, 255).Value =
        '                        SafeVarchar(r.StagedFileName, 255)

        '                    ' フラグ（Y/N 文字）
        '                    cmd.Parameters.Add(":p_recon", OracleDbType.Char, 1).Value = NormalizeYN(r.ReconcileFlag)
        '                    cmd.Parameters.Add(":p_fcst_recon", OracleDbType.Char, 1).Value = NormalizeYN(r.FcstReconcileFlag)
        '                    cmd.Parameters.Add(":p_hand", OracleDbType.Char, 1).Value = NormalizeYN(r.HandFlag)

        '                    ' ステータス
        '                    cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = r.Status

        '                    ' 監査系
        '                    cmd.Parameters.Add(":p_c_at", OracleDbType.Date).Value = r.CreatedAt
        '                    cmd.Parameters.Add(":p_c_uid", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.CreatedUserId, 9)
        '                    cmd.Parameters.Add(":p_c_pid", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.CreatedPgId, 150)
        '                    cmd.Parameters.Add(":p_u_at", OracleDbType.Date).Value = r.UpdatedAt
        '                    cmd.Parameters.Add(":p_u_uid", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.UpdatedUserId, 9)
        '                    cmd.Parameters.Add(":p_u_pid", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.UpdatedPgId, 150)

        '                    cmd.ExecuteNonQuery()
        '                Next

        '                tran.Commit()
        '            End Using
        '        End Using
        '    End Using
        'End Sub
#End Region

#Region "UPDATE"
        Public Function Update(conn As OracleConnection, tran As OracleTransaction,
                                Optional ByVal kImpFileStageId As Long? = Nothing,
                                Optional ByVal kFolderPath As String = Nothing,
                                Optional ByVal kFileName As String = Nothing,
                                Optional ByVal stagedFolderPath As String = Nothing,
                                Optional ByVal stagedFileName As String = Nothing,
                                Optional ByVal updatedAt As Date? = Nothing,
                                Optional ByVal status As String = Nothing) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                Dim prm As New List(Of OracleParameter)()
                sb.AppendLine("UPDATE imp_files_stage ")

                ' セット 
                sb.AppendLine("SET ")
                If stagedFolderPath IsNot Nothing Then
                    sb.AppendLine(" staged_folder_path = :p_stagedFolderPath, ")
                    prm.Add(New OracleParameter(":p_stagedFolderPath", OracleDbType.Varchar2) With {.Value = stagedFolderPath})
                End If
                If stagedFileName IsNot Nothing Then
                    sb.AppendLine(" staged_file_name = :p_stagedFileName, ")
                    prm.Add(New OracleParameter(":p_stagedFileName", OracleDbType.Varchar2) With {.Value = stagedFileName})
                End If
                If updatedAt IsNot Nothing Then
                    sb.AppendLine(" updated_at = :p_updatedAt, ")
                    prm.Add(New OracleParameter(":p_updatedAt", OracleDbType.Date) With {.Value = updatedAt})
                End If
                If status IsNot Nothing Then
                    sb.AppendLine(" status = :p_status, ")
                    prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
                End If

                ' 最後のカンマ削除
                If sb.Length > 0 Then
                    Dim startPos As Integer = Math.Max(0, sb.Length - 5)
                    sb.Replace(",", " ", startPos, sb.Length - startPos)
                End If

                sb.AppendLine(" WHERE 1=1 ")

                ' 絞り込み
                If kImpFileStageId IsNot Nothing Then
                    sb.AppendLine("AND imp_file_stage_id = :p_kImpFileStageId ")
                    prm.Add(New OracleParameter(":p_kImpFileStageId", OracleDbType.Int64) With {.Value = kImpFileStageId})
                End If

                If kFolderPath IsNot Nothing Then
                    sb.AppendLine("AND UPPER(folder_path) = :p_kFolderPath ")
                    prm.Add(New OracleParameter(":pkFolderPath", OracleDbType.Varchar2) With {.Value = kFolderPath})
                End If

                If kFileName IsNot Nothing Then
                    sb.AppendLine("AND UPPER(file_name) = :p_kFileName ")
                    prm.Add(New OracleParameter(":p_kFileName", OracleDbType.Varchar2) With {.Value = kFileName})
                End If

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    cmd.ExecuteNonQuery()
                End Using

            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function

#End Region

#Region "DELETE"
        ''' <summary>
        ''' customerSettingId と status で OrderStage レコードを削除する
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="impFileStageId"></param>
        ''' <param name="status"></param>
        ''' <returns></returns>
        Public Function Delete(conn As OracleConnection, tran As OracleTransaction,
                                Optional ByVal impFileStageId As Long? = Nothing,
                                Optional ByVal status As String = Nothing) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                sb.AppendLine("DELETE ")
                sb.AppendLine("FROM imp_files_stage ")
                sb.AppendLine("WHERE 1=1 ")
                Dim prm As New List(Of OracleParameter)()

                If impFileStageId IsNot Nothing Then
                    sb.AppendLine("AND imp_file_stage_id = :p_impFileStageId ")
                    prm.Add(New OracleParameter(":p_impFileStageId", OracleDbType.Int64) With {.Value = impFileStageId})
                End If
                If status IsNot Nothing Then
                    sb.AppendLine("AND UPPER(status) = :p_status ")
                    prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
                End If

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    cmd.ExecuteNonQuery()
                End Using

            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function
#End Region

        ''' <summary>
        ''' DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of ImpFilesStageRow)
            Dim osrs = New List(Of ImpFilesStageRow)()
            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next
            Return osrs
        End Function

        ''' <summary>
        ''' DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As ImpFilesStageRow
            If dt Is Nothing Then
                Return Nothing
            End If

            Dim osr = New ImpFilesStageRow
            osr.ImpFileStageId = dt.Field(Of Long)("imp_file_stage_id")
            osr.CustomerSettingId = dt.Field(Of Long?)("customer_setting_id")
            osr.FolderType = dt.Field(Of Int16?)("folder_type")
            osr.FolderPath = dt.Field(Of String)("folder_path")
            osr.FileName = dt.Field(Of String)("file_name")
            osr.StagedFolderPath = dt.Field(Of String)("staged_folder_path")
            osr.StagedFileName = dt.Field(Of String)("staged_file_name")
            osr.ReconcileFlag = dt.Field(Of String)("reconcile_flag")
            osr.FcstReconcileFlag = dt.Field(Of String)("fcst_reconcile_flag")
            osr.HandFlag = dt.Field(Of String)("hand_flag")
            osr.Status = dt.Field(Of String)("status")
            osr.CreatedAt = dt.Field(Of Date?)("created_at")
            osr.CreatedUserId = dt.Field(Of String)("created_user_id")
            osr.CreatedPgId = dt.Field(Of String)("created_pg_id")
            osr.UpdatedAt = dt.Field(Of Date?)("updated_at")
            osr.UpdatedUserId = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPgId = dt.Field(Of String)("updated_pg_id")
            Return osr
        End Function

        '2026/03/26 酒井 st
        Public Sub UpdateRange(ByVal tran As OracleTransaction, records As IEnumerable(Of ImpFilesStageRow))
            If records Is Nothing Then Return

            Const sql As String =
                        "  UPDATE imp_files_stage SET " &
                        "  hand_flag = :p_hand_flag, " &
                        "  status = :p_status, " &
                        "  updated_at = :p_updated_at, " &
                        "  updated_user_id = :p_updated_user_id, " &
                        "  updated_pg_id = :p_updated_pg_id " &
                        "  WHERE imp_file_stage_id = :p_imp_file_stage_id"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                For Each r In records
                    cmd.Parameters.Clear()

                    cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int32).Value = r.ImpFileStageId
                    cmd.Parameters.Add(":p_hand_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.HandFlag)

                    ' ステータス
                    cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.Status, 20)

                    ' 監査系
                    cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = r.UpdatedAt
                    cmd.Parameters.Add(":p_updated_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.UpdatedUserId, 9)
                    cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.UpdatedPgId, 150)

                    cmd.ExecuteNonQuery()
                Next

            End Using

        End Sub

        ''' <summary>
        ''' 取引先設定IDからフォルダパス＋ワークファイル名を取得（なければ空リスト）
        ''' </summary>
        Public Function GetFolderInfosByImpFileStageId(ByVal impFileStageId As Long) As List(Of FolderPathInfo)
            Dim list As New List(Of FolderPathInfo)()

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Dim sql As String =
                        "SELECT DISTINCT folder_path, staged_folder_path,staged_file_name " &
                        "FROM imp_files_stage " &
                        "WHERE imp_file_stage_id = :p_imp_file_stage_id"

                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId

                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        Dim ordPath = reader.GetOrdinal("folder_path")
                        Dim ordStagedPath = reader.GetOrdinal("staged_folder_path")
                        Dim ordStagedName = reader.GetOrdinal("staged_file_name")
                        While reader.Read()
                            Dim path As String =
                                    If(reader.IsDBNull(ordPath), Nothing, reader.GetString(ordPath)?.Trim())
                            Dim stagedpath As String =
                                    If(reader.IsDBNull(ordStagedPath), Nothing, reader.GetString(ordStagedPath)?.Trim())
                            Dim stagedName As String =
                                    If(reader.IsDBNull(ordStagedName), Nothing, reader.GetString(ordStagedName)?.Trim())


                            If Not String.IsNullOrEmpty(path) Then
                                list.Add(New FolderPathInfo With {
                                        .FolderPath = path,
                                        .Staged_FolderPath = stagedpath,
                                        .Staged_FileName = stagedName
                                    })
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return list
        End Function

        ''' <summary>
        ''' 取込ファイルワークテーブルを削除する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="impFileStageId">一時取込ファイルID</param>
        Public Sub DeleteImpFileStageRange(ByVal tran As OracleTransaction, ByVal impFileStageId As Long)



            Const sql As String = "
                DELETE FROM imp_files_stage 
                WHERE imp_file_stage_id = :p_imp_file_stage_id"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId
                cmd.ExecuteNonQuery()
            End Using

        End Sub

        ''' <summary>
        ''' 取引先設定IDからワークフォルダパス、ワークファイル名、受注区分を取得（なければ空リスト）
        ''' </summary>
        Public Function GetStageFolderInfosByImpFileStageId(ByVal impFileStageId As Long) As List(Of FolderPathInfo)
            Dim list As New List(Of FolderPathInfo)()

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Dim sql As String =
                        "SELECT DISTINCT staged_folder_path, staged_file_name, folder_type " &
                        "FROM imp_files_stage " &
                        "WHERE imp_file_stage_id = :p_imp_file_stage_id"

                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId

                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        Dim ordStagedPath = reader.GetOrdinal("staged_folder_path")
                        Dim ordStagedName = reader.GetOrdinal("staged_file_name")
                        Dim folderType = reader.GetOrdinal("folder_type")
                        While reader.Read()
                            Dim stagedpath As String =
                                    If(reader.IsDBNull(ordStagedPath), Nothing, reader.GetString(ordStagedPath)?.Trim())
                            Dim stagedName As String =
                                    If(reader.IsDBNull(ordStagedName), Nothing, reader.GetString(ordStagedName)?.Trim())
                            Dim ftype As Integer =
                                    If(reader.IsDBNull(folderType), 0, Convert.ToInt32(reader.GetValue(folderType)))

                            If Not String.IsNullOrEmpty(stagedpath) Then
                                list.Add(New FolderPathInfo With {
                                        .Staged_FolderPath = stagedpath,
                                        .Staged_FileName = stagedName,
                                        .folderType = ftype
                                    })
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return list
        End Function

        ''' <summary>
        ''' 取込ファイルテーブル(imp_files)に取込ファイルワークテーブル(imp_files_stage)のレコードを追加する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="impFileStageId">一時取込ファイルID</param>
        ''' <returns>新規作成された IMP_FILE_ID</returns>
        Public Function InsertImpFileFromStage(ByVal tran As OracleTransaction,
                                       ByVal impFileStageId As Long,
                                       ByVal updatedAt As DateTime,
                                       ByVal updateUserId As String,
                                       ByVal updatePgId As String) As Long

            Dim newImpFileId As Long = 0
            Dim conn As OracleConnection = tran.Connection

            'ID (Sequence) を取得する
            ' ※アイデンティティ列(ISEQ$$...)の場合は、直接 .NEXTVAL を実行して取得します　※取得先:IMP_FILESに設定されているシーケンスID
            '天方環境
            Const sqlGetId As String = "SELECT ""OMSDB"".""ISEQ$$_77189"".NEXTVAL FROM DUAL"
            '本番環境
            'Const sqlGetId As String = "SELECT ""OMSDB"".""ISEQ$$_66841"".NEXTVAL FROM DUAL"
            '検証環境
            'Const sqlGetId As String = "SELECT ""OMSTS"".""ISEQ$$_66717"".NEXTVAL FROM DUAL"
            Using cmdId As New OracleCommand(sqlGetId, conn)
                cmdId.Transaction = tran
                newImpFileId = Convert.ToInt64(cmdId.ExecuteScalar())
            End Using

            '取得した ID を含めて INSERT ... SELECT を実行する
            Dim sqlInsert As String = "
            INSERT INTO imp_files (
                imp_file_id,
                customer_setting_id, folder_type, folder_path, file_name,
                staged_folder_path, staged_file_name, reconcile_flag, fcst_reconcile_flag,
                hand_flag, created_at, created_user_id, created_pg_id,
                updated_at, updated_user_id, updated_pg_id
            )
            SELECT 
                :p_new_id,
                customer_setting_id, folder_type, folder_path, file_name,
                staged_folder_path, staged_file_name, reconcile_flag, fcst_reconcile_flag,
                hand_flag, :p_updated_at, :p_user_id, :p_pg_id,
                :p_updated_at, :p_user_id, :p_pg_id
            FROM imp_files_stage
            WHERE imp_file_stage_id = :p_imp_file_stage_id
              AND status = 'PARSED'"

            Using cmd As New OracleCommand(sqlInsert, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_new_id", OracleDbType.Long).Value = newImpFileId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.TimeStamp).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2).Value = updateUserId
                cmd.Parameters.Add(":p_pg_id", OracleDbType.Varchar2).Value = updatePgId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId

                cmd.ExecuteNonQuery()
            End Using

            Return newImpFileId
        End Function

        ''' <summary>
        ''' 正規データ更新 (IMP_FILE_STAGEテーブルのステータスをFAILEDに更新する)
        ''' </summary>
        ''' <param name="impFileStageId">現在処理中の一時取込ファイルID</param>
        Public Sub UpdateImpFileStageStatus(ByVal impFileStageId As Long)

            Const sql As String = "
                    UPDATE imp_files_stage
                    SET status = 'FAILED'
                    WHERE imp_file_stage_id = :p_imp_file_stage_id"

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Clear()

                    ' パラメータの設定
                    cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId

                    ' 実行
                    cmd.ExecuteNonQuery()

                End Using
            End Using

        End Sub

        '2026/03/26 酒井 ed

    End Class

    ' DTO（IMP_FILES_STAGE へ渡すため）
    Public Class ImpFilesStageRow
        Public Property ImpFileStageId As Long          ' IMP_FILE_STAGE_ID
        Public Property CustomerSettingId As Long?      ' CUSTOMER_SETTING_ID
        Public Property FolderType As Integer?          ' FOLDER_TYPE
        Public Property FolderPath As String            ' FOLDER_PATH
        Public Property FileName As String              ' FILE_NAME
        Public Property StagedFolderPath As String      ' STAGED_FOLDER_PATH
        Public Property StagedFileName As String        ' STAGED_FILE_NAME

        ' 画面選択に応じてセット
        Public Property ReconcileFlag As String         ' RECONCILE_FLAG
        Public Property FcstReconcileFlag As String     ' FCST_RECONCILE_FLAG
        Public Property HandFlag As String              ' HAND_FLAG

        ' ステータス
        Public Property Status As String                ' STATUS

        ' 監査系
        Public Property CreatedAt As DateTime?          ' CREATED_AT
        Public Property CreatedUserId As String         ' CREATED_USER_ID
        Public Property CreatedPgId As String           ' CREATED_PG_ID
        Public Property UpdatedAt As DateTime?          ' UPDATED_AT
        Public Property UpdatedUserId As String         ' UPDATED_USER_ID
        Public Property UpdatedPgId As String           ' UPDATED_PG_ID

        Public Sub New()

        End Sub

        Public Sub New(
                        CustomerSettingId As Long,
                        FolderType As Integer,
                        FolderPath As String,
                        FileName As String,
                        StagedFolderPath As String,
                        StagedFileName As String,
                        ReconcileFlag As String,
                        FcstReconcileFlag As String,
                        HandFlag As String,
                        Status As String,
                        CreatedAt As DateTime,
                        CreatedUserId As String,
                        CreatedPgId As String,
                        UpdatedAt As DateTime,
                        UpdatedUserId As String,
                        UpdatedPgId As String)
            Me.CustomerSettingId = CustomerSettingId
            Me.FolderType = FolderType
            Me.FolderPath = FolderPath
            Me.FileName = FileName
            Me.StagedFolderPath = StagedFolderPath
            Me.StagedFileName = StagedFileName
            Me.ReconcileFlag = ReconcileFlag
            Me.FcstReconcileFlag = FcstReconcileFlag
            Me.HandFlag = HandFlag
            Me.Status = Status
            Me.CreatedAt = CreatedAt
            Me.CreatedUserId = CreatedUserId
            Me.CreatedPgId = CreatedPgId
            Me.UpdatedAt = UpdatedAt
            Me.UpdatedUserId = UpdatedUserId
            Me.UpdatedPgId = UpdatedPgId

        End Sub

        Public Sub New(src As ImpFilesStageRow)

            Me.CustomerSettingId = src.CustomerSettingId
            Me.FolderType = src.FolderType
            Me.FolderPath = src.FolderPath
            Me.FileName = src.FileName
            Me.StagedFolderPath = src.StagedFolderPath
            Me.StagedFileName = src.StagedFileName
            Me.ReconcileFlag = src.ReconcileFlag
            Me.FcstReconcileFlag = src.FcstReconcileFlag
            Me.HandFlag = src.HandFlag
            Me.Status = src.Status
            Me.CreatedAt = src.CreatedAt
            Me.CreatedUserId = src.CreatedUserId
            Me.CreatedPgId = src.CreatedPgId
            Me.UpdatedAt = src.UpdatedAt
            Me.UpdatedUserId = src.UpdatedUserId
            Me.UpdatedPgId = src.UpdatedPgId

        End Sub
        ''' <summary>
        ''' ImpFilesStageRow から ImpFilesRowクラス変換
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        Public Shared Function ToImpFilesRow(src As ImpFilesStageRow) As ImpFilesRow

            Dim dst = New ImpFilesRow()

            dst.CustomerSettingId = src.CustomerSettingId
            dst.FolderType = src.FolderType
            dst.FolderPath = src.FolderPath
            dst.FileName = src.FileName
            dst.StagedFolderPath = src.StagedFolderPath
            dst.StagedFileName = src.StagedFileName
            dst.ReconcileFlag = src.ReconcileFlag
            dst.FcstReconcileFlag = src.FcstReconcileFlag
            dst.HandFlag = src.HandFlag
            dst.CreatedAt = src.CreatedAt
            dst.CreatedUserId = src.CreatedUserId
            dst.CreatedPgId = src.CreatedPgId
            dst.UpdatedAt = src.UpdatedAt
            dst.UpdatedUserId = src.UpdatedUserId
            dst.UpdatedPgId = src.UpdatedPgId
            Return dst

        End Function


    End Class

End Namespace
