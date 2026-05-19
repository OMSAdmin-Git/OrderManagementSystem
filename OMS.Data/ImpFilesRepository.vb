
Imports Oracle.ManagedDataAccess.Client
Imports System.Data
Imports System.Text
Imports OMS.Common
Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace OMS.Data
    Public Class ImpFilesRepository
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As ImpFilesRow) As String

            Dim records = New List(Of ImpFilesRow)()
            records.Add(row)
            Return InsertRange(conn, tran, records)

        End Function

        '''' <summary>
        '''' OrderRow class リストをDBに追加する (元のコード同等呼び出し
        '''' R.sagisaka Add
        '''' </summary>
        '''' <param name="records"></param>
        'Public Sub InsertRange(records As IEnumerable(Of ImpFilesRow))

        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            InsertRange(conn, tran, records.ToList())
        '            tran.Commit()
        '        End Using
        '    End Using
        'End Sub

        ''' <summary>
        ''' InsertRange
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns></returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of ImpFilesRow)) As String

            Dim errorMessage As String = ""
            If records Is Nothing Then Return "InsertRange ImpFilesRow: no record data. "
            Try

                'conn.Open()
                Const sql As String =
                "INSERT INTO imp_files (" &
                "  customer_setting_id, folder_type, folder_path, file_name, staged_folder_path, staged_file_name, " &
                "  reconcile_flag, fcst_reconcile_flag, hand_flag, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id " &
                ") VALUES (" &
                "  :p_csid, :p_ftype, :p_folder, :p_fname, :p_sfolder, :p_sfname, " &
                "  :p_recon, :p_fcst_recon, :p_hand, " &
                "  :p_c_at, :p_c_uid, :p_c_pid, " &
                "  :p_u_at, :p_u_uid, :p_u_pid " &
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

        ''' <summary>
        ''' filename で 取込ファイルワークテーブルを取得する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="fileName"></param>
        ''' <param name="folderPath"></param>
        ''' <returns></returns>
        Public Function GetImpFilesFilename(conn As OracleConnection, tran As OracleTransaction, Optional ByVal fileName As String = Nothing, Optional ByVal folderPath As String = Nothing) As ImpFilesRow

            Dim records = GetImpFiles(conn, tran, fileName:=fileName, folderPath:=folderPath)
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
        ''' <param name="impFileId"></param>
        ''' <param name="fileName"></param>
        ''' <param name="folderPath"></param>
        ''' <returns></returns>
        Public Function GetImpFiles(conn As OracleConnection, tran As OracleTransaction,
            Optional ByVal impFileId As Long? = Nothing,
            Optional ByVal fileName As String = Nothing,
            Optional ByVal folderPath As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT * ")
            sb.AppendLine("FROM imp_files ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pFilename As String = Utils.BuildLikePattern(fileName, LikeMode.Contains)
            Dim pFolderPath As String = Utils.BuildLikePattern(folderPath, LikeMode.Contains)
            Dim pImpFileId = impFileId

            If pFilename IsNot Nothing Then
                sb.AppendLine("AND UPPER(file_name) LIKE UPPER(:p_filename) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_filename", OracleDbType.Varchar2) With {.Value = pFilename})
            End If

            If pFolderPath IsNot Nothing Then
                sb.AppendLine("AND UPPER(folder_path) LIKE UPPER(:p_folderPath) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_folderPath", OracleDbType.Varchar2) With {.Value = pFolderPath})
            End If

            If impFileId IsNot Nothing Then
                sb.AppendLine("AND imp_file_id = :p_impFileId ")
                prm.Add(New OracleParameter(":p_impFileId", OracleDbType.Int64) With {.Value = pImpFileId})
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

        ''' <summary>
        ''' DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of ImpFilesRow)
            Dim osrs = New List(Of ImpFilesRow)()
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
        Public Function ToClass(dt As DataRow) As ImpFilesRow
            If dt Is Nothing Then
                Return Nothing
            End If

            Dim osr = New ImpFilesRow
            osr.ImpFileId = dt.Field(Of Long)("imp_file_id")
            osr.CustomerSettingId = dt.Field(Of Long?)("customer_setting_id")
            osr.FolderType = dt.Field(Of Int16?)("folder_type")
            osr.FolderPath = dt.Field(Of String)("folder_path")
            osr.FileName = dt.Field(Of String)("file_name")
            osr.StagedFolderPath = dt.Field(Of String)("staged_folder_path")
            osr.StagedFileName = dt.Field(Of String)("staged_file_name")
            osr.ReconcileFlag = dt.Field(Of String)("reconcile_flag")
            osr.FcstReconcileFlag = dt.Field(Of String)("fcst_reconcile_flag")
            osr.HandFlag = dt.Field(Of String)("hand_flag")
            osr.CreatedAt = dt.Field(Of Date?)("created_at")
            osr.CreatedUserId = dt.Field(Of String)("created_user_id")
            osr.CreatedPgId = dt.Field(Of String)("created_pg_id")
            osr.UpdatedAt = dt.Field(Of Date?)("updated_at")
            osr.UpdatedUserId = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPgId = dt.Field(Of String)("updated_pg_id")
            Return osr
        End Function

    End Class

    Public Class ImpFilesRow
        Public Property ImpFileId As Long               ' IMP_FILE_ID
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
            Me.CreatedAt = CreatedAt
            Me.CreatedUserId = CreatedUserId
            Me.CreatedPgId = CreatedPgId
            Me.UpdatedAt = UpdatedAt
            Me.UpdatedUserId = UpdatedUserId
            Me.UpdatedPgId = UpdatedPgId

        End Sub

        Public Sub New(src As ImpFilesRow)

            Me.CustomerSettingId = src.CustomerSettingId
            Me.FolderType = src.FolderType
            Me.FolderPath = src.FolderPath
            Me.FileName = src.FileName
            Me.StagedFolderPath = src.StagedFolderPath
            Me.StagedFileName = src.StagedFileName
            Me.ReconcileFlag = src.ReconcileFlag
            Me.FcstReconcileFlag = src.FcstReconcileFlag
            Me.HandFlag = src.HandFlag
            Me.CreatedAt = src.CreatedAt
            Me.CreatedUserId = src.CreatedUserId
            Me.CreatedPgId = src.CreatedPgId
            Me.UpdatedAt = src.UpdatedAt
            Me.UpdatedUserId = src.UpdatedUserId
            Me.UpdatedPgId = src.UpdatedPgId

        End Sub
    End Class
End Namespace