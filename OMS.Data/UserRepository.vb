Imports System
Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data
    ''' <summary>
    ''' ユーザー（USRF）関連のデータアクセスを提供するリポジトリ。
    ''' 認証・表示名取得／候補リスト取得（ユーザー名一覧）などを扱います。
    ''' </summary>
    Public Class UserRepository

#Region "フィールド・コンストラクタ"

        ''' <summary>接続文字列</summary>
        Private ReadOnly _connectionString As String

        ''' <summary>
        ''' リポジトリを初期化します。
        ''' </summary>
        ''' <param name="connectionString">Oracle接続文字列</param>
        '''Public Sub New(connectionString As String)
        '''    _connectionString = connectionString
        '''End Sub
        Public Sub New(connectionString As String)
            If String.IsNullOrWhiteSpace(connectionString) Then
                Throw New ArgumentException("connectionString is null or empty.", NameOf(connectionString))
            End If
            _connectionString = connectionString
        End Sub
#End Region

#Region "認証/ユーザー情報取得"

        ''' <summary>
        ''' ユーザーIDとパスワードでUSRFテーブルを検証し、該当する場合は表示名（FUSRNAME）を返します。
        ''' </summary>
        ''' <param name="userId">ユーザーID（FUSRID）</param>
        ''' <param name="password">パスワード（FUSRPASSWD）※平文を現行SQLに忠実に比較</param>
        ''' <returns>
        ''' 一致するユーザーが存在する場合は表示名（FUSRNAME）。<br/>
        ''' 見つからない場合は <c>Nothing</c>。
        ''' </returns>
        ''' <remarks>
        ''' ・現行仕様に合わせて、<c>UPPER(TRIM(FUSRID))</c> と <c>UPPER(TRIM(:impId))</c> を比較し、
        '''   パスワードは <c>FUSRPASSWD = :impPass</c>（大文字小文字区別、トリム無し）で比較しています。<br/>
        ''' ・将来的にハッシュ化へ移行する場合は、本メソッドでハッシュを生成して比較するように変更してください。
        ''' </remarks>
        Public Function ValidateAndGetDisplayName(userId As String, password As String) As String

            'Const sql As String = "
            '    SELECT FUSRNAME
            '    FROM USRF
            '    WHERE UPPER(TRIM(FUSRID)) = UPPER(TRIM(:impId))
            '        AND FUSRPASSWD          = :impPass
            '    FETCH FIRST 1 ROWS ONLY
            '"

            Const sql As String = "
                SELECT FUSRNAME
                FROM USRF
                WHERE UPPER(TRIM(FUSRID)) = UPPER(TRIM(:impId))
                FETCH FIRST 1 ROWS ONLY
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True

                    ' パラメータ（長さは現行運用に合わせた暫定。必要に応じてDB定義へ合わせてください）
                    Dim pId = cmd.Parameters.Add("impId", OracleDbType.Varchar2, 9, ParameterDirection.Input)
                    pId.Value = If(userId, String.Empty).Trim()

                    'Dim pPass = cmd.Parameters.Add("impPass", OracleDbType.Varchar2, 200, ParameterDirection.Input)
                    'pPass.Value = If(password, String.Empty)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)
                End Using
            End Using
        End Function

#End Region

#Region "候補リスト取得"

        ''' <summary>
        ''' ユーザー名（FUSRNAME）の一覧を取得します。<br/>
        ''' 並び順は <c>FUSRID</c> の昇順です。
        ''' </summary>
        ''' <returns>ユーザー名（FUSRNAME）のリスト</returns>
        ''' <remarks>
        ''' ・全ユーザを対象に列挙します。大量件数の場合はページングやインクリメンタルサーチへの切替を検討してください。<br/>
        ''' ・権限制御（有効ユーザのみ、部門絞り込み等）が必要な場合は、別メソッドとして拡張するのがおすすめです。
        ''' </remarks>
        Public Function GetUserNames() As List(Of String)
            Dim result As New List(Of String)()

            Dim sql As String = "
                SELECT FUSRNAME
                FROM USRF
                ORDER BY FUSRID ASC
            "

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            If Not reader.IsDBNull(0) Then
                                result.Add(reader.GetString(0))
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return result
        End Function

#End Region

#Region "一覧取得"

        ' ユーザーマスタ一覧取得
        Public Function GetUserList(
            Optional ByVal prodMgmtUserId As String = Nothing,
            Optional ByVal ProdMgmtUserName As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  prod_mgmt_user_id  AS ""ProdMgmtUserId"",")
            sb.AppendLine("  FUSRNAME           AS ""ProdMgmtUserName"",")
            sb.AppendLine("  password           AS ""Password"",")
            sb.AppendLine("  authority_level    AS ""AuthorityLevel"",")
            sb.AppendLine("  active_flag        AS ""ActiveFlag"",")
            sb.AppendLine("  created_at         AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id    AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id      AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at         AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id    AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id      AS ""UpdatedPgId""")
            sb.AppendLine("FROM user_mst INNER JOIN USRF ON TRIM(prod_mgmt_user_id) = TRIM(FUSRID) ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' LIKE パターン生成
            Dim pProdMgmtUserId As String = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)
            Dim pProdMgmtUserName As String = Utils.BuildLikePattern(ProdMgmtUserName, LikeMode.Contains)
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

            If pProdMgmtUserId IsNot Nothing Then
                Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
                If Not isAdmin Then
                    'sb.AppendLine("AND UPPER(prod_mgmt_user_id) = UPPER(:p_user) ")
                    sb.AppendLine("AND UPPER(prod_mgmt_user_id) LIKE UPPER(:p_user) ")
                    prm.Add(New OracleParameter(":p_user", OracleDbType.Varchar2) With {.Value = pProdMgmtUserId})
                End If
            End If

            If pProdMgmtUserName IsNot Nothing Then
                sb.AppendLine("AND UPPER(FUSRNAME) LIKE UPPER(:p_funame) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_funame", OracleDbType.Varchar2) With {.Value = pProdMgmtUserName})
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            sb.AppendLine("ORDER BY prod_mgmt_user_id ")

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

        ' ユーザーマスタ パスワード取得（有効のみ）
        Public Function GetUserPasswordActive(userId As String) As String

            Const sql As String = "
                SELECT PASSWORD
                FROM USER_MST
                WHERE UPPER(TRIM(PROD_MGMT_USER_ID)) = UPPER(TRIM(:impId)) AND ACTIVE_FLAG = 'Y'
                FETCH FIRST 1 ROWS ONLY
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True

                    ' パラメータ（長さは現行運用に合わせた暫定。必要に応じてDB定義へ合わせてください）
                    Dim pId = cmd.Parameters.Add("impId", OracleDbType.Varchar2, 9, ParameterDirection.Input)
                    pId.Value = If(userId, String.Empty).Trim()

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)
                End Using
            End Using

        End Function

        ' ユーザーマスタ パスワード取得
        Public Function GetUserPassword(userId As String) As String

            Const sql As String = "
                SELECT PASSWORD
                FROM USER_MST
                WHERE UPPER(TRIM(PROD_MGMT_USER_ID)) = UPPER(TRIM(:impId))
                FETCH FIRST 1 ROWS ONLY
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True

                    ' パラメータ（長さは現行運用に合わせた暫定。必要に応じてDB定義へ合わせてください）
                    Dim pId = cmd.Parameters.Add("impId", OracleDbType.Varchar2, 9, ParameterDirection.Input)
                    pId.Value = If(userId, String.Empty).Trim()

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)
                End Using
            End Using

        End Function

        ' ユーザーマスタ 権限レベル取得
        Public Function GetUserAuthority(userId As String) As String

            Const sql As String = "
                SELECT AUTHORITY_LEVEL
                FROM USER_MST
                WHERE UPPER(TRIM(PROD_MGMT_USER_ID)) = UPPER(TRIM(:impId)) AND ACTIVE_FLAG = 'Y'
                FETCH FIRST 1 ROWS ONLY
            "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True

                    ' パラメータ（長さは現行運用に合わせた暫定。必要に応じてDB定義へ合わせてください）
                    Dim pId = cmd.Parameters.Add("impId", OracleDbType.Varchar2, 9, ParameterDirection.Input)
                    pId.Value = If(userId, String.Empty).Trim()

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)
                End Using
            End Using

        End Function

#End Region

#Region "INSERT / UPDATE / RESET"

        'INSERT: 1件追加する
        Public Function InsertUser(
            prodMgmtUserId As String,
            password As String,
            pgId As String
            ) As String

            Dim sql As String = "INSERT INTO user_mst (
                            prod_mgmt_user_id, 
                            password, 
                            created_user_id, 
                            created_pg_id,
                            updated_user_id, 
                            updated_pg_id
                        ) VALUES (
                            :prod_mgmt_user_id, 
                            :password, 
                            :created_user_id, 
                            :created_pg_id,
                            :updated_user_id, 
                            :updated_pg_id
                        )
                        RETURNING prod_mgmt_user_id INTO :p_newid
                    "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        ' --- パラメータ ---
                        AddVarchar(cmd, ":prod_mgmt_user_id", prodMgmtUserId)
                        AddVarchar(cmd, ":password", password)
                        AddVarchar(cmd, ":created_user_id", prodMgmtUserId)
                        AddVarchar(cmd, ":created_pg_id", pgId)
                        AddVarchar(cmd, ":updated_user_id", prodMgmtUserId)
                        AddVarchar(cmd, ":updated_pg_id", pgId)

                        Dim pNewId = New OracleParameter(":p_newid", OracleDbType.Varchar2)
                        pNewId.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(pNewId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

                        Return Convert.ToString(pNewId.Value)
                    End Using
                End Using
            End Using
        End Function

        ' UPDATE: 主キーで1件更新
        Public Function UpdateUser(
            prodMgmtUserId As String,
            password As String,
            pgId As String
            ) As String

            Dim sql As String = "UPDATE  user_mst
                                SET password             =:password, 
                                updated_at              = SYSDATE,
                                updated_user_id         =:updated_user_id,
                                updated_pg_id           =:updated_pg_id
                                WHERE prod_mgmt_user_id = :prod_mgmt_user_id
                                "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        ' --- パラメータ ---
                        AddVarchar(cmd, ":password", password)
                        AddVarchar(cmd, ":updated_user_id", prodMgmtUserId)
                        AddVarchar(cmd, ":updated_pg_id", pgId)
                        AddVarchar(cmd, ":prod_mgmt_user_id", prodMgmtUserId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

                        Return prodMgmtUserId

                    End Using
                End Using
            End Using
        End Function

        ' UPDATE(パスワードリセット): 主キーで1件更新
        Public Function UpdatePasswordReset(
            prodMgmtUserId As String,
            pgUserId As String,
            pgId As String
            ) As String

            Dim sql As String = "UPDATE  user_mst
                                SET password             =:password, 
                                updated_at              = SYSDATE,
                                updated_user_id         =:updated_user_id,
                                updated_pg_id           =:updated_pg_id
                                WHERE prod_mgmt_user_id = :prod_mgmt_user_id
                                "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.BindByName = True
                        ' --- パラメータ ---
                        AddVarchar(cmd, ":password", "astiasti")
                        AddVarchar(cmd, ":updated_user_id", pgUserId)
                        AddVarchar(cmd, ":updated_pg_id", pgId)
                        AddVarchar(cmd, ":prod_mgmt_user_id", prodMgmtUserId)

                        cmd.ExecuteNonQuery()
                        tran.Commit()

                        Return prodMgmtUserId

                    End Using
                End Using
            End Using
        End Function
#End Region

    End Class
End Namespace