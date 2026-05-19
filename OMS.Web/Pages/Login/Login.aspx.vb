Imports System
Imports System.Web.Security
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client

Namespace Pages.Login
    Partial Public Class Login
        Inherits System.Web.UI.Page

#Region "フィールド"
        Private _logger As Logger
        Private _repo As UserRepository
#End Region

#Region "ページ初期化（Page_Init）"
        Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
            Dim logPath = Server.MapPath("~/App_Data/logs/app.log")
            _logger = New Logger(logPath)

            Dim connStr = System.Configuration.ConfigurationManager.ConnectionStrings("OMSConnection").ConnectionString
            _repo = New UserRepository(connStr)
        End Sub
#End Region

#Region "ボタンイベント"
        'ログインボタン
        Protected Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
            Dim userId As String = txtUser.Text.Trim()
            Dim password As String = txtPass.Text  ' パスワードは Trim しない

            ' 入力バリデーション（最低限）
            If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(password) Then
                lblError.Text = "ユーザー名とパスワードを入力してください。"
                Exit Sub
            End If

            Dim clientIp = GetClientIp()
            Dim ua = Request.UserAgent

            Try
                ' ▼ 1回のクエリで認証＋表示名取得
                Dim displayName = _repo.ValidateAndGetDisplayName(userId, password)
                If String.IsNullOrEmpty(displayName) Then
                    displayName = GetDisplayNameAdmin(userId, password)
                End If

                If Not String.IsNullOrEmpty(displayName) Then
                    _logger.Write($"LOGIN SUCCESS user={userId} ip={clientIp} ua={ua}")

                    'ログインユーザーがユーザーマスタ（USER_MST）に有効で登録されているか確認（登録されていたらパスワードを返す）
                    Dim strPassword As String = _repo.GetUserPasswordActive(userId)
                    If String.IsNullOrEmpty(strPassword) And displayName <> "admin" Then
                        lblError.Text = "ユーザーが登録されていません。パスワード設定を行ってください。"
                        Exit Sub
                    End If

                    'パスワードの確認
                    If password <> strPassword And displayName <> "admin" Then
                        lblError.Text = "ユーザー名またはパスワードが正しくありません。"
                        Exit Sub
                    End If

                    ' 認証チケット（Forms認証を使う場合）
                    FormsAuthentication.SetAuthCookie(userId, False)

                    ' TopMenu 用のセッション
                    Session("UserId") = userId
                    Session("UserName") = displayName
                    If displayName = "admin" Then
                        Session("UserAuthority") = "1"
                    Else
                        Session("UserAuthority") = _repo.GetUserAuthority(userId)
                    End If


                    Response.Redirect("../TopMenu.aspx", False)
                    Context.ApplicationInstance.CompleteRequest()
                    Return
                Else
                    _logger.Write($"LOGIN FAIL user={userId} ip={clientIp} ua={ua}")
                    lblError.Text = "ユーザー名またはパスワードが正しくありません。"
                End If

            Catch ex As OracleException
                _logger.Write($"LOGIN ERROR(Oracle) user={userId} ip={clientIp} code={ex.Number} msg={ex.Message}")
                lblError.Text = "システムエラーが発生しました。時間をおいて再度お試しください。"
            Catch ex As Exception
                _logger.Write($"LOGIN ERROR user={userId} ip={clientIp} msg={ex.Message}")
                lblError.Text = "システムエラーが発生しました。"
            End Try
        End Sub

        'パスワード設定ボタン
        Protected Sub btnPassword_Click(sender As Object, e As EventArgs) Handles btnPassword.Click
            Response.Redirect("./PasswordSetting.aspx", False)
        End Sub
#End Region

#Region "プライベートメソッド"
        Private Function GetClientIp() As String
            Dim ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            If Not String.IsNullOrEmpty(ip) Then
                ' 複数ある場合は最初を採用
                Return ip.Split(","c)(0).Trim()
            End If
            Return Request.ServerVariables("REMOTE_ADDR")
        End Function
#End Region

    End Class
End Namespace