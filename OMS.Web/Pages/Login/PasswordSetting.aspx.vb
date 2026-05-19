Imports System
Imports System.Web.Security
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client

Public Class PasswordSetting
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            txtUserID.Text = ""
            txtCurrentPassword.Text = ""
            txtNewPassword.Text = ""
            txtConfirmPassword.Text = ""
        End If
    End Sub

#Region "ページ初期化（Page_Init）"
    Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
        Dim connStr = System.Configuration.ConfigurationManager.ConnectionStrings("OMSConnection").ConnectionString
        _repo = New UserRepository(connStr)
    End Sub
#End Region

#Region "フィールド"
    Private _logger As Logger
    Private _repo As UserRepository
#End Region

#Region "メニュー遷移（ナビゲーション）"
    Protected Sub btnReturn_Click(sender As Object, e As EventArgs)
        Response.Redirect("./Login.aspx", False)
    End Sub
#End Region

#Region "ボタンイベント"
    '更新ボタン
    Protected Sub btnSavePasswordSetting_Click(sender As Object, e As EventArgs) Handles btnSavePasswordSetting.Click
        Dim userId As String = txtUserID.Text.Trim()
        Dim currentPassword As String = txtCurrentPassword.Text  ' パスワードは Trim しない
        Dim newPassword As String = txtNewPassword.Text
        Dim confirmPassword As String = txtConfirmPassword.Text
        Dim pgId As String = "SaveUserSetting"
        Dim pgIdDetail As String = String.Empty

        ' 入力バリデーション（最低限）
        If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(newPassword) OrElse String.IsNullOrEmpty(confirmPassword) Then
            lblResult.Text = ""
            lblError.Text = "ユーザー名とパスワードを入力してください。"
            Exit Sub
        End If

        ' STRAMMIC側にユーザーが存在するか確認
        Dim displayName = _repo.ValidateAndGetDisplayName(userId, newPassword)
        If String.IsNullOrEmpty(displayName) Then
            lblResult.Text = ""
            lblError.Text = "ユーザーが存在しません。"
            Exit Sub
        End If

        ' パスワードと確認用パスワードの一致確認
        If newPassword <> confirmPassword Then
            lblResult.Text = ""
            lblError.Text = "新しいパスワードと確認用パスワードが一致しません。"
            Exit Sub
        End If

        ' パスワードの入力規則チェック（8桁以上12桁以下、半角大文字・半角小文字・半角数字・記号（-(ハイフン)または_(アンダーバー)）を1文字以上含める）
        Dim passwordPattern As String = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[-_])[A-Za-z\d-_]{8,12}$"
        If Not System.Text.RegularExpressions.Regex.IsMatch(newPassword, passwordPattern) Then
            lblResult.Text = ""
            lblError.Text = "パスワードの入力規則に違反しています。"
            Exit Sub
        End If

        ' USER_MSTにユーザーが登録されているか確認（されている場合は現在のパスワードを取得）
        Dim strPassword As String = _repo.GetUserPassword(userId)

        Try
            If String.IsNullOrEmpty(strPassword) Then
                '新規ユーザー追加
                pgIdDetail = "Insert"
                _repo.InsertUser(
                        userId,
                        newPassword,
                        pgId:=pgId & "(" & pgIdDetail & ")"
                    )
            Else
                '既存ユーザー更新
                'パスワードの確認
                If currentPassword <> strPassword Then
                    lblResult.Text = ""
                    lblError.Text = "現在のパスワードが正しくありません。"
                    Exit Sub
                End If
                pgIdDetail = "Update"
                _repo.UpdateUser(
                        userId,
                        newPassword,
                        pgId:=pgId & "(" & pgIdDetail & ")"
                    )
            End If
            lblResult.Text = "パスワードの設定が完了しました。"
            lblError.Text = ""
        Catch ex As OracleException
            lblResult.Text = ""
            lblError.Text = "システムエラーが発生しました。時間をおいて再度お試しください。"
        Catch ex As Exception
            lblResult.Text = ""
            lblError.Text = "システムエラーが発生しました。"
        End Try
    End Sub
#End Region
End Class