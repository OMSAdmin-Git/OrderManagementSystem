Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.CustomerSetting

Namespace Pages.Masters.Folder
    Public Class FolderSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        ' id1 = FolderId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property FolderId As Long
            Get
                Dim s As String = If(Request.QueryString("id1"), "").Trim()
                Dim v As Long = 0
                Long.TryParse(s, v)
                Return v
            End Get
        End Property
#End Region

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                PageHelpers.SetUserName(Me, lblUser)
                LoadFolderHeader(FolderId)
            End If
        End Sub
#End Region

#Region "フォルダマスタデータ"
        Private Sub LoadFolderHeader(folderId As Long)

            If folderId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New FolderRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetFolder(folderId)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                SetHeaderControls(Nothing)
                lblError.Text = "選択されたデータが見つかりません。"
                Exit Sub
            End If

            SetHeaderControls(dt.Rows(0))
        End Sub

        ' 選択データセット（Nothing なら空クリア）
        Private Sub SetHeaderControls(r As DataRow)

            txtCustomerCode.Text = If(GetStr(r, "CustomerCode"), "")
            txtCustomerName.Text = If(GetStr(r, "CustomerName"), "")
            txtProfitCenter.Text = If(GetStr(r, "ProfitCenter"), "")
            txtCustomerUnitName.Text = If(GetStr(r, "CustomerUnitName"), "")

            Dim folderTypeText As String = ""
            Dim folderTypeStr As String = GetStr(r, "FolderType")
            Dim folderTypeVal As Integer

            If Integer.TryParse(folderTypeStr, folderTypeVal) Then
                If FolderTypeMap.ContainsKey(folderTypeVal) Then
                    folderTypeText = FolderTypeMap(folderTypeVal)
                Else
                    folderTypeText = $"不明({folderTypeVal})"
                End If
            Else
                folderTypeText = ""
            End If
            txtFolderType.Text = folderTypeText

            txtFolderPath.Value = If(GetStr(r, "FolderPath"), "")
            ddlActiveFlag.SelectedValue = If(GetStr(r, "ActiveFlag"), "")
            txtUpdatedAt.Text = If(GetStr(r, "UpdatedAt"), "")
            txtUpdatedUserName.Text = If(GetStr(r, "UpdatedUserId"), "")

        End Sub

        ' DataRowの項目を文字列で安全取得
        Private Function GetStr(r As DataRow, columnName As String) As String
            If r Is Nothing Then Return Nothing
            If Not r.Table.Columns.Contains(columnName) Then Return Nothing
            Dim v = r(columnName)
            If v Is DBNull.Value OrElse v Is Nothing Then Return Nothing
            Return v.ToString()
        End Function
#End Region

#Region "ボタンイベント"
        ' 保存ボタン
        Protected Sub btnSaveFolderSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try

                If FolderId <= 0 Then
                    lblError.Text = "対象IDが不正です。再度ページを開き直してください。"
                    Return
                End If

                ' 入力値取得
                Dim folderPath As String = (If(txtFolderPath.Value, "")).Trim()
                Dim activeFlag As String = If(ddlActiveFlag.SelectedValue, "").Trim().ToUpperInvariant()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim pgId As String = "FolderSetting(Update)"

                ' 入力確認
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New FolderRepository(Utils.GetConnectionString())

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repo.ExistsFolderPath(
                    folderPath:=folderPath,
                    excludeFolderId:=FolderId
                )
                If existsOther Then
                    lblError.Text = "同一（フォルダパス）の登録が見つかりました。"
                    Return
                End If

                ' UPDATE
                Dim affected As Integer = repo.UpdateFolderWithConcurrency(
                    folderId:=FolderId,
                    folderPath:=folderPath,
                    activeFlag:=activeFlag,
                    loginUserId:=loginUserId,
                    programId:=pgId
                )

                If affected = -1 Then
                    lblError.Text = "対象データが見つかりませんでした。"
                    Return
                ElseIf affected = 0 Then
                    lblError.Text = "他のユーザーにより更新されています。再読み込みしてからやり直してください。"
                    Return
                End If

                lblResult.Text = "登録情報を更新しました。"
                LoadFolderHeader(FolderId)

            Catch ex As ApplicationException
                lblError.Text = Server.HtmlEncode(ex.Message)

            Catch ex As Oracle.ManagedDataAccess.Client.OracleException
                lblError.Text = "DBエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)

            Catch ex As Exception
                lblError.Text = "予期しないエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)
            End Try
        End Sub
#End Region

#Region "OnPreRender"
        Protected Overrides Sub OnPreRender(e As EventArgs)
            MyBase.OnPreRender(e)
            ' lblError が DOM に出るタイミングで確実に ClientID が決まっている
            btnSaveFolderSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace