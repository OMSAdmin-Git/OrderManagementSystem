Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.CustomerSetting

Namespace Pages.Masters.ImpRule
    Public Class ImpRuleSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        ' id1 = ImpRuleId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property ImpRuleId As Long
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
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
                LoadImpRuleHeader(ImpRuleId)
            End If
        End Sub
#End Region

#Region "取込条件マスタデータ"
        Private Sub LoadImpRuleHeader(impRuleId As Long)

            If impRuleId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New ImpRuleRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetImpRule(impRuleId)

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
            ddlProratedType.SelectedValue = If(GetStr(r, "ProratedType"), "")
            ddlReconcileFlag.SelectedValue = If(GetStr(r, "ReconcileFlag"), "")
            ddlFcstReconcileFlag.SelectedValue = If(GetStr(r, "FcstReconcileFlag"), "")
            ddlReconcileType.SelectedValue = If(GetStr(r, "ReconcileType"), "")
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
        Protected Sub btnSaveImpRuleSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try

                If ImpRuleId <= 0 Then
                    lblError.Text = "対象IDが不正です。再度ページを開き直してください。"
                    Return
                End If

                ' 入力値取得
                Dim proratedTypeValueRaw As String = If(ddlProratedType.SelectedValue, "").Trim()
                Dim proratedType As Integer
                If Not Integer.TryParse(proratedTypeValueRaw, proratedType) Then
                    lblError.Text = "分割区分（ProratedType）の選択が不正です。"
                    Return
                End If

                Dim reconcileFlag As String = If(ddlReconcileFlag.SelectedValue, "").Trim()
                Dim fcstReconcileFlag As String = If(ddlFcstReconcileFlag.SelectedValue, "").Trim()

                Dim reconcileTypeValueRaw As String = If(ddlReconcileType.SelectedValue, "").Trim()
                Dim reconcileType As Integer
                If Not Integer.TryParse(reconcileTypeValueRaw, reconcileType) Then
                    lblError.Text = "消込条件（ReconcileType）の選択が不正です。"
                    Return
                End If
                Dim activeFlag As String = If(ddlActiveFlag.SelectedValue, "").Trim().ToUpperInvariant()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim pgId As String = "FolderSetting(Update)"

                ' 入力確認
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New ImpRuleRepository(Utils.GetConnectionString())

                ' UPDATE
                Dim affected As Integer = repo.UpdateImpRuleWithConcurrency(
                    impRuleId:=ImpRuleId,
                    proratedType:=proratedType,
                    reconcileFlag:=reconcileFlag,
                    fcstReconcileFlag:=fcstReconcileFlag,
                    reconcileType:=reconcileType,
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
                LoadImpRuleHeader(ImpRuleId)

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
            btnSaveImpRuleSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace