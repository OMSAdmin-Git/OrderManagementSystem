Imports System.Data
Imports System.Text
Imports System.Web.Script.Services
Imports System.Web.Services
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client

Namespace Pages.Masters.CustomerUnit
    Public Class CustomerUnit
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        ' id1 = CustomerUnitId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property CustomerUnitId As Long
            Get
                Dim s As String = If(Request.QueryString("id1"), "").Trim()
                Dim v As Long = 0
                Long.TryParse(s, v)
                Return v
            End Get
        End Property
#End Region

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                PageHelpers.SetUserName(Me, lblUser)
                LoadCustomerUnitHeader(CustomerUnitId)
            End If
        End Sub
#End Region

#Region "注文工場／担当者マスタデータ"
        Private Sub LoadCustomerUnitHeader(customerUnitId As Long)

            If customerUnitId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New CustomerUnitRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetCustomerUnit(customerUnitId)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                SetHeaderControls(Nothing)
                lblError.Text = "選択されたデータが見つかりません。"
                Exit Sub
            End If

            SetHeaderControls(dt.Rows(0))
        End Sub

        ' 選択データセット
        Private Sub SetHeaderControls(r As DataRow)
            txtCustomerUnitId.Text = If(GetStr(r, "CustomerUnitId"), "")
            txtCustomerUnitName.Value = If(GetStr(r, "CustomerUnitName"), "")
            ddlActiveFlag.SelectedValue = GetStr(r, "ActiveFlag")
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
        Protected Sub btnSaveCustomerUnitSetting_Click(sender As Object, e As EventArgs)

            lblError.Text = ""
            lblResult.Text = ""

            Try

                If CustomerUnitId <= 0 Then
                    lblError.Text = "対象IDが不正です。再度ページを開き直してください。"
                    Return
                End If

                ' 入力値取得
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim activeFlag As String = If(ddlActiveFlag.SelectedValue, "").Trim().ToUpperInvariant()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim pgId As String = "CustomerUnitSetting(Update)"

                ' 入力確認
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New CustomerUnitRepository(Utils.GetConnectionString())

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repo.ExistsCustomerUnit(
                    customerUnitName:=customerUnitName,
                    excludeCustomerUnitId:=CustomerUnitId
                )
                If existsOther Then
                    lblError.Text = "同一の設定が見つかりました。"
                    Return
                End If

                ' UPDATE
                Dim affected As Integer = repo.UpdateCustomerUnitWithConcurrency(
                    customerUnitId:=CustomerUnitId,
                    customerUnitName:=customerUnitName,
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
                LoadCustomerUnitHeader(CustomerUnitId)

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
            btnSaveCustomerUnitSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace