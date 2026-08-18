Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.SuzukiSpiritsConversion
    Public Class SuzukiSpiritsConversionSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        ' id1 = ConversionId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property ConversionId As Long
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
                LoadSuzukiSpiritsConversionHeader(ConversionId)
            End If
        End Sub
#End Region

#Region "スズキSPIRITS納入ホーム変換マスタデータ"
        Private Sub LoadSuzukiSpiritsConversionHeader(conversionId As Long)

            If conversionId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New SuzukiSpiritsConversionRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetSuzukiSpiritsConversion(conversionId)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                SetHeaderControls(Nothing)
                lblError.Text = "選択されたデータが見つかりません。"
                Exit Sub
            End If

            SetHeaderControls(dt.Rows(0))
        End Sub

        ' 選択データセット（Nothing なら空クリア）
        Private Sub SetHeaderControls(r As DataRow)

            txtDeliveryCodeOrder.Text = If(GetStr(r, "deliveryCodeOrder"), "")
            txtDeliveryCodePlan.Text = If(GetStr(r, "deliveryCodePlan"), "")
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
        ' 保存ボタン（未実装）
        Protected Sub btnSaveSuzukiSpiritsConversionSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try

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
            btnSaveSuzukiSpiritsConversionSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace