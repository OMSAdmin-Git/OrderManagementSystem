Imports DocumentFormat.OpenXml.Vml.Wordprocessing
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.CustomerSetting

Namespace Pages.Masters.SuzukiSpiritsConversion
    Public Class SuzukiSpiritsConversionSettingCreate
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
            End If
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnSuzukiSpiritsConversionList_Click(sender As Object, e As EventArgs)
            Response.Redirect("SuzukiSpiritsConversionList.aspx")
        End Sub

        ' 登録ボタン（未実装）
        Protected Sub btnCreateSuzukiSpiritsConversionSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try
                '入力取得
                Dim deliveryCodeOrder As String = (If(txtDeliveryCodeOrder.Value, "")).Trim()
                Dim deliveryCodePlan As String = (If(txtDeliveryCodePlan.Value, "")).Trim()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim programId As String = "FolderSetting(Update)"


                ' 必須チェック
                ' JSで処理しているためなし。必要なら追加。

                Dim Deli As New SuzukiSpiritsConversionRepository(Utils.GetConnectionString())

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = Deli.ExistsSuzukiSpiritsConversion(
                deliveryCodeOrder:=deliveryCodeOrder,
                    excludeConversionId:=0 '新規のため除外なし
                )

                If existsOther = True Then
                    lblError.Text = "同一（納入指示納入先コード）の登録が見つかりました。"
                    Return
                End If

                Dim activeflag As String = "Y"
                Dim newId As Long = Deli.InsertSuzukiSpiritsConversionNullable(
                    deliveryCodePlan:=deliveryCodePlan,
                    deliveryCodeOrder:=deliveryCodeOrder,
                    activeFlag:=activeflag,
                    loginUserId:=loginUserId,
                    programId:=programId
                )

                If newId <= 0 Then
                    Throw New ApplicationException("入力情報の登録に失敗しました。")
                End If

                ' 完了メッセージ
                lblResult.Text = "入力情報を登録しました。"


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
            btnCreateSuzukiSpiritsConversionSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace

