Imports OMS.Common
Imports OMS.Data

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

