Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.CustomerUnit
    Public Class CustomerUnitSettingCreate
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
            End If
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnCustomerUnitList_Click(sender As Object, e As EventArgs)
            Response.Redirect("CustomerUnitList.aspx")
        End Sub

        ' 登録ボタン
        Protected Sub btnCreateCustomerUnitSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try
                ' 入力取得
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim programId As String = "CustomerUnitSettingCreate(Insert)"

                ' 必須チェック
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New CustomerUnitRepository(Utils.GetConnectionString())

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repo.ExistsCustomerUnit(
                    customerUnitName:=customerUnitName,
                    excludeCustomerUnitId:=0
                )
                If existsOther Then
                    lblError.Text = "同一（注文工場／担当者名）の登録が見つかりました。"
                    Return
                End If

                ' INSERT
                Dim activeFlag As String = "Y"
                Dim newId As Long = repo.InsertCustomerUnitNullable(
                    customerUnitName:=customerUnitName,
                    activeFlag:=activeFlag,
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
            btnCreateCustomerUnitSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace