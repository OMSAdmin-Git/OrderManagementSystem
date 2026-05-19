Imports OMS.Common
Imports OMS.Data
Imports System.Data
Imports System.Text

Namespace Pages.Masters.CustomerSetting
    Public Class CustomerSettingCreate
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
                ' 検索候補を初期化
                LoadSearchConditionLists()
            End If
        End Sub
#End Region

#Region "検索候補：リスト初期化"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim repo As New CustomerRepository(Utils.GetConnectionString())

            Dim customerCodeList As List(Of String) = repo.GetStraCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = repo.GetStraProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

            lstCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnCustomerList_Click(sender As Object, e As EventArgs)
            Response.Redirect("CustomerList.aspx")
        End Sub

        ' 登録ボタン
        Protected Sub btnCreateCustomerSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try
                ' 入力取得
                Dim customerCode As String = (If(txtCustomerCode.Value, "")).Trim()
                Dim profitCenterRaw As String = (If(txtProfitCenter.Value, "")).Trim()
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim prodMgmtUserName As String = (If(txtProdMgmtUserName.Value, "")).Trim()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim programId As String = "CustomerSettingCreate(Insert)"

                ' 必須チェック
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New CustomerRepository(Utils.GetConnectionString())

                ' 変換/解決
                If Not repo.ExistsCustomerCode(customerCode) Then
                    lblError.Text = "取引先コードが正しく選択されていません。候補から選択してください。"
                    Return
                End If

                ' PC
                Dim profitCenter As String = If(String.IsNullOrWhiteSpace(profitCenterRaw), Nothing, profitCenterRaw)
                If Not String.IsNullOrWhiteSpace(profitCenter) Then
                    If Not repo.ExistsProfitCenter(profitCenter) Then
                        lblError.Text = "PCが正しく選択されていません。候補から選択してください。"
                        Return
                    End If
                End If

                Dim customerUnitIdNullable As Long? = Nothing
                If Not String.IsNullOrWhiteSpace(customerUnitName) Then
                    Dim customerUnitId As Long = repo.GetCustomerUnitIdByName(customerUnitName) ' 未存在/複数は ApplicationException
                    If customerUnitId <= 0 Then
                        lblError.Text = "注文工場／担当者名が正しく選択されていません。候補から選択してください。"
                        Return
                    End If
                    customerUnitIdNullable = customerUnitId
                End If
                Dim prodMgmtUserId As String = repo.GetProdMgmtUserIdByName(prodMgmtUserName).Trim()
                If String.IsNullOrWhiteSpace(prodMgmtUserId) Then
                    lblError.Text = "生産管理担当者名が正しく選択されていません。候補から選択してください。"
                    Return
                End If

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repo.ExistsCustomerSetting(
                    customerCode:=customerCode,
                    profitCenter:=profitCenter,
                    customerUnitId:=If(customerUnitIdNullable.HasValue, customerUnitIdNullable.Value, CType(Nothing, Long?)),
                    excludeCustomerSettingId:=0 ' 新規のため除外なし
                )
                If existsOther Then
                    lblError.Text = "同一（取引先コード、PC、注文工場／担当者名）の登録が見つかりました。"
                    Return
                End If

                ' INSERT
                Dim activeFlag As String = "Y"
                Dim newId As Long = repo.InsertCustomerSettingNullable(
                    customerCode:=customerCode,
                    profitCenter:=profitCenter,               ' Nothing → DB NULL
                    customerUnitId:=customerUnitIdNullable,   ' Nothing → DB NULL
                    prodMgmtUserId:=prodMgmtUserId,
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
            btnCreateCustomerSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

    End Class
End Namespace