Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.InfoType
    Public Class InfoTypeSettingCreate
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

            Dim customerCodeList As List(Of String) = repo.GetCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = repo.GetProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

            lstCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnInfoTypeList_Click(sender As Object, e As EventArgs)
            Response.Redirect("InfoTypeList.aspx")
        End Sub

        ' 登録ボタン
        Protected Sub btnCreateInfoTypeSetting_Click(sender As Object, e As EventArgs)
            lblError.Text = ""
            lblResult.Text = ""

            Try
                ' 入力取得
                Dim customerCode As String = (If(txtCustomerCode.Value, "")).Trim()
                Dim profitCenterRaw As String = (If(txtProfitCenter.Value, "")).Trim()
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim customerInfoType As String = (If(txtCustomerInfoType.Value, "")).Trim()
                Dim infoType As String = (If(ddlInfoType.SelectedValue, "")).Trim().ToUpperInvariant()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim programId As String = "InfoTypeSettingCreate(Insert)"

                ' 必須チェック
                ' JSで処理しているためなし。必要なら追加。

                Dim repoCust As New CustomerRepository(Utils.GetConnectionString())
                Dim repoInfoType As New InfoTypeRepository(Utils.GetConnectionString())

                ' 変換/解決
                Dim profitCenter As String = If(String.IsNullOrWhiteSpace(profitCenterRaw), Nothing, profitCenterRaw)
                Dim customerUnitIdNullable As Long? = Nothing
                If Not String.IsNullOrWhiteSpace(customerUnitName) Then
                    Dim customerUnitId As Long = repoCust.GetCustomerUnitIdByName(customerUnitName)
                    If customerUnitId <= 0 Then
                        lblError.Text = "注文工場／担当者名が正しく選択されていません。候補から選択してください。"
                        Return
                    End If
                    customerUnitIdNullable = customerUnitId
                End If

                ' 取引先設定IDの取得
                Dim customerSettingIdNullable As Long? = repoCust.FindCustomerSettingIdSimple(
                    customerCode:=customerCode,
                    profitCenter:=profitCenter,
                    customerUnitId:=If(customerUnitIdNullable.HasValue, customerUnitIdNullable.Value, CType(Nothing, Long?))
                )

                If Not customerSettingIdNullable.HasValue Then
                    lblError.Text = "該当する取引先設定が見つかりません。取引先コード／PC／注文工場／担当者名を確認してください。"
                    Return
                End If

                Dim customerSettingId As Long = customerSettingIdNullable.Value

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repoInfoType.ExistsInfoType(
                    customerSettingId:=customerSettingId,
                    customerInfoType:=customerInfoType,
                    excludeInfoTypeId:=0
                )
                If existsOther Then
                    lblError.Text = "同一（取引先コード、PC、注文工場／担当者名、取引先情報区分）の設定が見つかりました。"
                    Return
                End If

                ' INSERT
                Dim activeFlag As String = "Y"
                Dim newId As Long = repoInfoType.InsertInfoTypeNullable(
                    customerSettingId:=customerSettingId,
                    customerInfoType:=customerInfoType,
                    infoType:=infoType,
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
            btnCreateInfoTypeSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region
    End Class
End Namespace