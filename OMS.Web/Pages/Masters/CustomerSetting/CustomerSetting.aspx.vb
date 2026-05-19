Imports System.Data
Imports System.Text
Imports System.Web.Script.Services
Imports System.Web.Services
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client

Namespace Pages.Masters.CustomerSetting
    Public Class CustomerSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        ' id1 = CustomerSettingId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property CustomerSettingId As Long
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
                LoadSearchConditionLists()
                LoadCustomerHeader(CustomerSettingId)
            End If
        End Sub
#End Region

#Region "検索候補：リスト初期化"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim custRepo As New CustomerRepository(Utils.GetConnectionString())
            Dim userRepo As New UserRepository(Utils.GetConnectionString())

            Dim customerCodeList As List(Of String) = custRepo.GetStraCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = custRepo.GetStraProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = custRepo.GetCustomerUnitNames(loginUserId)
            Dim prodMgmtUserName As List(Of String) = userRepo.GetUserNames()

            lstCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
            lstProdMgmtUserName.InnerHtml = BuildOptions(prodMgmtUserName)
        End Sub
#End Region

#Region "取引先設定マスタデータ"
        Private Sub LoadCustomerHeader(customerSettingId As Long)

            If customerSettingId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New CustomerRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetCustomerSetting(customerSettingId)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                SetHeaderControls(Nothing)
                lblError.Text = "選択されたデータが見つかりません。"
                Exit Sub
            End If

            SetHeaderControls(dt.Rows(0))
        End Sub

        ' 選択データセット
        Private Sub SetHeaderControls(r As DataRow)
            txtCustomerCode.Value = If(GetStr(r, "CustomerCode"), "")
            txtCustomerName.Text = If(GetStr(r, "CustomerName"), "")
            txtProfitCenter.Value = If(GetStr(r, "ProfitCenter"), "")
            txtCustomerUnitName.Value = If(GetStr(r, "CustomerUnitName"), "")
            txtProdMgmtUserName.Value = If(GetStr(r, "UserName"), "")
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
        Protected Sub btnSaveCustomerSetting_Click(sender As Object, e As EventArgs)

            lblError.Text = ""
            lblResult.Text = ""

            Try

                If CustomerSettingId <= 0 Then
                    lblError.Text = "対象IDが不正です。再度ページを開き直してください。"
                    Return
                End If

                ' 入力値取得
                Dim customerCode As String = (If(txtCustomerCode.Value, "")).Trim()
                Dim profitCenterRaw As String = (If(txtProfitCenter.Value, "")).Trim()
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim prodMgmtUserName As String = (If(txtProdMgmtUserName.Value, "")).Trim()
                Dim activeFlag As String = If(ddlActiveFlag.SelectedValue, "").Trim().ToUpperInvariant()

                ' ログイン情報
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim pgId As String = "CustomerSetting(Update)"

                ' 入力確認
                ' JSで処理しているためなし。必要なら追加。

                Dim repo As New CustomerRepository(Utils.GetConnectionString())

                If Not repo.ExistsCustomerCode(customerCode) Then
                    lblError.Text = "取引先コードが正しく選択されていません。候補から選択してください。"
                    Return
                End If

                Dim profitCenter As String = If(String.IsNullOrWhiteSpace(profitCenterRaw), Nothing, profitCenterRaw)
                If Not String.IsNullOrWhiteSpace(profitCenter) Then
                    If Not repo.ExistsProfitCenter(profitCenter) Then
                        lblError.Text = "PCが正しく選択されていません。候補から選択してください。"
                        Return
                    End If
                End If

                Dim customerUnitIdNullable As Long? = Nothing
                If Not String.IsNullOrWhiteSpace(customerUnitName) Then
                    Dim customerUnitId As Long = repo.GetCustomerUnitIdByName(customerUnitName) ' 未存在/複数なら ApplicationException を投げる契約なら OK
                    If customerUnitId > 0 Then
                        customerUnitIdNullable = customerUnitId
                    Else
                        lblError.Text = "注文工場／担当者名が正しく選択されていません。候補から選択してください。"
                        Return
                    End If
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
                    excludeCustomerSettingId:=CustomerSettingId
                )
                If existsOther Then
                    lblError.Text = "同一（取引先コード、PC、注文工場／担当者名）の設定が見つかりました。"
                    Return
                End If

                ' UPDATE
                Dim affected As Integer = repo.UpdateCustomerSettingWithConcurrency(
                    customerSettingId:=CustomerSettingId,
                    customerCode:=customerCode,
                    profitCenter:=profitCenter,
                    customerUnitId:=customerUnitIdNullable,
                    prodMgmtUserId:=prodMgmtUserId,
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
                LoadCustomerHeader(CustomerSettingId)

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
            ' lblError の ClientID が確定したタイミングで data-error-label-id を付与
            btnSaveCustomerSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region
    End Class
End Namespace