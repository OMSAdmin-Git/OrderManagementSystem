Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.CustomerSetting
Imports System.Data
Imports System.Text

Namespace Pages.Masters.CustomerUnit
    Public Class CustomerUnitList
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補を読み込み
                LoadSearchConditionLists()

                ' 初期表示
                BindCustomerUnitListGrid(activeFlag:="Y")
            End If
        End Sub
#End Region

#Region "検索条件：候補読み込み"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim repo As New CustomerRepository(Utils.GetConnectionString())

            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)
            lstSearchCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "GridView バインド"
        ' 用途ベース命名：CustomerUnit 一覧グリッドにデータを流し込む
        Private Sub BindCustomerUnitListGrid(
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        )
            Dim repo As New CustomerUnitRepository(Utils.GetConnectionString())
            Dim dt As DataTable =
                repo.GetCustomerUnitList(
                    customerUnitName:=customerUnitName,
                    activeFlag:=activeFlag
                )

            gvCustomerUnitList.DataSource = dt
            gvCustomerUnitList.DataBind()
        End Sub
#End Region

#Region "検索系 ボタンイベント"
        ' 検索ボタン
        Protected Sub btnSearchGv_ServerClick(sender As Object, e As EventArgs)
            Dim customerUnitName As String = NullIfWhite(txtSearchCustomerUnitName.Value)
            Dim activeFlag As String = If(chkSearchActiveOnly.Checked, "Y", Nothing)

            BindCustomerUnitListGrid(
                customerUnitName:=customerUnitName,
                activeFlag:=activeFlag
            )
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_ServerClick(sender As Object, e As EventArgs)
            txtSearchCustomerUnitName.Value = ""
            chkSearchActiveOnly.Checked = True

            BindCustomerUnitListGrid(activeFlag:="Y")
        End Sub
#End Region

#Region "画面遷移"
        ' マスタメニューへ
        Protected Sub btnMasterMenu_Click(sender As Object, e As EventArgs)
            Response.Redirect("../MasterMenu.aspx")
        End Sub

        ' 新規登録
        Protected Sub btnCustomerUnitSettingCreate_Click(sender As Object, e As EventArgs)
            Response.Redirect("CustomerUnitSettingCreate.aspx")
        End Sub
#End Region

    End Class
End Namespace