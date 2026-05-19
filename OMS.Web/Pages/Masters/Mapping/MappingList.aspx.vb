Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.Folder
Imports System.Data
Imports System.Text

Namespace Pages.Masters.Mapping
    Public Class MappingList
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                PageHelpers.SetUserName(Me, lblUser)
                LoadSearchConditionLists()
                BindMappingListGrid(activeFlag:="Y")
            End If
        End Sub
#End Region

#Region "検索条件：候補読み込み"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim repo As New CustomerRepository(Utils.GetConnectionString())

            Dim customerCodeList As List(Of String) = repo.GetCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = repo.GetProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

            lstSearchCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstSearchProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstSearchCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "GridView バインド"
        Private Sub BindMappingListGrid(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        )
            Dim repo As New MappingRepository(Utils.GetConnectionString())
            Dim dt As DataTable =
                repo.GetMappingList(
                    customerCode:=customerCode,
                    customerName:=customerName,
                    profitCenter:=profitCenter,
                    customerUnitName:=customerUnitName,
                    prodMgmtUserId:=PageHelpers.GetUserId(Me),
                    activeFlag:=activeFlag
                )

            gvMappingList.DataSource = dt
            gvMappingList.DataBind()
        End Sub
#End Region

#Region "検索系 ボタンイベント"
        ' 検索ボタン
        Protected Sub btnSearchGv_ServerClick(sender As Object, e As EventArgs)
            Dim customerCode As String = NullIfWhite(txtSearchCustomerCode.Value)
            Dim customerName As String = NullIfWhite(txtSearchCustomerName.Value)
            Dim profitCenter As String = NullIfWhite(txtSearchProfitCenter.Value)
            Dim customerUnitName As String = NullIfWhite(txtSearchCustomerUnitName.Value)
            Dim activeFlag As String = If(chkSearchActiveOnly.Checked, "Y", Nothing)

            BindMappingListGrid(
                customerCode,
                customerName,
                profitCenter,
                customerUnitName,
                activeFlag
            )
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_ServerClick(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""
            chkSearchActiveOnly.Checked = True

            BindMappingListGrid(activeFlag:="Y")
        End Sub
#End Region

#Region "画面遷移"
        ' マスタメニューへ
        Protected Sub btnMasterMenu_Click(sender As Object, e As EventArgs)
            Response.Redirect("../MasterMenu.aspx")
        End Sub

        ' 新規登録
        Protected Sub btnMappingSettingCreate_Click(sender As Object, e As EventArgs)
            Response.Redirect("MappingSettingCreate.aspx")
        End Sub
#End Region

    End Class
End Namespace