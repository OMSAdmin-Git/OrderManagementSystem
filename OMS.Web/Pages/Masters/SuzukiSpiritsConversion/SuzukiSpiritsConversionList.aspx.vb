Imports System.Data
Imports System.Text
Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.File
Imports OMS.Web.Pages.Masters.SuzukiSpiritsConversion

Namespace Pages.Masters.SuzukiSpiritsConversion
    Public Class SuzukiSpiritsConversionList
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
                ' 初期表示
                BindSuzukiSpiritsConversionListGrid(activeFlag:="Y")
            End If
        End Sub
#End Region

#Region "GridView バインド"
        ' 用途ベース命名：SuzukiSpiritsConversion 一覧グリッドにデータを流し込む
        Private Sub BindSuzukiSpiritsConversionListGrid(
            Optional ByVal deliveryCodePlan As String = Nothing,
            Optional ByVal deliveryCodeOrder As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        )
            Dim repo As New SuzukiSpiritsConversionRepository(Utils.GetConnectionString())
            Dim dt As DataTable =
                repo.GetSuzukiSpiritsConversionList(
                    deliveryCodePlan:=deliveryCodePlan,
                    deliveryCodeOrder:=deliveryCodeOrder,
                    activeFlag:=activeFlag
                )

            gvSuzukiSpiritsConversionList.DataSource = dt
            gvSuzukiSpiritsConversionList.DataBind()
        End Sub
#End Region

#Region "検索系 ボタンイベント"
        ' 検索ボタン
        Protected Sub btnSearchGv_ServerClick(sender As Object, e As EventArgs)
            Dim deliveryCodeOrder As String = NullIfWhite(txtSearchDeliveryCodeOrder.Value)
            Dim deliveryCodePlan As String = NullIfWhite(txtSearchDeliveryCodePlan.Value)
            Dim activeFlag As String = If(chkSearchActiveOnly.Checked, "Y", Nothing)

            BindSuzukiSpiritsConversionListGrid(
                deliveryCodePlan:=deliveryCodePlan,
                deliveryCodeOrder:=deliveryCodeOrder,
                activeFlag:=activeFlag
            )
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_ServerClick(sender As Object, e As EventArgs)
            txtSearchDeliveryCodeOrder.Value = ""
            txtSearchDeliveryCodePlan.Value = ""
            chkSearchActiveOnly.Checked = True

            BindSuzukiSpiritsConversionListGrid(activeFlag:="Y")
        End Sub
#End Region

#Region "画面遷移"
        ' マスタメニューへ
        Protected Sub btnMasterMenu_Click(sender As Object, e As EventArgs)
            Response.Redirect("../MasterMenu.aspx")
        End Sub

        ' 新規登録
        Protected Sub btnSuzukiSpiritsConversionSettingCreate_Click(sender As Object, e As EventArgs)
            Response.Redirect("SuzukiSpiritsConversionSettingCreate.aspx")
        End Sub

#End Region
    End Class
End Namespace