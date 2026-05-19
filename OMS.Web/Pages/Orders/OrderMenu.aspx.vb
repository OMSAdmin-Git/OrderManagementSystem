Imports OMS.Common

Namespace Pages.Orders
    Public Class OrderMenu
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
            End If
        End Sub
#End Region

#Region "メニュー遷移（ナビゲーション）"
        ' トップへ
        Protected Sub btnTopMenu_Click(sender As Object, e As EventArgs)
            Navigate("../TopMenu.aspx")
        End Sub

        ' 受注登録メニュー
        Protected Sub btnOrderImportStage_ServerClick(sender As Object, e As EventArgs)
            Navigate("OrderImportStage.aspx")
        End Sub

        Protected Sub btnOrderImport_ServerClick(sender As Object, e As EventArgs)
            Navigate("OrderImport.aspx")
        End Sub

        Protected Sub btnDueSetFromImport_ServerClick(sender As Object, e As EventArgs)
            Navigate("DueDateSettingFromImport.aspx")
        End Sub

        ' 生産計画メニュー
        Protected Sub btnProdPlan_ServerClick(sender As Object, e As EventArgs)
            Navigate("ProdPlan.aspx")
        End Sub

        Protected Sub btnDueSetFromProdPlan_ServerClick(sender As Object, e As EventArgs)
            Navigate("DueDateSettingFromProdPlan.aspx")
        End Sub

        ' 受注出力メニュー
        Protected Sub btnOrderExport_ServerClick(sender As Object, e As EventArgs)
            Navigate("OrderExport.aspx")
        End Sub

        ' 社内調整
        Protected Sub btnSelfFcst_ServerClick(sender As Object, e As EventArgs)
            Navigate("SelfFcst.aspx")
        End Sub
#End Region

#Region "ナビゲーション共通"
        ' 相対パスでの遷移を共通化（将来：ログ記録・権限チェック等もここで一元化可能）
        Private Sub Navigate(relativeUrl As String)
            Response.Redirect(relativeUrl)
        End Sub
#End Region

    End Class
End Namespace