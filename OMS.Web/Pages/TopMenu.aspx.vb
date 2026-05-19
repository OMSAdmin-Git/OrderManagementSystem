Imports OMS.Common
Imports OMS.Data

Namespace Pages
    Partial Public Class TopMenu
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
            End If
        End Sub
#End Region

#Region "メニュー遷移（ナビゲーション）"
        Protected Sub btnOrder_ServerClick(sender As Object, e As EventArgs)
            Navigate("~/Pages/Orders/OrderMenu.aspx")
        End Sub

        Protected Sub btnMaster_ServerClick(sender As Object, e As EventArgs)
            Navigate("~/Pages/Masters/MasterMenu.aspx")
        End Sub

        Protected Sub btnLogout_Click(sender As Object, e As EventArgs)
            ' セッション破棄 → ログインへ
            Session.Clear()
            Session.Abandon()
            Navigate("~/Pages/Login/Login.aspx")
        End Sub
#End Region

#Region "ナビゲーション共通"
        ' 画面遷移の共通化（将来的にログ記録や権限チェックを入れる場合にもここで一元管理）
        Private Sub Navigate(relativeUrl As String)
            Response.Redirect(relativeUrl)
        End Sub
#End Region

    End Class
End Namespace