Imports OMS.Common

Namespace Pages.Masters
    Public Class MasterMenu
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
                ' ログインユーザーの権限に応じてメニューの表示/非表示を切り替える（例：管理者のみユーザーマスタを表示）
                Dim authority As String = TryCast(Page.Session("UserAuthority"), String)
                If authority = "1" Then
                    btnUserList.Visible = True
                Else
                    btnUserList.Visible = False
                End If
            End If
        End Sub
#End Region

#Region "メニュー遷移（ナビゲーション）"
        ' トップメニューへ
        Protected Sub btnTopMenu_Click(sender As Object, e As EventArgs)
            Navigate("../TopMenu.aspx")
        End Sub

        ' 顧客一覧
        Protected Sub btnCustomerList_ServerClick(sender As Object, e As EventArgs)
            Navigate("CustomerSetting/CustomerList.aspx")
        End Sub

        ' 顧客ユニット一覧
        Protected Sub btnCustomerUnitList_ServerClick(sender As Object, e As EventArgs)
            Navigate("CustomerUnit/CustomerUnitList.aspx")
        End Sub

        ' フォルダ一覧
        Protected Sub btnFolderList_ServerClick(sender As Object, e As EventArgs)
            Navigate("Folder/FolderList.aspx")
        End Sub

        ' ファイル一覧
        Protected Sub btnFileList_ServerClick(sender As Object, e As EventArgs)
            Navigate("File/FileList.aspx")
        End Sub

        ' 情報区分一覧
        Protected Sub btnInfoTypeList_ServerClick(sender As Object, e As EventArgs)
            Navigate("InfoType/InfoTypeList.aspx")
        End Sub

        ' マッピング一覧
        Protected Sub btnMappingList_ServerClick(sender As Object, e As EventArgs)
            Navigate("Mapping/MappingList.aspx")
        End Sub

        ' 取込ルール一覧
        Protected Sub btnImpRuleList_ServerClick(sender As Object, e As EventArgs)
            Navigate("ImpRule/ImpRuleList.aspx")
        End Sub

        ' 生産計画ルール一覧
        Protected Sub btnProdPlanRuleList_ServerClick(sender As Object, e As EventArgs)
            Navigate("ProdPlanRule/ProdPlanRuleList.aspx")
        End Sub
        ' ユーザー一覧
        Protected Sub btnUserList_ServerClick(sender As Object, e As EventArgs)
            Navigate("User/UserList.aspx")
        End Sub
#End Region

#Region "ナビゲーション共通"
        ' 相対パスでの遷移を共通化（将来ログや権限チェックを入れる場合もここで一元化）
        Private Sub Navigate(relativeUrl As String)
            Response.Redirect(relativeUrl)
        End Sub
#End Region

    End Class
End Namespace