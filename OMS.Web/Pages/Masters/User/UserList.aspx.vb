Imports System.Data
Imports System.Text
Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.CustomerSetting
Imports Oracle.ManagedDataAccess.Client

Public Class UserList
    Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ' ユーザー名表示
            PageHelpers.SetUserName(Me, lblUser)

            '検索条件初期化
            txtSearchUserId.Value = ""
            txtSearchUserName.Value = ""
            chkSearchActiveOnly.Checked = True

            ' 初期表示
            BindUserListGrid(ActiveFlag:="Y")
        End If
    End Sub
#End Region

#Region "画面遷移"
    ' マスタメニューへ
    Protected Sub btnMasterMenu_Click(sender As Object, e As EventArgs)
        Response.Redirect("../MasterMenu.aspx")
    End Sub
#End Region
#Region "GridView バインド"
    ' 用途ベース命名：ユーザーマスタ一覧のグリッドにデータを流し込む
    Private Sub BindUserListGrid(
            Optional ByVal ProdMgmtUserId As String = Nothing,
            Optional ByVal ProdMgmtUserName As String = Nothing,
            Optional ByVal Password As String = Nothing,
            Optional ByVal AuthorityLevel As String = Nothing,
            Optional ByVal ActiveFlag As String = Nothing,
            Optional ByVal CreatedAt As String = Nothing,
            Optional ByVal CreatedUserId As String = Nothing,
            Optional ByVal CreatedPgId As String = Nothing,
            Optional ByVal UpdatedAt As String = Nothing,
            Optional ByVal UpdatedUserId As String = Nothing,
            Optional ByVal UpdatedPgId As String = Nothing
        )
        Dim repo As New UserRepository(Utils.GetConnectionString())
        Dim dt As DataTable =
                repo.GetUserList(
                    prodMgmtUserId:=ProdMgmtUserId,
                    ProdMgmtUserName:=ProdMgmtUserName,
                    activeFlag:=ActiveFlag
                    )

        gvUserList.DataSource = dt
        gvUserList.DataBind()
    End Sub
#End Region

#Region "ボタンイベント"
    ' 検索ボタン
    Protected Sub btnSearchGv_ServerClick(sender As Object, e As EventArgs)
        Dim userId As String = NullIfWhite(txtSearchUserId.Value)
        Dim userName As String = NullIfWhite(txtSearchUserName.Value)
        Dim activeFlag As String = If(chkSearchActiveOnly.Checked, "Y", Nothing)

        BindUserListGrid(
                ProdMgmtUserId:=userId,
                ProdMgmtUserName:=userName,
                ActiveFlag:=activeFlag
            )
    End Sub

    ' クリアボタン
    Protected Sub btnDefaultGv_ServerClick(sender As Object, e As EventArgs)
        txtSearchUserId.Value = ""
        txtSearchUserName.Value = ""
        chkSearchActiveOnly.Checked = True

        BindUserListGrid(ActiveFlag:="Y")
    End Sub

    ' パスワードリセットボタン
    Protected Sub gvUserList_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        Dim repo As New UserRepository(Utils.GetConnectionString())
        Dim loginUserId As String = PageHelpers.GetUserId(Me)
        Dim pgId As String = "SaveUserSetting(Reset)"

        If e.CommandName = "PassReset" Then
            Dim prodMgmtUserId As String = e.CommandArgument.ToString()

            Try
                repo.UpdatePasswordReset(
                        prodMgmtUserId,
                        loginUserId,
                        pgId:=pgId
                    )

                lblResult.Text = "パスワードを初期化しました。"
                lblError.Text = ""
                Dim userId As String = NullIfWhite(txtSearchUserId.Value)
                Dim userName As String = NullIfWhite(txtSearchUserName.Value)
                Dim activeFlag As String = If(chkSearchActiveOnly.Checked, "Y", Nothing)

                BindUserListGrid(
                        ProdMgmtUserId:=userId,
                        ProdMgmtUserName:=userName,
                        ActiveFlag:=activeFlag
                    )
            Catch ex As OracleException
                lblResult.Text = ""
                lblError.Text = "システムエラーが発生しました。時間をおいて再度お試しください。"
            Catch ex As Exception
                lblResult.Text = ""
                lblError.Text = "システムエラーが発生しました。"
            End Try


        End If
    End Sub
#End Region

End Class