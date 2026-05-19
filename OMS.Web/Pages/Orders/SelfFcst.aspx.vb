Imports System
Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.Http
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports OMS.Common
Imports OMS.Data

Namespace Pages.Orders
    Public Class SelfFcst
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補をページ内メソッドで初期化（Utils.BuildOptions を使用）
                LoadSearchConditionLists()

                ' 初期表示（一覧）
                BindSelfFcstGrid()
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

            lstSearchCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstSearchProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstSearchCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "ナビゲーション / 検索ボタン"
        ' 受注メニューボタン
        Protected Sub btnOrderMenu_Click(sender As Object, e As EventArgs)
            Response.Redirect("OrderMenu.aspx")
        End Sub

        ' 検索ボタン
        Protected Sub btnSearchGv_Click(sender As Object, e As EventArgs)
            Dim customerCode As String = NullIfWhite(txtSearchCustomerCode.Value)
            Dim customerName As String = NullIfWhite(txtSearchCustomerName.Value)
            Dim profitCenter As String = NullIfWhite(txtSearchProfitCenter.Value)
            Dim customerUnitName As String = NullIfWhite(txtSearchCustomerUnitName.Value)

            BindSelfFcstGrid(customerCode, customerName, profitCenter, customerUnitName)
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_Click(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""

            BindSelfFcstGrid()
        End Sub
#End Region

#Region "GridViewバインド／イベント"
        ' ステージ一覧をバインド
        Private Sub BindSelfFcstGrid(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing
        )
            Dim repo As New OrderHistoryRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetOrdersHistory(
                customerCode:=customerCode,
                customerName:=customerName,
                profitCenter:=profitCenter,
                selfFcstFlag:="Y",
                customerUnitName:=customerUnitName,
                prodMgmtUserId:=PageHelpers.GetUserId(Me),
                status:="DUE_SET",
                activeFlag:="Y"
            )

            gvSelfFcst.DataSource = dt
            gvSelfFcst.DataBind()
        End Sub

        ' 行データバウンド（ヘッダーの一括チェック連動）
        Protected Sub gvSelfFcst_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvSelfFcst.RowDataBound
            If e.Row.RowType = DataControlRowType.DataRow Then
                ' 処理対象列
                Dim chkImport As CheckBox = TryCast(e.Row.FindControl("chkSelfFcst"), CheckBox)
                If chkImport IsNot Nothing Then
                    chkImport.InputAttributes("onclick") =
                        $"OMS.Grid.updateHeader('{gvSelfFcst.ClientID}', 'chkSelfFcstAll', 'chkSelfFcst');"
                End If
            End If
        End Sub
#End Region

#Region "ボタンイベント"
        ' 削除ボタン
        Protected Sub btnSelfFcstDelete_Click(sender As Object, e As EventArgs)
            lblResult.Text = ""
            lblError.Text = ""

            Try
                ' GridViewからチェック済みの OrderId を収集
                Dim orderIds As New List(Of Long)()
                For Each row As GridViewRow In gvSelfFcst.Rows
                    If row.RowType <> DataControlRowType.DataRow Then Continue For
                    Dim chk As CheckBox = TryCast(row.FindControl("chkSelfFcst"), CheckBox)
                    If chk IsNot Nothing AndAlso chk.Checked Then
                        Dim keyObj As Object = gvSelfFcst.DataKeys(row.RowIndex).Value ' DataKeyNames="OrderId" 前提
                        If keyObj IsNot Nothing AndAlso keyObj IsNot DBNull.Value Then
                            Dim oid As Long
                            If Long.TryParse(keyObj.ToString(), oid) Then
                                orderIds.Add(oid)
                            End If
                        End If
                    End If
                Next

                If orderIds.Count = 0 Then
                    lblError.Text = "処理対象が選択されていません。"
                    Return
                End If

                ' 更新
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim programId As String = "SelfFcst(Delete)"
                Dim newActiveFlag As String = "N"

                ' 一括更新（生産計画以前のレコード）
                Dim repoO As New OrderRepository(Utils.GetConnectionString())
                Dim affected As Integer = repoO.UpdateOrdersActiveFlagByOrderIds(
                    orderIds:=orderIds,
                    newActiveFlag:=newActiveFlag,
                    loginUserId:=loginUserId,
                    programId:=programId
                )

                '' 一括更新（生産計画以降のレコード）
                'Dim repoH As New OrderHistoryRepository(Utils.GetConnectionString())
                'Dim affected As Integer = repoH.UpdateOrdersActiveFlagByOrderIds(
                '    orderIds:=orderIds,
                '    newActiveFlag:=newActiveFlag,
                '    loginUserId:=loginUserId,
                '    programId:=programId
                ')

                lblResult.Text = $"自社予測データ を {affected} 件削除しました。"

                ' 再描画
                BindSelfFcstGrid()

            Catch ex As Oracle.ManagedDataAccess.Client.OracleException
                lblError.Text = "DBエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)
            Catch ex As Exception
                lblError.Text = "予期しないエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)
            End Try
        End Sub
#End Region

    End Class
End Namespace