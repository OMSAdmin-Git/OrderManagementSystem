Imports System.Formats.Asn1
Imports System.IO
Imports DocumentFormat.OpenXml.VariantTypes
Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters.File
Imports Oracle.ManagedDataAccess.Client

Namespace Pages.Orders
    Public Class OrderExport
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補を初期化
                LoadSearchConditionLists()

                ' データバインド
                'gvSelectCustomers_Init()
                ' 初期表示（一覧）
                BindSelectCustomersGrid()
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
            'gvSelectCustomers_Init(customerCode, customerName, profitCenter, customerUnitName)
            BindSelectCustomersGrid(customerCode, customerName, profitCenter, customerUnitName)
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_Click(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""
            'gvSelectCustomers_Init()
            BindSelectCustomersGrid()
        End Sub
#End Region

#Region "GridView バインド / イベント"
        '' GridViewデータバインド
        'Private Sub gvSelectCustomers_Init(
        '    Optional ByVal customerCode As String = Nothing,
        '    Optional ByVal customerName As String = Nothing,
        '    Optional ByVal profitCenter As String = Nothing,
        '    Optional ByVal customerUnitName As String = Nothing
        ')
        '    Dim repo As New CustomerRepository(Utils.GetConnectionString())
        '    Dim dt As DataTable = repo.GetCustomerList(
        '                                customerCode:=customerCode,
        '                                customerName:=customerName,
        '                                profitCenter:=profitCenter,
        '                                customerUnitName:=customerUnitName,
        '                                prodMgmtUserId:=PageHelpers.GetUserId(Me))
        '    gvSelectCustomers.DataSource = dt
        '    gvSelectCustomers.DataBind()
        'End Sub
        Private Sub BindSelectCustomersGrid(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing
        )
            Dim repo As New CustomerRepository(Utils.GetConnectionString())
            Dim dt = repo.GetCustomerList(
                customerCode:=customerCode,
                customerName:=customerName,
                profitCenter:=profitCenter,
                customerUnitName:=customerUnitName,
                prodMgmtUserId:=PageHelpers.GetUserId(Me)
            )

            gvSelectCustomers.DataSource = dt
            gvSelectCustomers.DataBind()
        End Sub

        ' GridViewヘッダーバインド
        Protected Sub gvSelectCustomers_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvSelectCustomers.RowDataBound
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox = TryCast(e.Row.FindControl("chkOrderExport"), CheckBox)
                If chk IsNot Nothing Then
                    ' 個別のチェック操作時にヘッダー状態を更新
                    chk.InputAttributes("onclick") =
                        $"OMS.Grid.updateHeader('{gvSelectCustomers.ClientID}', 'chkOrderExportAll', 'chkOrderExport');"
                End If
            End If
        End Sub
#End Region

#Region "CSV出力 / エラー出力"
        ' CSV出力ボタン
        Protected Sub btnOrderExport_Click(sender As Object, e As EventArgs)
            lblResult.Text = ""
            lblError.Text = ""
            'lblError.Text = "開発未着手"
            Dim fileList As List(Of String) = New List(Of String)()
            Dim strPath = Server.MapPath("~/App_Data/Files/")
            Dim FileDate = DateTime.Now
            Dim unofficial = 0
            Dim confirmed = 0
            Dim count = 0
            Dim valid = 0
            Dim errors As New List(Of String)()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)

            ' 値取得
            Dim ProcessingStartDate As Date = DateTime.Now

            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(Utils.GetConnectionString())
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()
            ' Table class access 
            ' 受注
            Dim repo As New OrderRepository(Utils.GetConnectionString())
            ' 受注ワーク
            Dim reps As New OrderStageRepository(Utils.GetConnectionString())
            ' 受注履歴
            Dim reph = New OrderHistoryRepository(Utils.GetConnectionString())
            ' 生産計画条件
            Dim prodp = New ProdPlanRuleRepository(Utils.GetConnectionString())
            ' 分割条件
            Dim splitp = New SplitCaseRepository(Utils.GetConnectionString())
            ' 処理対象となる CustomerSettingId リスト
            Dim idList As New List(Of Long)

            Try
                '値取得
                '[処理開始日時]を取得する。

                'DELETE
                '(生産計画ワーク)
                ' ①	STATUS（ステータス） = 'POST_PLAN_DUE_SET' 	
                errors.Add(reps.Delete(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, status:="POST_PLAN_DUE_SET"))
                If (CheckError(errors)) Then
                    ' エラー
                End If
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                ' 処理対象 がチェックされている行
                For Each row In gvSelectCustomers.Rows

                    Dim chk = TryCast(row.FindControl("chkOrderExport"), WebControls.CheckBox)

                    If chk IsNot Nothing Then
                        ' チェックがある時
                        If chk.Checked Then

                            Dim idx As Integer = row.RowIndex
                            Dim keys = gvSelectCustomers.DataKeys(idx)

                            ' 取引先設定ID 取得＆安全化
                            'Dim profitCenter As String = gvSelectCustomers

                            Dim customerSettingId As Long
                            Dim csidObj = keys("CustomerSettingId")
                            If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), customerSettingId) Then
                                errors.Add($"Row {idx}：CustomerSettingIdが不正")
                                Continue For
                            End If
                            csidObj = keys("CustomerCode")
                            Dim customerCode As String = csidObj.ToString()
                            csidObj = keys("ProfitCenter")
                            Dim profitCenter As String = csidObj.ToString()
                            Dim updatedAt = DateTime.Now
                            ' Customer Setting ID リストに追加
                            idList.Add(customerSettingId)

                            ' 生産計画ワーク追加
                            'PROD_PLAN（生産計画テーブル）から以下全ての条件を満たすレコードをPROD_PLAN_STAGE（生産計画ワークテーブル）へ追加する。	
                            '①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 
                            '②	STATUS（ステータス） = 'POST_PLAN_DUE_SET'
                            '③	ACTIVE_FLAG（有効フラグ）= 'Y'
                            Dim orderStageRowsOrg = repo.ToClass(repo.GetOrders(conn, tran, OrderRepository.OrdersTable.ProductPlan, status:="POST_PLAN_DUE_SET", activeFlag:="Y", customerSettingId:=customerSettingId))
                            '-----------------------------
                            ' INSERT 生産計画ワークテーブル
                            '-----------------------------
                            errors.Add(reps.InsertRange(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderStageRowsOrg))
                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            '以下項目をキーとしてPROD_PLAN_HISTORY（生産計画履歴）とPROD_PLAN_STAGE（生産計画ワークテーブル）を比較し、
                            '該当するPROD_PLAN_STAGEのレコードを更新する。

                            ' 出荷状況チェック
                            '[PROD_PLAN_HISTORY]
                            'CUSTOMER_SETTING_ID(取引先設定ID)  処理中のCUSTOMER_SETTING_ID
                            'STATUS(ステータス)                 'EXPORTED'
                            'ACTIVE_FLAG(有効フラグ)             'Y'
                            Dim ordersh = reph.GetOrders(conn, tran, OrderHistoryRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId, status:="EXPORTED", activeFlag:="Y")
                            '[PROD_PLAN_STAGE]
                            '書式・条件（PROD_PLAN_STAGE）処理中のCUSTOMER_SETTING_ID
                            'STATUS(ステータス)                 'POST_PLAN_DUE_SET'
                            'CTIVE_FLAG(有効フラグ)             'Y'
                            Dim orderss = reps.GetOrders(conn, tran, OrderHistoryRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId, status:="POST_PLAN_DUE_SET", activeFlag:="Y")

                            '一致比較
                            'CUSTOMER_SETTING_ID（取引先設定ID）
                            'CUSTOMER_ORDER_NO(客先発注No)
                            'ITEM_NO（品目No）
                            Dim updatedRows = reps.ToClass((From rowT In orderss.AsEnumerable()
                                                            Join rowH In ordersh.AsEnumerable()
                                                           On rowT.Field(Of Long)("customer_setting_id") Equals rowH.Field(Of Long)("customer_setting_id") And
                                                                  rowT.Field(Of String)("customer_order_no") Equals rowH.Field(Of String)("customer_order_no") And
                                                                  rowT.Field(Of String)("item_no") Equals rowH.Field(Of String)("item_no")
                                                            Select rowT).
                                                            GroupBy(Function(r) r.Field(Of Long)("stage_id")).
                                                            Select(Function(g) g.First()))
                            'ACTIVE_FLAG(有効フラグ)            'N'
                            'UPDATED_AT(更新日時)              [処理開始日時]
                            'UPDATED_USER_ID(更新ユーザーID)        [ログインユーザーID]
                            'UPDATED_PG_ID(更新プログラムID)      'OrderExport'
                            errors.Add(reps.UpdateByOrderId(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, updatedRows, activeFlag:="N", updatedPgId:="OrderExport", updatedUserId:=loginUserId, updatedAt:=updatedAt))

                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                        End If
                    End If
                Next

                '出荷状況エラーリスト出力	
                'PROD_PLAN_STAGE_VIEW（生産計画出力一覧）をExcel出力する。	
                Dim repos = New OrderStraRepository(Utils.GetConnectionString())
                Dim ErrorRows = repos.GetOrderStage(conn, tran, status:="POST_PLAN_DUE_SET", activeFlag:="N")
                errors.Add(OrderProductionPlanExcelFile.ShippingStatusErrorExcelOut(strPath, FileDate, repos.ToClass(ErrorRows)))
                If (CheckError(errors)) Then
                    ' エラー
                End If
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                Dim trfilename = OrderProductionPlanExcelFile.GetErrorListExcelFilename(strPath, FileDate)
                'Utils.FileTransfer(Response, Server, trfilename)
                fileList.Add(trfilename)

                'CSV出力(内示)
                'CSVファイルへPROD_PLAN_STRA_VIEW（生産計画出力一覧）を書き出し、ブラウザで設定されているダウンロードフォルダへ出力する。
                'Dim PlanRows = repos.GetOrderStras(conn, tran, demandStatus:="F", status:="POST_PLAN_DUE_SET", activeFlag:="Y")
                Dim formatEx As New List(Of (name As String, format As String)) From {("SHIP_PLAN_DATE", "yyyyMMdd")}
                Dim sql = " SELECT * 
                            FROM prod_plan_stra_view 
                            WHERE demand_status = 'F' "
                Dim delimiter = DELIMITER_COMMA
                Dim enclosure = QUOTE_NONE
                Dim headerYN = "N"
                Dim lineEnding = NEWLINE_CRLF
                Dim charset = ENCODING_UTF8_BOM
                Dim processDate = FileDate
                Dim fileBaseName = "A-R-COEDI-F_"
                'ExportQueryToCsvResponse2(sql, delimiter, enclosure, headerYN, lineEnding, charset, processDate, fileBaseName:=fileBaseName, formatEx:=formatEx)

                trfilename = repo.GetProdPlanStraViewCsvFilename(strPath, fileBaseName, FileDate)
                errors.Add(repo.ProdPlanStraViewCsvFile(sql, delimiter, enclosure, headerYN, lineEnding, charset, processDate, filename:=trfilename, formatEx:=formatEx))
                fileList.Add(trfilename)

                'CSV出力(確定/納入指示)
                sql = " SELECT * 
                            FROM prod_plan_stra_view 
                            WHERE  demand_status = 'O' "
                fileBaseName = "A-R-COEDI-O_"
                'ExportQueryToCsvResponse2(sql, delimiter, enclosure, headerYN, lineEnding, charset, processDate, fileBaseName:=fileBaseName, formatEx:=formatEx)

                trfilename = repo.GetProdPlanStraViewCsvFilename(strPath, fileBaseName, FileDate)
                errors.Add(repo.ProdPlanStraViewCsvFile(sql, delimiter, enclosure, headerYN, lineEnding, charset, processDate, filename:=trfilename, formatEx:=formatEx))
                fileList.Add(trfilename)

                'UPDATE
                '生産計画ワーク
                '[抽出条件]
                'status POST_PLAN_DUE_SET
                'activeFlag Y
                '[更新]
                'STATUS(ステータス)              EXPORTED
                'UPDATED_AT(更新日時)            処理開始日時
                'UPDATED_USER_ID(更新ユーザーID) ログインユーザーID
                'UPDATED_PG_ID(更新プログラムID) OrderExport
                'errors.Add(reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, kStatus:="POST_PLAN_DUE_SET", kActiveFlag:="Y", status:="EXPORTED", updatedAt:=FileDate, updatedUserId:=loginUserId, updatedPgId:="OrderExport"))
                'If (CheckError(errors)) Then
                '    ' エラー
                '    DBError(tran)
                'End If

                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If

                Dim rowsu = reps.ToClass(reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, status:="POST_PLAN_DUE_SET", activeFlag:="Y"))


                For Each row In rowsu
                    errors.Add(reps.Update(conn, tran, OrderRepository.OrdersTable.ProductPlan, kOrderId:=row.OrderId, status:="EXPORTED", updatedAt:=FileDate, updatedUserId:=loginUserId, updatedPgId:="OrderExport"))
                Next
                If (CheckError(errors)) Then
                    ' エラー
                    DBError(tran)
                End If
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                'UPDATE
                '生産計画

                'PROD_PLAN_STAGE
                'STATUS(ステータス)
                'UPDATED_AT(更新日時)
                'UPDATED_USER_ID(更新ユーザーID)
                'UPDATED_PG_ID(更新プログラムID)
                Dim rows = reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, status:="EXPORTED", activeFlag:="Y")
                Dim orderRows = reps.ToClass(rows)
                count += orderRows.Count
                For Each row In orderRows
                    errors.Add(repo.Update(conn, tran, OrderRepository.OrdersTable.ProductPlan, kOrderId:=row.OrderId, status:=row.Status, updatedAt:=row.UpdatedAt, updatedUserId:=row.UpdatedUserId, updatedPgId:=row.UpdatedPgId))
                Next
                If (CheckError(errors)) Then
                    ' エラー
                    DBError(tran)
                End If
                valid += orderRows.Count
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                ' 作成したファイルの転送 (Zip)
                '受注ファイル出力_出力日時(yyyyMMddhhmmss).zip
                Dim orderFilename = repo.GeOrderZipFilename("受注ファイル出力_出力日時", FileDate)
                Utils.FilesTransfer(Response, Server, fileList, orderFilename)

                '完了メッセージ表示
                'lblResult.Text = "ファイル出力完了しました。"
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If

            Catch ex As Exception
                Dim m = ex.Message
                errors.Add(ex.Message)
                'lblError.Text = ex.Message
            Finally
                ' 取引先単位で 確定/取り消し が必要なので ループ内で 行う必要がある
                ' Commit/Rollback  Transaction
                If errors.Count = 0 Then
                    tran.Commit()
                Else
                    If (errors.Count > 0) Then
                        'lblError.Text = errors(0)
                        lblError.Text = String.Join(vbCrLf, errors)
                    End If
                    tran.Rollback()
                End If
                If (count = 0 And valid = 0) Then
                    lblResult.Text = $"有効なデータが無かったため受注データの出力は行われませんでした。"
                Else
                    lblResult.Text = $"{count}件中{valid}件の受注データの出力を行いました。"
                End If
                tran.Dispose()
                conn.Close()
                conn.Dispose()
            End Try

        End Sub
        ''' <summary>
        ''' Error メッセージ以外を削除する エラー数を返す
        ''' </summary>
        ''' <param name="errors"></param>
        ''' <returns>true:error false:non error</returns>
        Private Function CheckError(ByRef errors As List(Of String)) As Boolean

            Dim errorsCheck = New List(Of String)
            errors.ForEach(Sub(x)
                               If x <> "" Then
                                   errorsCheck.Add(x)
                               End If
                           End Sub)
            errors.Clear()
            ' 参照元 更新
            errors = errorsCheck
            Return errors.Count <> 0
        End Function

        ''' <summary>
        ''' DB error
        ''' </summary>
        ''' <param name="tran"></param>
        Private Sub DBError(tran As OracleTransaction)

            tran.Rollback()

        End Sub

        ' エラーリスト出力ボタン
        Protected Sub btnExportErrorList_Click(sender As Object, e As EventArgs)
            lblResult.Text = ""
            lblError.Text = ""
            'lblError.Text = "開発未着手"
            Dim errors As New List(Of String)()

            ' 値取得
            Dim ProcessingStartDate As Date = DateTime.Now

            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(Utils.GetConnectionString())
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()
            ' Table class access                            
            ' 受注ワーク
            Dim reps As New OrderStageRepository(Utils.GetConnectionString())

            '出荷状況エラーリスト出力	
            'PROD_PLAN_STRA_VIEW（生産計画出力一覧）をExcel出力する。	
            Dim repos = New OrderStraRepository(Utils.GetConnectionString())
            'Dim ErrorRows = repos.GetOrderStras(conn, tran, status:="POST_PLAN_DUE_SET")
            Dim ErrorRows = repos.GetOrderStage(conn, tran, status:="POST_PLAN_DUE_SET", activeFlag:="N")
            Dim FileDate = DateTime.Now
            Dim strPath = Server.MapPath("~/App_Data/Files/")

            Try


                errors.Add(OrderProductionPlanExcelFile.ShippingStatusErrorExcelOut(strPath, FileDate, repos.ToClass(ErrorRows)))
                If (CheckError(errors)) Then
                    ' エラー
                End If
            Catch ex As Exception
                Dim m = ex.Message
                errors.Add(ex.Message)
                'lblError.Text = ex.Message
            Finally
                If (errors.Count > 0) Then
                    'lblError.Text = errors(0)
                    lblError.Text = String.Join(vbCrLf, errors)
                End If
            End Try

        End Sub
#End Region

    End Class
End Namespace