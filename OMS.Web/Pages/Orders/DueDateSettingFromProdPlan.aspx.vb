Imports System
Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data
Imports OMS.Data.OrderDiferenceExcelFile
Imports OMS.Data.SectmRepository
Imports OMS.Web.Pages.Masters.File
Imports Oracle.ManagedDataAccess.Client

Namespace Pages.Orders
    Public Class DueDateSettingFromProdPlan
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補の初期化
                LoadSearchConditionLists()

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
            Dim customerNameList As List(Of String) = repo.GetCustomerNames(loginUserId)

            lstSearchCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstSearchProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstSearchCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
            lstSearchCustomerName.InnerHtml = BuildOptions(customerNameList)
        End Sub

        ' <option> の生成を共通化
        Private Function BuildOptions(values As IEnumerable(Of String)) As String
            Dim sb As New StringBuilder()
            For Each v In values
                sb.AppendFormat("<option value=""{0}"" />", v)
            Next
            Return sb.ToString()
        End Function
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

            BindSelectCustomersGrid(customerCode, customerName, profitCenter, customerUnitName)
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_Click(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""

            BindSelectCustomersGrid()
        End Sub
#End Region

#Region "GridView バインド / イベント"
        ' Customer候補の一覧をバインド
        Private Sub BindSelectCustomersGrid(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing
        )
            Dim repo As New CustomerRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetCustomerList(
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
        Protected Sub gvSelectCustomers_RowDataBound(sender As Object, e As GridViewRowEventArgs) _
            Handles gvSelectCustomers.RowDataBound

            If e.Row.RowType = DataControlRowType.DataRow Then
                'Dim chk As CheckBox = TryCast(e.Row.FindControl("chkDueDateSetting"), CheckBox)
                Dim chk As WebControls.CheckBox = TryCast(e.Row.FindControl("chkDueDateSetting"), WebControls.CheckBox)
                If chk IsNot Nothing Then
                    ' 個別のチェック操作時にヘッダー状態を更新
                    chk.InputAttributes("onclick") =
                        $"OMS.Grid.updateHeader('{gvSelectCustomers.ClientID}', 'chkDueDateSettingAll', 'chkDueDateSetting');"
                End If
            End If
        End Sub
#End Region

#Region "納期設定 / 差異出力"
        ' 納期設定ボタン
        Protected Sub btnDueDateSetting_Click(sender As Object, e As EventArgs)
            'lblError.Text = "開発未着手"

            Dim errors As New List(Of String)()
            Dim valid As Integer = 0
            Dim count As Integer = 0
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
            ' ファイルリスト
            Dim fileList As List(Of String) = New List(Of String)()

            Try
                ' 処理対象 がチェックされている行
                For Each row In gvSelectCustomers.Rows

                    Dim chk = TryCast(row.FindControl("chkDueDateSetting"), WebControls.CheckBox)

                    If chk IsNot Nothing Then
                        ' チェックがある時
                        If chk.Checked Then

                            Dim idx As Integer = row.RowIndex
                            Dim keys = gvSelectCustomers.DataKeys(idx)

                            ' 取引先設定ID 取得＆安全化
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

                            ' 生産計画ワーク削除
                            ' ①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 	
                            ' ②	STATUS（ステータス） = 'PLAN_SET'	
                            errors.Add(reps.Delete(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId, status:="PLAN_SET"))
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            repo.Update(conn, tran, OrderRepository.OrdersTable.ProductPlan, kOrderId:=577, status:="PLAN_SET")
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If

                            ' 生産計画ワーク追加 (生産計画より)
                            ' ①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 
                            ' ②	STATUS（ステータス） = 'PLAN_SET'
                            ' ③	ACTIVE_FLAG（有効フラグ）= 'Y'
                            Dim orders = repo.GetOrders(conn, tran, OrderRepository.OrdersTable.ProductPlan, status:="PLAN_SET", activeFlag:="Y", customerSettingId:=customerSettingId)
                            Dim orderRows = repo.ToClass(orders)
                            count += orderRows.Count
                            '#If DEBUG Then
                            '                            ' #### DEBUG 評価
                            '                            If (orderRows.Count <> 0) Then
                            '                                orderRows(0).DueDate = New DateTime(2026, 5, 1)
                            '                                orderRows(0).ShipDate = New DateTime(2026, 5, 1)
                            '                                orderRows(0).ShipScheduledDate = New DateTime(2026, 5, 1)
                            '                            End If
                            '                            ' #### DEBUG
                            '#End If
                            errors.Add(reps.InsertRange(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderRows))
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If

                            ' STRA 納期設定 生産計画ワーク
                            ' 1. SHIP_SCHEDULED_DATE(出荷予定日) ORDERS_STAGE.SHIP_SCHEDULED_DATE + USRDEFFLDF.FUSRDEC1
                            ' 2. SHIP_DATE(出荷日) ORDERS_STAGE.SHIP_DATE +SHPROUTM.FTRANLT
                            ' 3. STATUS(ステータス) 'POST_PLAN_DUE_SET'
                            ' 4. UPDATED_AT(更新日時)
                            ' 5. UPDATED_USER_ID(更新ユーザーID)
                            ' 6. UPDATED_PG_ID(更新プログラムID)
                            ' 7. CUSTOMER_SETTING_ID(取引先設定ID)
                            ' 対象の生産計画テーブルorders の受注ID(ORDER_ID)を 生産計画ワークから探して
                            ' 納期設定値を上書きする
                            Dim repu = New UsrDeffIdfRepository(Utils.GetConnectionString())
                            Dim sectm = New SectmRepository(Utils.GetConnectionString())
                            Dim shproutm = New ShproutmRepository(Utils.GetConnectionString())
                            For Each orderRow In orderRows
                                ' STRAMMIC DB チェック  
                                Dim deliveryCode = orderRow.DeliveryCode
                                errors.Add(shproutm.CheckShippingDestination(conn, tran, customerCode, deliveryCode))
                                If (CheckError(errors)) Then
                                    ' エラー DB更新無効
                                    DBError(tran)
                                    Continue For
                                End If

                                ' 納期計算
                                Dim itemNo As String = orderRow.ItemNo
                                'FUSRDEC1 ((A)品揃リードタイム)
                                Dim assortLeadTime = shproutm.GetAssortmentLeadTime(customerCode, profitCenter, itemNo)
                                'FTRANLT(輸送L/T)
                                Dim transferLeadTime = shproutm.GetTransferLeadTime(orderRow.ShipTo, orderRow.ShipStockLocation)
                                Dim orderid = orderRow.OrderId

                                Dim shipScheduledDate = orderRow.ShipScheduledDate
                                Dim shipDateOrg = orderRow.ShipDate
                                '#If DEBUG Then
                                '                                ' #### DEBUG 評価
                                '                                transferLeadTime = 3
                                '                                assortLeadTime = 3
                                '                                ' #### DEBUG
                                '#End If
                                ' 日付確認
                                If (shipScheduledDate Is Nothing) Or (shipDateOrg Is Nothing) Then
                                    errors.Add("出荷日/出荷予定日 の設定が間違っています。")
                                End If
                                If (CheckError(errors)) Then
                                    ' エラー DB更新無効
                                    DBError(tran)
                                    Continue For
                                End If

                                Dim calType = "00001"
                                Dim cal = New CalenderRepository(Utils.GetConnectionString())
                                Dim shipDate = cal.AddWorkingDays(conn, tran, calType, shipScheduledDate.Value, assortLeadTime)
                                Dim dueDate = cal.AddWorkingDays(conn, tran, calType, shipDateOrg.Value, transferLeadTime)

                                ' UPDATE(生産計画)
                                Dim status = "POST_PLAN_DUE_SET"
                                Dim updateAt = ProcessingStartDate
                                Dim updateUserId = PageHelpers.GetUserId(Me)
                                Dim updatePgId = "DueDateSetting(PostPlan)"

                                ' SHIP_DATE(出荷日)
                                ' DUE_DATE(希望納期)
                                ' STATUS(ステータス)
                                ' UPDATED_AT(更新日時)
                                ' UPDATED_USER_ID(更新ユーザーID)
                                ' UPDATED_PG_ID(更新プログラムID)
                                ' CUSTOMER_SETTING_ID(取引先設定ID)
                                errors.Add(reps.UpdateDeadline(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderId:=orderid, shipDate:=shipDate, dueDate:=dueDate, status:=status, updatedAt:=updateAt, updatedUserId:=updateUserId, updatedPgId:=updatePgId))
                                '#If DEBUG Then
                                '                                ' #### DEBUG
                                '                                tran.Commit()
                                '                                tran = conn.BeginTransaction()
                                '                                ' #### DEBUG
                                '#End If
                                If (CheckError(errors)) Then
                                    ' エラー DB更新無効
                                    DBError(tran)
                                    Continue For
                                End If
                            Next

                            ' Update
                            ' 正規データ更新 生産計画
                            ' CUSTOMER_SETTING_ID と STATUS='POST_PLAN_DUE_SET'を抽出し
                            ' PROD_PLAN_ID が共通するレコードの フィールドを更新
                            ' SHIP_SCHEDULED_DATE(出荷予定日)
                            ' SHIP_DATE(出荷日)
                            ' STATUS(ステータス)
                            ' UPDATED_AT(更新日時)
                            ' UPDATED_USER_ID(更新ユーザーID)
                            ' UPDATED_PG_ID(更新プログラムID)
                            ' 2026/06/03 SQL 更新に変更
                            errors.Add(repo.ProductionPlanUpdate(conn, tran, customerSettingId))
                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If

                            'For Each orderRow In orderRows
                            '    Dim dt = reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderId:=orderRow.OrderId)
                            '    Dim orderStageRows = reps.ToClass(dt)
                            '    If orderStageRows Is Nothing Then
                            '        Exit Try
                            '    End If
                            '    Dim orderStageRow = orderStageRows(0)

                            '    errors.Add(repo.UpdateDeadline(conn, tran, OrderRepository.OrdersTable.ProductPlan, orderId:=orderStageRow.OrderId, orderStageRow.ShipScheduledDate, shipDate:=orderStageRow.ShipDate, status:=orderStageRow.Status, updatedAt:=orderStageRow.UpdatedAt, updatedUserId:=orderStageRow.UpdatedUserId, updatedPgId:=orderStageRow.UpdatedPgId))
                            '    '#If DEBUG Then
                            '    '                                ' #### DEBUG
                            '    '                                tran.Commit()
                            '    '                                tran = conn.BeginTransaction()
                            '    '                                ' #### DEBUG
                            '    '#End If
                            '    If (CheckError(errors)) Then
                            '        ' エラー DB更新無効
                            '        DBError(tran)
                            '        Continue For
                            '    End If
                            'Next

                            ' 生産計画履追加
                            ' 生産計画データ取得
                            ' CUSTOMER_SETTING_ID(取引先設定ID)
                            ' STATUS(ステータス) 'DUE_SET'
                            orders = repo.GetOrders(conn, tran, OrderRepository.OrdersTable.ProductPlan, status:="POST_PLAN_DUE_SET", customerSettingId:=customerSettingId)
                            errors.Add(reph.InsertRange(conn, tran, OrderHistoryRepository.OrdersTable.ProductPlan, repo.ToClass(orders)))
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            If (CheckError(errors)) Then
                                ' エラー DB更新無効
                                DBError(tran)
                                Continue For
                            End If

                            ' 生産計画差異取得
                            ' 生産計画差異取得(内示)
                            Dim difu = reph.UnNoticeDifferenceToClass(reph.AfterProductionPlanningUnofficialNoticeDifference(conn, tran, customerSettingId))
                            ' 生産計画差異取得(確定/納入指示)
                            Dim difd = reph.InstructionDifferenceToClass(reph.AfterProductionPlanningDeliveryInstructionDifference(conn, tran, customerSettingId))
                            ' 生産計画差異リスト出力
                            Dim strPath = Server.MapPath("~/App_Data/Files/")
                            Dim filename = CreateOorderDiferenceExcelFile(strPath, DiffFileTiminge.AfterProductionPlanning, ProcessingStartDate, difu, difd, customerSettingId)
                            fileList.Add(filename)
                            'If (filename <> "") Then
                            '    Utils.FileTransfer2(Response, Server, filename)
                            'Else
                            '    ' Excel ファイル作成 Error
                            '    errors.Add("STRA納期設定ページ_生産計画後 Excel ファイル作成エラー")
                            '    If (CheckError(errors)) Then
                            '        ' エラー DB更新無効
                            '        DBError(tran)
                            '        Continue For
                            '    End If
                            'End If
                            valid += orderRows.Count
                        Else
                        End If
                    End If
                    ' 1取引先終了時
                    tran.Commit()
                    tran = conn.BeginTransaction()
                Next
                Dim orderFilename = repo.GeOrderZipFilename("生産計画差異リスト", ProcessingStartDate)
                'Utils.FilesTransfer(Response, Server, fileList, orderFilename)
                ' 別ページでDownload 処理を行う (FileList file はDownload 処理内で削除する)
                Dim fileListName = IO.Path.Combine(Server.MapPath("~/App_Data/Files/"), Utils.GetTempFileName("FileList.txt"))
                Utils.SaveFileList(fileListName, fileList)
                Dim url As String = $"DownloadProcess.ashx?file={HttpUtility.UrlEncode(orderFilename)}&list={HttpUtility.UrlEncode(fileListName)}"
                Dim script As String = $"document.getElementById('downloadFrame').src = '{url}';"
                ClientScript.RegisterStartupScript(Me.GetType(), "downloadScript", script, True)

            Catch ex As Exception
                Dim err = ex.Message
                errors.Add(err)
            Finally
                lblError.Text = ""
                If (errors.Count > 0) Then
                    'lblError.Text = errors(0)
                    lblError.Text = String.Join(vbCrLf, errors)
                End If
                If (count = 0 And valid = 0) Then
                    lblResult.Text = $"有効なデータが無かったため納期設定は行われませんでした。"
                Else
                    lblResult.Text = $"{count}件中{valid}件の納期設定を行いました。"
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

        ' 差異リスト出力ボタン
        Protected Sub btnExportDiffList_Click(sender As Object, e As EventArgs)
            Dim errors As New List(Of String)()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)

            ' 値取得
            Dim ProcessingStartDate As Date = DateTime.Now
            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(Utils.GetConnectionString())
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()
            ' Table class access 
            ' 受注履歴
            Dim reph = New OrderHistoryRepository(Utils.GetConnectionString())
            ' ファイルリスト
            Dim fileList As List(Of String) = New List(Of String)()
            Try
                ' 処理対象 がチェックされている行
                For Each row In gvSelectCustomers.Rows

                    Dim chk = TryCast(row.FindControl("chkDueDateSetting"), WebControls.CheckBox)

                    If chk IsNot Nothing Then
                        Dim idx As Integer = row.RowIndex
                        Dim keys = gvSelectCustomers.DataKeys(idx)
                        ' 取引先設定ID 取得＆安全化
                        Dim customerSettingId As Integer
                        Dim csidObj = keys("CustomerSettingId")
                        If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), customerSettingId) Then
                            errors.Add($"Row {idx}：CustomerSettingIdが不正")
                            Continue For
                        End If

                        ' チェックがある時
                        If chk.Checked Then
                            ' 受注差異取得
                            ' 受注差異取得(内示)
                            Dim difu = reph.UnNoticeDifferenceToClass(reph.AfterProductionPlanningUnofficialNoticeDifference(conn, tran, customerSettingId))
                            ' 受注差異取得(確定/納入指示)
                            Dim difd = reph.InstructionDifferenceToClass(reph.AfterProductionPlanningDeliveryInstructionDifference(conn, tran, customerSettingId))
                            ' 受注差異リスト出力
                            Dim strPath = Server.MapPath("~/App_Data/Files/")
                            Dim filename = CreateOorderDiferenceExcelFile(strPath, DiffFileTiminge.AfterProductionPlanning, ProcessingStartDate, difu, difd, customerSettingId)
                            fileList.Add(filename)
                            'If (filename <> "") Then
                            '    Utils.FileTransfer2(Response, Server, filename)
                            'Else
                            '    ' Excel ファイル作成 Error
                            'End If
                        End If
                    End If
                Next
                Dim repo = New OrderRepository(Utils.GetConnectionString())
                Dim orderFilename = repo.GeOrderZipFilename("生産計画差異リスト", ProcessingStartDate)
                'Utils.FilesTransfer(Response, Server, fileList, orderFilename)
                ' 別ページでDownload 処理を行う
                Dim fileListName = IO.Path.Combine(Server.MapPath("~/App_Data/Files/"), Utils.GetTempFileName("FileList.txt"))
                Utils.SaveFileList(fileListName, fileList)
                Dim url As String = $"DownloadProcess.ashx?file={HttpUtility.UrlEncode(orderFilename)}&list={HttpUtility.UrlEncode(fileListName)}"
                Dim script As String = $"document.getElementById('downloadFrame').src = '{url}';"
                ClientScript.RegisterStartupScript(Me.GetType(), "downloadScript", script, True)
            Catch
            Finally
                tran.Dispose()
                conn.Close()
                conn.Dispose()
            End Try
        End Sub
#End Region

    End Class
End Namespace