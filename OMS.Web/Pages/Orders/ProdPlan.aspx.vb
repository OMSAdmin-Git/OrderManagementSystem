Imports System
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DocumentFormat.OpenXml.Drawing.Diagrams
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Office2010.Excel
Imports DocumentFormat.OpenXml.Office2013.Excel
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.VariantTypes
Imports DocumentFormat.OpenXml.Vml
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data
Imports OMS.Web.Pages.Masters
Imports OMS.Web.Pages.Masters.CustomerSetting
Imports OMS.Web.Pages.Masters.Folder
Imports Oracle.ManagedDataAccess.Client
Imports CheckBox = System.Web.UI.WebControls.CheckBox

Namespace Pages.Orders
    Public Class ProdPlan
        Inherits System.Web.UI.Page

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補を初期化（Utils.BuildOptions で <option> を生成）
                LoadSearchConditionLists()

                ' 初期表示（一覧）
                BindSelectCustomersGrid()

                lblProdPlanContinueMessage.Visible = False
                btnProdPlanOK.Visible = False
                btnProdPlanNO.Visible = False

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

            lstSearchCustomerCode.InnerHtml = Utils.BuildOptions(customerCodeList)
            lstSearchProfitCenter.InnerHtml = Utils.BuildOptions(profitCenterList)
            lstSearchCustomerUnitName.InnerHtml = Utils.BuildOptions(customerUnitNameList)
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
        ' 顧客候補の一覧をバインド
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

        ' 行データバウンドイベント
        Protected Sub gvSelectCustomers_RowDataBound(sender As Object, e As GridViewRowEventArgs) _
            Handles gvSelectCustomers.RowDataBound

            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox = TryCast(e.Row.FindControl("chkProdPlan"), CheckBox)
                If chk IsNot Nothing Then
                    ' 個別のチェック操作時にヘッダー状態を更新
                    chk.InputAttributes("onclick") =
                        $"OMS.Grid.updateHeader('{gvSelectCustomers.ClientID}', 'chkProdPlanAll', 'chkProdPlan');"
                End If
            End If
        End Sub
#End Region

#Region "アクション（生産計画／Excel出力／Excel取込）"

        Protected Sub btnTest1_Click(sender As Object, e As EventArgs) Handles btnTest1.Click
            'btnTest1.Visible = False
            'lblProdPlanContinueMessage.Visible = False
            'btnProdPlanOK.Visible = False
            'btnProdPlanNO.Visible = False

            btnProdPlanNO_Click(sender, e)

        End Sub
        Protected Sub btnTest2_Click(sender As Object, e As EventArgs) Handles btnTest1.Click
            lblProdPlanContinueMessage.Visible = True
            btnProdPlanOK.Visible = True
            btnProdPlanNO.Visible = True
            btnTest1.Visible = True

        End Sub


        ' 生産計画ボタン
        Protected Sub btnProdPlan_Click(sender As Object, e As EventArgs) Handles btnProdPlan.Click
            'lblError.Text = "開発未着手"

            lblError.Text = ""
            lblResult.Text = ""
            Dim errors As New List(Of String)()
            Dim valid As Integer = 0
            Dim count As Integer = 0
            Dim loginUserId As String = PageHelpers.GetUserId(Me)

            ' 値取得
            Dim ProcessingStartDate As Date = New Date(DateTime.Now.Year, DateAndTime.Now.Month, DateAndTime.Now.Day)

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
                ' 処理対象 がチェックされている行
                For Each row In gvSelectCustomers.Rows

                    Dim chk = TryCast(row.FindControl("chkProdPlan"), CheckBox)

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

                            ' Customer Setting ID リストに追加
                            idList.Add(customerSettingId)

                            ' 生産計画ワーク削除
                            ' ①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 	
                            errors.Add(reps.Delete(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId))
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            'repo.Update(conn, tran, OrderRepository.OrdersTable.Orders, kOrderId:=21, status:="PROCESSED")
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            If (CheckError(errors)) Then
                                ' エラー
                            End If

                            ' 生産計画ワーク追加
                            ' ①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 
                            ' ②	STATUS（ステータス） = 'DUE_SET'
                            ' ③	ACTIVE_FLAG（有効フラグ）= 'Y'
                            Dim orders = repo.GetOrders(conn, tran, OrderRepository.OrdersTable.Orders, status:="DUE_SET", activeFlag:="Y", customerSettingId:=customerSettingId)
                            Dim orderRows = repo.ToClass(orders)
                            errors.Add(reps.InsertRange(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderRows))
                            '#If DEBUG Then
                            '                            ' #### DEBUG
                            '                            tran.Commit()
                            '                            tran = conn.BeginTransaction()
                            '                            ' #### DEBUG
                            '#End If
                            If (CheckError(errors)) Then
                                ' エラー    
                            End If

                            ' 生産計画条件マスタ取得 (正常確認)
                            ' ①取引先設定ID：    選択されたデータのCUSTOMER_SETTING_ID（取引先設定ID）   
                            ' ②受注区分：        選択されたデータのORDER_TYPE（受注区分）    
                            ' ③分割区分：        選択されたデータのPRORATED_TYPE（分割区分）
                            For Each orderRow In orderRows
                                count = count + 1
                                Dim orderType = orderRow.OrderType
                                Dim proratedType = orderRow.ProratedType
                                'Dim prodPlanRule = prodp.GetProdPlanRule(conn, tran, customerSettingId, orderType, proratedType)
                                Dim prodPlanRule = prodp.GetProdPlanRule(conn, tran, customerSettingId)
                                If (prodPlanRule IsNot Nothing) Then
                                    Dim splitCaseFlag = prodPlanRule.SplitCaseFlag
                                    Dim prodPlanRuleId = prodPlanRule.ProdPlanRuleId
                                    ' SPLIT_CASE_FLAG 分割パターンを使用するとき
                                    If (splitCaseFlag = "Y") Then
                                        Dim splitCase = splitp.GetSplitCaseRuleListSortByPriolity(conn, tran, prodPlanRuleId)
                                        If (splitCase.Count <> 0) Then
                                            valid = valid + 1
                                        Else
                                        End If
                                    Else
                                        ' 分割していないときは +1
                                        valid = valid + 1
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
                ' Customer Setting ID リスト重複をまとめる
                Dim idSelect = idList.Distinct().ToList()

                ' judge
                tran.Commit()
                tran.Dispose()
                conn.Close()
                conn.Dispose()

                ' SPLIT_CASE_FLAG = Y で分割指定が正しいとき
                ' SPLIT_CASE_FLAG = N のとき
                If (count = valid) Then
                    ' 分割処理を呼び出す
                    btnProdPlanOK_Click(sender, e)
                Else
                    ' SPLIT_CASE_FLAG = Y で分割指定がないとき
                    ' 確認メッセージを表示する
                    ' 生産計画条件マスタに登録されていないデータを検出しました。​生産計画を続行しますか？
                    lblProdPlanContinueMessage.Visible = True
                    btnProdPlanOK.Visible = True
                    btnProdPlanNO.Visible = True
                End If
            Catch ex As Exception
                Dim m = ex.Message
                errors.Add(ex.Message)
            Finally
                ' 取引先単位で 確定/取り消し が必要なので ループ内で 行う必要がある
                ' Commit/Rollback  Transaction
                If errors.Count = 0 Then
                    'tran.Commit()
                Else
                    'tran.Rollback()
                    lblError.Text = String.Join(vbCrLf, errors)
                End If
                'tran.Dispose()
                'conn.Close()    
                'conn.Dispose()
            End Try

        End Sub

        ' いいえボタン click event handler
        Protected Sub btnProdPlanNO_Click(sender As Object, e As EventArgs) Handles btnProdPlanNO.Click

            ' 確認ボタン/メッセージ非表示
            lblProdPlanContinueMessage.Visible = False
            btnProdPlanOK.Visible = False
            btnProdPlanNO.Visible = False
        End Sub

        ' はいボタン click event handler
        Protected Sub btnProdPlanOK_Click(sender As Object, e As EventArgs) Handles btnProdPlanOK.Click

            'lblError.Text = "開発未着手"
            lblError.Text = ""
            lblResult.Text = ""
            Dim errors As New List(Of String)()
            Dim valid As Integer = 0
            Dim count As Integer = 0
            Dim loginUserId As String = PageHelpers.GetUserId(Me)

            ' 確認ボタン/メッセージ非表示
            ' File転送が入るため コードでコントロールを非表示にすることができないため
            ' aspx 内で OK Click のevent 内に非表示処理を入れ java で実行する

            ' 値取得(処理開始日 時間 00:00:00 に丸める )
            'Dim ProcessingStartDate As Date = DateSerial(Year(DateTime.Now), Month(DateTime.Now), Day(DateTime.Now))
#If DEBUG Then  '#DEBUG
            Dim ProcessingStartDate As Date = DateSerial(2026, 3, 1)
#End If
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
                ' 処理対象 がチェックされている行
                For Each row In gvSelectCustomers.Rows

                    Dim chk = TryCast(row.FindControl("chkProdPlan"), WebControls.CheckBox)

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

                            ' Customer Setting ID リストに追加
                            idList.Add(customerSettingId)

                            '' 受注ワーク追加
                            '' ①	CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID（取引先設定ID） 
                            '' ②	STATUS（ステータス） = 'DUE_SET'
                            '' ③	ACTIVE_FLAG（有効フラグ）= 'Y'
                            'Dim ordersRowsOrg = repo.ToClass(repo.GetOrders(conn, tran, status:="DUE_SET", activeFlag:="Y", customerSettingId:=customerSettingId))

                            ''------------------------------
                            '' ORDER_STAGE　Insert Range 取引先単位 全レコード追加
                            ''------------------------------
                            'errors.Add(reps.InsertRange(conn, tran, ordersRowsOrg))
                            'If (CheckError(errors)) Then
                            '    ' エラー
                            'End If

                            ' 仕様書では ワークレコードは 取得済みになっているためここで取得する
                            ' 生産計画ワーク取得
                            ' Orders -> OrdersStage に追加した OrderRow 
                            ' 作業用に 追加した OrderStege レコードを取得
                            Dim orderStageRowsOrg = reps.ToClass(reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, status:="DUE_SET", activeFlag:="Y", customerSettingId:=customerSettingId))

                            ' 生産計画条件マスタ 振り分け リスト取得 
                            ' CUSTOMER_SETTING_ID で抽出した Order を 下記の条件でグループ分けを行う
                            ' CUSTOMER_SETTING_ID（取引先設定ID）
                            ' CUSTOMER_ITEM_NO（客先品目No）
                            ' ORDER_TYPE（受注区分）
                            ' PRORATED_TYPE（分割区分）
                            ' SHIP_SCHEDULED_DATE（出荷予定日）のyyyy/mm

                            ' 上記 Order.Field でグループ分け
                            Dim ordersStageRowsGroup = orderStageRowsOrg.GroupBy(Function(x) New With {Key x.CustomerSettingId, Key x.CustomerItemNo, Key x.OrderType, Key x.ProratedType, Key x.ShipScheduledDate.Value.Year, Key x.ShipScheduledDate.Value.Month}).ToList()
                            ' 当月 顧客単位の OrderStage リスト 
                            Dim ordersStageRows = New List(Of OrdersStageRow)
                            ' グループ分け　Loop
                            'For Each og In orderRowsGroup
                            For Each og In ordersStageRowsGroup
                                ordersStageRows.Clear()
                                ' グループ内 order record リスト作成 (yyyymm でグループ分けされたものを取得)
                                For Each ri In og.ToList()
                                    ordersStageRows.Add(ri)
                                Next
                                ' 更新しない項目については、出荷予定日の昇順で最初のレコードの情報を反映させる。
                                Dim ordersStageRow = ordersStageRows.ToList().OrderBy(Function(x) x.ShipScheduledDate).FirstOrDefault()
                                Dim prodPlanRule = prodp.GetProdPlanRule(conn, tran, customerSettingId)
                                'prodPlanRule = Nothing
                                If (prodPlanRule IsNot Nothing) Then
                                    Dim splitCaseFlag = prodPlanRule.SplitCaseFlag
                                    'Dim prodPlanRuleId = prodPlanRule.ProdPlanRuleId
                                    ' SPLIT_CASE_FLAG 分割パターンを使用するとき
                                    If (splitCaseFlag = "Y") Then
                                        'Dim splitCase = splitp.GetSplitCaseRuleListSortByPriolity(prodPlanRuleId)
                                        'If (splitCase IsNot Nothing) Then
                                        'Else
                                        'End If
                                    End If
                                    ' 分割フラグ確認 (SPLIT_FLAG)
                                    Dim splitFlag = prodPlanRule.SplitFlag
                                    Dim orderType = ordersStageRow.OrderType
                                    Dim proratedType = ordersStageRow.ProratedType
                                    If (splitFlag = "Y") Then
                                        ' 2026/03/09 仕様変更分 ----
                                        ' 内示指定の時 かつ 日割り指定以外
                                        If (orderType = 1 And proratedType = 2) Then
                                            ' 2026/03/09 仕様変更分 ------

                                            ' 分割方法
                                            Dim splitMethodType = prodPlanRule.SplitMethodType

                                            ' 任意分割を使用する場合 ※この段階では注文数がわからないため何もしない
                                            'Dim splitCaseFlag = prodPlanRule.SplitCaseFlag
                                            If (splitCaseFlag = "Y") Then
                                            End If

                                            '●	生産計画の適用
                                            '１.CALEM（カレンダーファイル）から対象月の稼働日でレコードを作成する。	
                                            Dim startDate As Date
                                            Dim endDate As Date

                                            ' 日割り期間 を計算
                                            Dim splitStartType = prodPlanRule.SplitStartType
                                            Dim shipScheduledDate = ordersStageRow.ShipScheduledDate
                                            'SPLIT_START_TYPE（分割開始区分）出荷予定日 期間を決定
                                            Select Case (splitStartType)
                                                Case 1
                                                    '   1(月初)        ：  月初～月末                         （例：2026年3月が対象の場合、3/1～3/31）
                                                    startDate = DateSerial(Year(shipScheduledDate), Month(shipScheduledDate), 1)
                                                    endDate = DateSerial(Year(shipScheduledDate), Month(shipScheduledDate), DateTime.DaysInMonth(Year(shipScheduledDate), Month(shipScheduledDate)))
                                                Case 2
                                                    '   2(前月第4週)    ：  前月第4週～当月第3週              （例：2026年3月が対象の場合、2/23～3/20）
                                                    startDate = Get4WeekOfLastMonth(shipScheduledDate)
                                                    endDate = Get3WeekOfThisMonth(shipScheduledDate)
                                                Case 3
                                                    '   3(納期の4週前)  ：  出荷予定日4週前～出荷予定日1週前  （例：出荷予定日が2026/3/13の場合、2/13～3/6）
                                                    startDate = Get4WeeksBefore(shipScheduledDate)
                                                    endDate = Get1WeeksBefore(shipScheduledDate)
                                            End Select

                                            ' 需要数 集計数
                                            Dim qty = ordersStageRows.Sum(Function(x) x.DemandQty)
                                            ' 端数の行先 1: 先頭 2: 最後尾
                                            Dim carryToType = prodPlanRule.CarryToType
                                            ' 丸め単位
                                            Dim unit = prodPlanRule.SplitRoudingUnit
                                            ' 分割比
                                            Dim splitRationType = prodPlanRule.SplitRationType
                                            ' 1、2 以外は 1に固定
                                            splitRationType = If(splitRationType <> 1 And splitRationType <> 2, 1, splitRationType)

                                            ' 丸め 1以下の場合 強制的に 1
                                            If (unit < 1) Then
                                                unit = 1
                                            End If
                                            ' 分割条件 読み込み
                                            Dim prodPlanRuleId = prodPlanRule.ProdPlanRuleId
                                            Dim splitCase = splitp.GetSplitCaseRuleListSortByPriolity(conn, tran, prodPlanRuleId)
                                            ' Split case Flag 判断
                                            If (splitCaseFlag = "Y") Then
                                                ' 条件がない場合は prodPlanRule を使用する
                                                For Each sc In splitCase
                                                    Dim judge = False

                                                    Select Case sc.QtyConditionType
                                                        Case "GE"   ' 超え
                                                            If (qty > sc.Qty) Then
                                                                judge = True
                                                            End If
                                                        Case "GT"   ' 以上
                                                            If (qty > sc.Qty) Then
                                                                judge = True
                                                            End If
                                                        Case "LT"   ' 未満
                                                            If (qty < sc.Qty) Then
                                                                judge = True
                                                            End If
                                                        Case "LE"   ' 以下
                                                            If (qty < sc.Qty) Then
                                                                judge = True
                                                            End If
                                                    End Select
                                                    ' 分割条件 あり
                                                    If (judge = True) Then
                                                        ' 分割条件が確定をする
                                                        splitMethodType = sc.SplitMethodType
                                                        Exit For
                                                    End If
                                                Next
                                            End If

                                            ' 週まるめ 時 (開始が月をさかのぼることがある、SPLIT_START_TYPEとは両立できない)
                                            If (splitMethodType = 5) Then
                                                ' start / end が変わる可能性がある
                                                Dim spDate = ordersStageRow.ShipScheduledDate
                                                ' 週先頭の月曜日
                                                startDate = GetDateMonday(spDate)
                                                endDate = DateSerial(Year(shipScheduledDate), Month(shipScheduledDate), DateTime.DaysInMonth(Year(shipScheduledDate), Month(shipScheduledDate)))
                                            End If

                                            ' 対象 ProdPlanStage のDemandQtyをクリアして 削除対象とする
                                            ' 2026/3/31 追加 DemandQty クリア
                                            For Each tg In ordersStageRows
                                                reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, kOrderId:=tg.OrderId, demandQty:=0)
                                                'ordersStageRows.ForEach(Sub(x) x.DemandQty = 0)
                                            Next

                                            ' 指定期間の カレンダーを取得(仕様書では カレンダーは もっと上で取得することになっているが 期間が判明するここで取得する)
                                            Dim calm As CalenderRepository = New CalenderRepository(Utils.GetConnectionString())
                                            Dim calender = calm.GetCalenderList(conn, tran, startDate, endDate)
                                            ' 営業日リスト
                                            Dim businessDayList As List(Of CalenderRow) = calender.Where(Function(x) x.HolidayFlag = "W").ToList()
                                            ' 営業日日数
                                            Dim businessDay = businessDayList.Count

                                            'Order_Stage テーブルに 対象月の稼働日でレコードを作成
                                            Dim ordersStage = New List(Of OrdersStageRow)()
                                            ' 抽出しているレコードの Calender 該当日を取得
                                            Dim targetDay = ordersStageRows.Join(businessDayList, Function(ord) ord.ShipScheduledDate, Function(bus) bus.DefDate, Function(ord, bus) bus).ToList()
                                            ' Base order record
                                            Dim baseOrder = New OrdersStageRow(ordersStageRow)
                                            For Each bday In businessDayList
                                                ' すでにレコードがある日付
                                                If (targetDay.Any(Function(x) x.DefDate = bday.DefDate)) Then
                                                Else
                                                    ' 未登録の稼働日
                                                    baseOrder.ShipScheduledDate = bday.DefDate  ' 営業日
                                                    baseOrder.UpdatedAt = ProcessingStartDate   ' 更新日
                                                    baseOrder.UpdatedUserId = loginUserId
                                                    baseOrder.UpdatedPgId = "ProductionPlanning(Execute)"
                                                    baseOrder.DemandQty = 0
                                                    ' 営業日の OrderStage Recode をリストに追加
                                                    ordersStage.Add(New OrdersStageRow(baseOrder))
                                                End If
                                            Next
                                            '------------------------------
                                            ' PROD_PLAN_STAGE　Insert Range 全営業日追加
                                            '------------------------------
                                            errors.Add(reps.InsertRange(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, ordersStage))
                                            If (CheckError(errors)) Then
                                                ' エラー
                                                DBError(tran)
                                                Continue For
                                            End If
                                            ' 読み直し
                                            Dim dt = reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, orderId:=baseOrder.OrderId)
                                            If (CheckError(errors)) Then
                                                ' エラー 
                                                DBError(tran)
                                                Continue For
                                            End If
                                            ordersStage.Clear()
                                            ordersStage = reps.ToClass(dt)
                                            '#If DEBUG Then
                                            '                                            ' #### DEBUG
                                            '                                            tran.Commit()
                                            '                                            tran = conn.BeginTransaction()
                                            '                                            ' #### DEBUG
                                            '#End If
                                            ' 分割方法確認 (SPLIT_METHOD_TYPE)
                                            Select Case (splitMethodType)
                                                ' 何もしない
                                                Case 0
                                                Case 6
                                                '日割り
                                                Case 1
                                                    Dim ans = GetRoundCalc(qty, businessDay, unit, splitRationType)
                                                    Dim roundup = ans.roundedUp     ' 注文/day
                                                    Dim days = ans.days             ' 期間
                                                    Dim fraction = ans.fraction     ' 端数

                                                    For cnt = 0 To days - 1
                                                        Dim setDay = GetDesignationDay(calender, "W", cnt)
                                                        Dim dqty = roundup
                                                        ' 先頭端数の時
                                                        If (cnt = 0) And (carryToType = 1) Then
                                                            dqty = fraction
                                                        End If
                                                        ' 先後尾端数の時
                                                        If (cnt = days - 1) And (carryToType = 2) Then
                                                            dqty = fraction
                                                        End If
                                                        Dim tg = ordersStage.Find(Function(x) x.ShipScheduledDate = setDay)
                                                        tg.DemandQty = dqty
                                                        tg.ShipPlanDate = setDay
                                                        'tg.CustomerOrderLineNo = "" ' ????? 仕様不明
                                                        ' 分割比
                                                        If (splitRationType = 2) Then
                                                            If (cnt = 0) Then
                                                                tg.DemandQty += roundup
                                                            End If
                                                        End If
                                                    Next
                                                '4分割
                                                Case 2
                                                    Dim ans = GetRoundCalc(qty, 4, unit, splitRationType)
                                                    Dim roundup = ans.roundedUp     ' 注文/day
                                                    Dim days = ans.days             ' 期間
                                                    Dim fraction = ans.fraction     ' 端数

                                                    For cnt = 0 To days - 1
                                                        Dim setDay = GetDesignationDay(calender, "W", cnt * (businessDay / 4))
                                                        Dim dqty = roundup
                                                        ' 先頭端数の時
                                                        If (cnt = 0) And (carryToType = 1) Then
                                                            dqty = fraction
                                                        End If
                                                        ' 先後尾端数の時
                                                        If (cnt = days - 1) And (carryToType = 2) Then
                                                            dqty = fraction
                                                        End If
                                                        Dim tg = ordersStage.Find(Function(x) x.ShipScheduledDate = setDay)
                                                        tg.DemandQty = dqty
                                                        tg.ShipPlanDate = setDay
                                                        'tg.CustomerOrderLineNo = "" ' ????? 仕様不明
                                                        ' 分割比
                                                        If (splitRationType = 2) Then
                                                            If (cnt = 0) Then
                                                                tg.DemandQty += roundup
                                                            End If
                                                        End If
                                                    Next

                                                '3分割
                                                Case 3
                                                    Dim ans = GetRoundCalc3D(qty, 4, unit)
                                                    Dim roundup = ans.roundedUp     ' 注文/day
                                                    Dim days = ans.days             ' 期間
                                                    Dim fraction = ans.fraction     ' 端数

                                                    For cnt = 0 To days ' 日付は月 4分割 のため
                                                        Dim setDay = GetDesignationDay(calender, "W", cnt * (businessDay / 4))
                                                        Dim dqty = roundup
                                                        ' 先頭端数の時
                                                        If (cnt = 0) And (carryToType = 1) Then
                                                            dqty = fraction
                                                        End If
                                                        ' 先後尾端数の時
                                                        If (cnt = days) And (carryToType = 2) Then
                                                            dqty = fraction
                                                        End If

                                                        ' 第2週をパス
                                                        If (cnt = 1) Then
                                                            Continue For
                                                        End If
                                                        Dim tg = ordersStage.Find(Function(x) x.ShipScheduledDate = setDay)
                                                        tg.DemandQty = dqty
                                                        tg.ShipPlanDate = setDay
                                                        'tg.CustomerOrderLineNo = "" ' ????? 仕様不明
                                                    Next

                                                '2分割
                                                Case 4
                                                    Dim ans = GetRoundCalc(qty, 2, unit, splitRationType)
                                                    Dim roundup = ans.roundedUp     ' 注文/day
                                                    Dim days = ans.days             ' 期間
                                                    Dim fraction = ans.fraction     ' 端数

                                                    For cnt = 0 To days - 1
                                                        Dim setDay = GetDesignationDay(calender, "W", cnt * (businessDay / 2))
                                                        Dim dqty = roundup
                                                        ' 先頭端数の時
                                                        If (cnt = 0) And (carryToType = 1) Then
                                                            dqty = fraction
                                                        End If
                                                        ' 先後尾端数の時
                                                        If (cnt = days - 1) And (carryToType = 2) Then
                                                            dqty = fraction
                                                        End If
                                                        Dim tg = ordersStage.Find(Function(x) x.ShipScheduledDate = setDay)
                                                        tg.DemandQty = dqty
                                                        tg.ShipPlanDate = setDay
                                                        'tg.CustomerOrderLineNo = "" ' ????? 仕様不明
                                                        ' 分割比
                                                        If (splitRationType = 2) Then
                                                            If (cnt = 0) Then
                                                                tg.DemandQty += roundup
                                                            End If
                                                        End If
                                                    Next
                                                '週まるめ
                                                Case 5
                                                    ' 月内の 週数
                                                    Dim wCount = GetWeeksInMonth(ordersStageRow.ShipScheduledDate)
                                                    Dim weekStart = startDate
                                                    ' 週 回数
                                                    For i As Integer = 0 To wCount - 1
                                                        'Dim wOrders = orderRows.FindAll(Function(x) x.ShipScheduledDate >= wStartDate And x.ShipScheduledDate <= wEndDate)
                                                        'Dim wOrders = ordersStageRows.FindAll(Function(x) x.ShipScheduledDate >= wStartDate And x.ShipScheduledDate <= wEndDate)
                                                        Dim wOrders = ordersStageRows.FindAll(Function(x) x.ShipScheduledDate >= weekStart And x.ShipScheduledDate <= DateSerial(Year(weekStart), Month(weekStart), Day(weekStart) + 6))
                                                        Dim dqty = wOrders.Sum(Function(x) x.DemandQty)
                                                        ' 週単位で 週内の DemandQty に 0代入 
                                                        wOrders.ForEach(Sub(x) x.DemandQty = 0)
                                                        ' その週の先頭営業日を取得 
                                                        '@@@@ 4/1 水曜日の場合  現状 [要確認]
                                                        If (dqty <> 0) Then
                                                            Dim setDay = GetDateWeekBusinessDay(calender, weekStart)
                                                            Dim tg = ordersStage.Find(Function(x) x.ShipScheduledDate = setDay)
                                                            tg.DemandQty = dqty
                                                            tg.ShipPlanDate = setDay
                                                            'tg.CustomerOrderLineNo = "" ' ????? 仕様不明
                                                        End If
                                                        ' 次の週
                                                        weekStart = DateSerial(Year(weekStart), Month(weekStart), Day(weekStart) + 7)
                                                    Next
                                            End Select
                                            '
                                            ' orderStageRow は DB から取得した内容 (処理経過で Dqty などは 変更される)
                                            ' 最終的に生成されるのは ordersStage で 開始日から終了日までの営業日レコードデータが
                                            ' 生成される。(不要なレコードも 含まれる。 DemandQty 0 のものも含まれるため)
                                            '----------------------------
                                            ' Update
                                            '----------------------------
                                            '【３．処理日以前のDEMAND_QTY（需要数）を集約する。】
                                            '・以下の条件に合致するデータが存在する場合、以下の処理を行う。
                                            '-条件：[処理開始日]より過去のSHIP_SCHEDULED_DATE（出荷予定日）に需要数が割り当てられたレコードが存在する
                                            '-処理：[処理開始日]より過去の出荷予定日に割り当てられた需要数を[処理開始日]のデータに合算する
                                            '
                                            '[処理開始日]の出荷予定日のレコードが存在しない場合は、新規にレコードを作成し、
                                            '需要数以外の項目は集約対象期間の昇順で先頭のレコードを参照する
                                            '（例：処理開始日が3/10の場合、3/1～3/9に割り当てられた需要数を3/10に合算する）
                                            '集約されたレコードの需要数をブランクに更新する
                                            Dim oldOrderOrder = ordersStage.FindAll(Function(x) x.ShipScheduledDate < ProcessingStartDate).OrderBy(Function(x) x.ShipScheduledDate).ToList()
                                            If (oldOrderOrder.Count > 0) Then
                                                ' 作業日に該当する レコードを取得
                                                Dim thisRecord = ordersStage.Find(Function(x) x.ShipScheduledDate = ProcessingStartDate)
                                                ' 
                                                If (thisRecord Is Nothing) Then
                                                    ' 昇順先頭レコードを先頭レコードをもとに作成
                                                    thisRecord = New OrdersStageRow(oldOrderOrder(0))
                                                    thisRecord.ShipScheduledDate = ProcessingStartDate
                                                    ordersStage.Add(thisRecord)
                                                    ' 分割期間よりも以前の場合 prod_plan_stage にレコードが無いため追加する
                                                    errors.Add(reps.Insert(conn, tran, OrderRepository.OrdersTable.ProductPlan, thisRecord))
                                                    '#If DEBUG Then
                                                    '                                                    ' #### DEBUG
                                                    '                                                    tran.Commit()
                                                    '                                                    tran = conn.BeginTransaction()
                                                    '                                                    ' #### DEBUG
                                                    '#End If
                                                    If (CheckError(errors)) Then
                                                        ' エラー
                                                        DBError(tran)
                                                        Continue For
                                                    End If
                                                End If
                                                ' 作業日以前の Qty をまとめる
                                                thisRecord.DemandQty = oldOrderOrder.Sum(Function(x) x.DemandQty)
                                                ' 作業日以前の Qty を0クリア
                                                oldOrderOrder.ForEach(Sub(x) x.DemandQty = 0)
                                            End If

                                            For Each tg In ordersStage
                                                ' OrdersStage 更新
                                                errors.Add(reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan,
                                                                        kOrderStageId:=tg.StageId,
                                                                        kOrderId:=tg.OrderId,
                                                                        kShipScheduledDate:=tg.ShipScheduledDate,
                                                                        demandQty:=tg.DemandQty,
                                                                        orderNo:=tg.OrderNo,
                                                                        shipPlanDate:=tg.ShipScheduledDate,
                                                                        updatedAt:=tg.UpdatedAt,
                                                                        updatedUserId:=tg.UpdatedUserId,
                                                                        updatedPgId:=tg.UpdatedPgId))
                                            Next
                                            '#If DEBUG Then
                                            '                                            ' #### DEBUG
                                            '                                            tran.Commit()
                                            '                                            tran = conn.BeginTransaction()
                                            '                                            ' #### DEBUG
                                            '#End If
                                            If (CheckError(errors)) Then
                                                ' エラー
                                                DBError(tran)
                                                Continue For
                                            End If

                                            ' 不要なレコードを削除する (DEMAND_QTY（需要数）がブランク)
                                            '----------------------------
                                            ' Delete 1レコードずつDelete
                                            '----------------------------
                                            ' 生産計画ワーク削除
                                            errors.Add(reps.Delete(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId, demandQty:=0))
                                            '#If DEBUG Then
                                            '                                            ' #### DEBUG
                                            '                                            tran.Commit()
                                            '                                            tran = conn.BeginTransaction()
                                            '                                            ' #### DEBUG
                                            '#End If
                                            If (CheckError(errors)) Then
                                                ' エラー
                                                DBError(tran)
                                                Continue For
                                            End If
                                        End If
                                    End If
                                End If
                                ' ステータス更新 STATUS "PLAN_SET"
                                '----------------------------
                                ' Update STATUS "PLAN_SET" 一括Update
                                '----------------------------
                                ' 生産計画ワーク更新
                                'Dim urows = reps.ToClass(reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, customerSettingId:=customerSettingId, status:="DUE_SET"))
                                'For Each trow In urows
                                'errors.Add(reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, kOrderStageId:=trow.StageId, status:="PLAN_SET"))
                                'Next

                                ' 一括更新では駄目 2026/5/14
                                'errors.Add(reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, kCustomerSettingId:=customerSettingId, kStatus:="DUE_SET", status:="PLAN_SET"))

                                ' 1レコード単位で更新
                                For Each tg In ordersStageRows
                                    errors.Add(reps.Update(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, kOrderId:=tg.OrderId, kCustomerSettingId:=customerSettingId, kStatus:="DUE_SET", status:="PLAN_SET"))
                                    If (CheckError(errors)) Then
                                        ' エラー
                                        DBError(tran)
                                        Continue For
                                    End If
                                Next

                                '#If DEBUG Then
                                '                                ' #### DEBUG
                                '                                tran.Commit()
                                '                                tran = conn.BeginTransaction()
                                '                                ' #### DEBUG
                                '#End If
                                If (CheckError(errors)) Then
                                    ' エラー
                                    DBError(tran)
                                    Continue For
                                End If
                            Next
                        End If
                    End If
                Next

                ' Customer Setting ID リスト重複をまとめる
                Dim idSelectList = idList.Distinct().ToList()

                ' Orders 削除 OrderStage に日割り登録したレコードを削除する
                ' PROD_PLAN（生産計画テーブル）のCUSTOMER_SETTING_ID（取引先設定ID）が処理対象と一致するレコード
                '----------------------------
                ' Delete 登録済みの Orders を 生産計画から 削除
                '----------------------------
                errors.Add(repo.DeleteRegisteredOrders(conn, tran, OrderRepository.OrdersTable.ProductPlan, idSelectList))
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                If (CheckError(errors)) Then
                    ' エラー
                    DBError(tran)
                End If
                '---------------------------- 
                ' Insert 登録済みの ProdPlanStageからPlodPlan へ登録
                '----------------------------
                '   OrdersRow の Listを取得   OrderStage のDataTable を OrderStageRow のクラス変換をして OrdersRow へ変換する
                Dim registerdOrders = OrdersStageRow.ToOrdersRow(reps.ToClass(reps.GetRegisteredOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, idSelectList)))
                errors.Add(repo.InsertRange(conn, tran, OrderRepository.OrdersTable.ProductPlan, registerdOrders))
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If

                If (CheckError(errors)) Then
                    ' エラー
                    DBError(tran)
                End If

                ' Insert 登録済みの ProdPlanStageからPlodPlanHistory へ登録
                errors.Add(reph.InsertRange(conn, tran, OrderRepository.OrdersTable.ProductPlan, registerdOrders))
                '#If DEBUG Then
                '                ' #### DEBUG
                '                tran.Commit()
                '                tran = conn.BeginTransaction()
                '                ' #### DEBUG
                '#End If
                If (CheckError(errors)) Then
                    ' エラー
                    DBError(tran)
                End If
                '----------------------------
                ' Excel 出力
                '----------------------------
                'Dim strPath = Server.MapPath("~/App_Data/Files/")
                'Dim filename = GetExcelFilename()
                ''errors.Add(exout.OrderExcelFile(filename, registerdOrders))
                'errors.Add(ProductPlanExcelOut(IO.Path.Combine(strPath, filename), registerdOrders))
                'If (CheckError(errors)) Then
                '    ' エラー
                'End If
                '' Excel ファイル転送
                'If (filename <> "") Then
                '    Utils.FileTransfer2(Response, Server, IO.Path.Combine(strPath, filename))
                'Else
                '    ' Excel ファイル作成 Error
                '    errors.Add("Error: Excel 生産計画リスト出力エラー")
                'End If

                Dim fileList As List(Of String) = New List(Of String)()
                Dim filename = IO.Path.Combine(Server.MapPath("~/App_Data/Files/"), GetExcelFilename())
                'errors.Add(exout.OrderExcelFile(filename, registerdOrders))
                errors.Add(ProductPlanExcelOut(filename, registerdOrders))
                If (CheckError(errors)) Then
                    ' エラー
                End If
                ' Excel ファイル転送
                fileList.Add(filename)
                Dim orderFilename = repo.GeOrderZipFilename("生産計画_出力日時", ProcessingStartDate)
                ' 別ページでDownload 処理を行う
                Dim fileListName = IO.Path.Combine(Server.MapPath("~/App_Data/Files/"), Utils.GetTempFileName("FileList.txt"))
                Utils.SaveFileList(fileListName, fileList)
                Dim url As String = $"DownloadProcess.ashx?file={HttpUtility.UrlEncode(orderFilename)}&list={HttpUtility.UrlEncode(fileListName)}"
                Dim script As String = $"document.getElementById('downloadFrame').src = '{url}';"
                ClientScript.RegisterStartupScript(Me.GetType(), "downloadScript", script, True)
            Catch ex As Exception
                errors.Add(ex.Message)
            Finally
                ' 取引先単位で 確定/取り消し が必要なので ループ内で 行う必要がある
                ' Commit/Rollback  Transaction
                If errors.Count = 0 Then
                    tran.Commit()
                Else
                    lblError.Text = String.Join(vbCrLf, errors)
                    tran.Rollback()
                End If
                tran.Dispose()
                conn.Close()
                conn.Dispose()
                btnProdPlanOK.Visible = False
                btnProdPlanNO.Visible = False
                lblProdPlanContinueMessage.Visible = False
            End Try

        End Sub

        ''' <summary>
        ''' Error メッセージ以外を削除する エラー数を返す
        ''' </summary>
        ''' <param name="errors"></param>.
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
        ''' <summary>
        ''' Excel出力ボタン event handler
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Sub btnExportProdPlanList_Click(sender As Object, e As EventArgs)

            'lblError.Text = "開発未着手"
            lblError.Text = ""
            lblResult.Text = ""
            Dim errors As New List(Of String)()
            Dim customerSettingIdList As List(Of Long) = New List(Of Long)()
            Dim repo As New OrderRepository(Utils.GetConnectionString())
            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(Utils.GetConnectionString())
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()

            ' 処理対象となる CustomerSettingId リスト
            Dim idList As New List(Of Long)

            ' Customer setting id を収集する
            ' 処理対象 がチェックされている行
            For Each row In gvSelectCustomers.Rows
                Dim chk = TryCast(row.FindControl("chkDueDateSetting"), WebControls.CheckBox)
                If chk IsNot Nothing Then
                    Dim idx As Integer = row.RowIndex
                    Dim keys = gvSelectCustomers.DataKeys(idx)
                    ' チェックがある時
                    If chk.Checked Then
                        Dim customerSettingId As Long
                        Dim csidObj = keys("CustomerSettingId")
                        If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), customerSettingId) Then
                            errors.Add($"Row {idx}：CustomerSettingIdが不正")
                            Continue For
                        End If
                        ' Customer Setting ID リストに追加
                        idList.Add(customerSettingId)
                    End If
                End If
            Next

            ' Customer Setting ID リスト重複をまとめる
            Dim idSelectList = idList.Distinct().ToList()
            Dim strPath = Server.MapPath("~/App_Data/Files/")
            'Dim path = GetWorkPath()
            Dim filename = GetExcelFilename()
            'errors.Add(exout.OrderExcelFile(filename, registerdOrders))

            Dim registerdOrders = repo.ToClass(repo.GetRegisteredOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, "PLAN_SET", idSelectList))
            errors.Add(ProductPlanExcelOut(IO.Path.Combine(strPath, filename), registerdOrders))
            If (CheckError(errors)) Then
                ' エラー
            End If
            ' Excel ファイル転送
            If (filename <> "") Then
                Utils.FileTransfer2(Response, Server, IO.Path.Combine(strPath, filename))
            Else
                ' Excel ファイル作成 Error
                errors.Add("Error: Excel 生産計画リスト出力エラー")
            End If

        End Sub

        '''' <summary>
        '''' 生産計画 Excel ファイル出力
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="filename"></param>
        '''' <returns></returns>
        'Private Function ProductPlanExcelOut(conn As OracleConnection, tran As OracleTransaction, filename As String) As String

        '    Return ProductPlanExcelOut(conn, tran, OrderRepository.OrdersTable.ProductPlan, filename)

        'End Function
        ''' <summary>
        ''' Orders stage の excel 出力
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="registerdOrders"></param>
        ''' <returns></returns>
        Private Function ProductPlanExcelOut(filename As String, registerdOrders As List(Of OrdersStageRow)) As String

            Dim excel = New OrderProductionPlanExcelFile()
            Return excel.OrdersStageExcelFile(filename, registerdOrders)

        End Function
        'End Function
        ''' <summary>
        ''' Orders stage の excel 出力
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="registerdOrders"></param>
        ''' <returns></returns>
        Private Function ProductPlanExcelOut(filename As String, registerdOrders As List(Of OrdersRow)) As String
            Dim path = IO.Path.GetDirectoryName(filename)
            ' フォルダがない場合作成
            If (Not IO.File.Exists(path)) Then
                EnsureDirectory(path)
            End If
            Dim excel = New OrderProductionPlanExcelFile()
            Return excel.OrdersExcelFile(filename, registerdOrders)

        End Function

        '''' <summary>
        '''' 生産計画 Excel ファイル出力
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="filename"></param>
        '''' <returns></returns>
        'Private Function ProductPlanExcelOut(conn As OracleConnection, tran As OracleTransaction, type As OrderRepository.OrdersTable, filename As String) As String

        '    Dim err As String = ""
        '    Dim customerSettingIdList As List(Of Long) = New List(Of Long)()

        '    ' 一覧で Check されている項目 の CustomerSettingId を収集
        '    For Each row In gvSelectCustomers.Rows
        '        Dim chk = TryCast(row.FindControl("chkDueDateSetting"), WebControls.CheckBox)
        '        If chk IsNot Nothing Then
        '            If chk.Checked Then
        '                Dim idx As Integer = row.RowIndex
        '                Dim keys = gvSelectCustomers.DataKeys(idx)
        '                Dim customerSettingId As Long
        '                Dim csidObj = keys("CustomerSettingId")
        '                If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), customerSettingId) Then
        '                    err += $"Row {idx}：CustomerSettingIdが不正"
        '                    Continue For
        '                End If
        '                customerSettingIdList.Add(customerSettingId)
        '            End If
        '        End If
        '    Next

        '    ' Orders（受注テーブル）
        '    ' 選択されたデータのCUSTOMER_SETTING_ID（取引先設定ID）が一致する
        '    ' STATUS（ステータス）が[PLAN_SET]で登録されているレコード
        '    Dim repo As New OrderRepository(Utils.GetConnectionString())
        '    Dim additionalConditions As String = Nothing
        '    customerSettingIdList.Select(Function(x, index) New With {Key index, .Number = x}) _
        '                         .ToList() _
        '                         .ForEach(Sub(x)
        '                                      If x.index = 0 Then
        '                                          additionalConditions &= $"AND customer_setting_id = {x.Number} " '& Environment.NewLine
        '                                      Else
        '                                          additionalConditions &= $"OR customer_setting_id = {x.Number} " '& Environment.NewLine
        '                                      End If
        '                                  End Sub)
        '    Dim registerdOrders = repo.ToClass(repo.GetOrders(conn, tran, status:="PLAN_SET", additionalConditions:=additionalConditions))
        '    Dim excel = New OrderProductionPlanExcelFile()
        '    err += excel.OrderExcelFile(filename, registerdOrders)
        '    Return err

        'End Function
        ''' <summary>
        ''' 生産計画 Excel ファイル名取得
        ''' </summary>
        ''' <returns></returns>
        Private Function GetExcelFilename() As String
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim ext = ".xlsx"
            '生産計画_[ログインユーザーID]_yyyyMMddHHmmss.xlsx
            Dim filename As String = $"生産計画_{loginUserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}"

            Return filename

        End Function

        ''' <summary>
        ''' Excel取込ボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Sub btnImportProdPlanList_Click(sender As Object, e As EventArgs) Handles Button1.Click
            'lblError.Text = "開発未着手"
            lblError.Text = ""
            lblResult.Text = ""
            Dim errors As New List(Of String)()

            ' カレントディレクトリを取得する。
            Dim url As New Uri(Server.MapPath("."))                             ' url.AbsolutePath "I:/VS2022SRC_NET/VM-KAI05/OrderManagementSystem/OrderManagementSystem/OMS.Web/Pages/Orders"
            Dim path As String = GetWorkPath()  'url.LocalPath                  ' フォルダ: ユーザーのローカルパス
            Dim fileName As String = FileUpload1.PostedFile.FileName            ' "受注フォーマット変換仕様まとめ案_20260218.xlsx"
            'Dim fullFileName As String = System.IO.Path.Combine(path, fileName) ' "I:\VS2022SRC_NET\VM-KAI05\OrderManagementSystem\OrderManagementSystem\OMS.Web\Pages\Orders\受注フォーマット変換仕様まとめ案_20260218.xlsx"

            If (fileName = "") Then
                lblError.Text = "取り込むExcelファイルを指定してください。"
                Return
            End If

            ' ファイルを保存する。
            Dim strPath = Server.MapPath("~/App_Data/Files/")
            FileUpload1.SaveAs(IO.Path.Combine(strPath, fileName))

            'If (FileUpload1.HasFile) // ファイルがアップロードされているか Then
            '                {
            '    String Path = @"C:\temp";    //保存先ディレクトリのパス指定
            '    String fileName = FileUpload1.PostedFile.FileName;    //ファイル名を取得
            '    String filePath = Path.Combine(Path, fileName);    //パス+ファイル名
            '    FileUpload1.SaveAs(filePath);    //保存を実行
            '}

            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            ' Oracle connection/Transaction
            Dim conn As New OracleConnection(Utils.GetConnectionString())
            conn.Open()
            Dim tran As OracleTransaction = conn.BeginTransaction()
            Dim repir As New ImpRunRepository(Utils.GetConnectionString())
            Dim repis As New ImpFilesStageRepository(Utils.GetConnectionString())
            Dim repif As New ImpFilesRepository(Utils.GetConnectionString())
            Dim repo As New OrderRepository(Utils.GetConnectionString())
            Dim reps As New OrderStageRepository(Utils.GetConnectionString())
            Dim reph = New OrderHistoryRepository(Utils.GetConnectionString())
            Dim importCount = 0
            Try

                '実行管理
                'IMP_RUN（取込実行テーブル）に新規レコードを作成する。
                'STARTED_AT         実行開始日時        SYSDATE
                'STATUS             ステータス          'RUNNNING'
                'STARTED_USER_ID    実行開始ユーザーID  [ログインユーザーID]
                'STARTED_PG_ID      実行プログラムID    'ProductionPlanningImport(Execute)'
                '----------------------------
                ' Insert 取込実行 IMP_RUN
                '----------------------------
#If DEBUG Then
                'Dim s = repir.GetImpRunIdNext(conn, tran)
                'repir.Delete(conn, tran, impRunId:="0")
#End If
                Dim startedAt = DateTime.Now
                Dim impRun = New ImpRunRow(startedAt, "RUNNING", loginUserId, "ProductionPlanningImport(Execute)")
                ' 次のIndexNo.
                'impRun.ImpRunId = repir.GetImpRunIdNext(conn, tran)
                errors.Add(repir.Insert(conn, tran, impRun))
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                If (CheckError(errors)) Then
                    ' エラー
                    Return
                End If
                ' 読み直し
                impRun = repir.GetImpRunRow(conn, tran, startedAt:=startedAt, startedUserId:=loginUserId)
                'impRun = repir.GetImpRun(conn, tran, startedAt:=startedAt)
                If (impRun Is Nothing) Then
                    errors.Add("Excel 取り込み ユーザー設定がありません。")
                End If
                If (CheckError(errors)) Then
                    ' エラー
                    'Return
                    Throw New Exception()
                End If

                Dim impRunId = impRun.ImpRunId
                'データインサート
                'IMP_FILES_STAGE（取込ファイルワークテーブル）に追加する。
                'CUSTOMER_SETTING_ID     取引先設定ID        9999999999                                  固定値
                'FOLDER_TYPE             フォルダ区分        4                                           固定値
                'FOLDER_PATH             フォルダパス        [フォルダパス]                              取得ファイル ?
                'FILE_NAME               ファイル名          [ファイル名]                                取得ファイル
                'STAGED_FOLDER_PATH      ワークフォルダパス  [フォルダパス]                              取得ファイル ?
                'STAGED_FILE_NAME        ワークファイル名    [ファイル名]                                取得ファイル
                'RECONCILE_FLAG          消込フラグ          'S'                                         固定値
                'FCST_RECONCILE_FLAG     内示消込フラグ      'N'                                         固定値
                'HAND_FLAG               ハンドフラグ        'Y'                                         固定値
                'STATUS                  ステータス          'PARSED'                                    固定値
                'CREATED_AT              登録日時            [処理日時]                                  OS日時
                'CREATED_USER_ID         登録ユーザーID      [ログインユーザーID]                        ログイン情報
                'CREATED_PG_ID           登録プログラムID    'ProductionPlanningImport(Execute)'         固定値
                'UPDATED_AT              更新日時            [処理日時]                                  OS日時
                'UPDATED_USER_ID         更新ユーザーID      [ログインユーザーID]                        ログイン情報
                'UPDATED_PG_ID           更新プログラムID    'ProductionPlanningImport(Execute)'         固定値
                '----------------------------
                ' Insert 取込実行 IMP_FILES_STAGE
                '----------------------------
                Dim osTime = DateTime.Now
                'path = ""
                Dim impFIlesStage = New ImpFilesStageRow(9999999999, 4, path, fileName, path, fileName, "S", "N", "Y", "PARSED", osTime, loginUserId, "ProductionPlanningImport(Execute)", osTime, loginUserId, "ProductionPlanningImport(Execute)")
                errors.Add(repis.Insert(conn, tran, impFIlesStage))
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                If (CheckError(errors)) Then
                    ' エラー
                    Throw New Exception()
                End If
                ' DB から 読み直し (ImpFileStageId 取得するため
                impFIlesStage = repis.GetImpFilesStageFilename(conn, tran, fileName:=fileName)

                'PROD_PLAN_STAGE（生産計画ワークテーブル）より、全レコードを削除する。
                '----------------------------
                ' Delete 生産計画ワーク
                '----------------------------
                'errors.Add(reps.Delete(conn, tran, OrderStageRepository.OrdersTable.ProductPlan))
                errors.Add(reps.Truncate(conn, tran, OrderStageRepository.OrdersTable.ProductPlan))
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                If (CheckError(errors)) Then
                    ' エラー
                    Throw New Exception()
                End If
                '選択されたExcelファイルからデータを取り込む。
                '----------------------------
                ' Insert 生産計画ワーク 
                '----------------------------
                Dim excel = New OrderProductionPlanExcelFile()
                Dim ordersRows = excel.OrderExcelFile(IO.Path.Combine(strPath, fileName))
                importCount = ordersRows.Count
                If (importCount = 0) Then
                    errors.Add($"{fileName}ファイルの取り込みはデータが無いため中止しました。")
                    Throw New Exception()
                End If
                errors.Add(reps.InsertRange(conn, tran, OrderStageRepository.OrdersTable.ProductPlan, ordersRows))
                If (CheckError(errors)) Then
                    ' エラー
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                '----------------------------
                ' エラーチェック
                '----------------------------
                ' 1.入力内容の真偽は判断できないため、必須チェック（NULLでないか）のみ
                ' 2.確定[ORDER_TYPE=2]、納入指示{ORDER_TYPE=3]のレコードについては
                ' CUSTOMER_ORDER_NO（客先発注No）が必須となるため、
                ' 対象の受注区分は客先発注Noを必須チェックに加える
                Dim errorCount = 0
                errorCount = ordersRows.Where(Function(x) x.IsCorrect()).Count()
                If (errorCount <> 0) Then
                    ' エラー
                    ' IMP_FILES_STAGE から追加したレコードを削除
                    ' 実行管理更新へ
                    Throw New Exception()
                End If

                '選択したExcelファイルを完了フォルダにコピーする。
                'コピーしたExcelファイルは以下の命名規則によってリネームする。
                '   [元のファイル名]_[ログインユーザーID]_yyyyMMddhhmmss.xlsx
                '   （例：[生産計画01.xlsx]→[生産計画01_ast04710_20260304164802.xlsx]）
                '----------------------------
                ' ファイルコピー
                '----------------------------
                'Dim name As String = TryCast(Page.Session("UserName"), String)
                Dim pathDone As String = IO.Path.Combine(strPath, "Copy")
                EnsureDirectory(pathDone)
                'Dim pathDone As String = GetDonePath()
                Dim newFilename As String = IO.Path.GetFileNameWithoutExtension(fileName) & loginUserId & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xlsx"
                ' コピー先に 同名ファイルがある
                If (IO.File.Exists(IO.Path.Combine(pathDone, fileName)) Or IO.File.Exists(IO.Path.Combine(pathDone, newFilename))) Then
                    ' エラー
                    'Throw New Exception()
                End If
                IO.File.Copy(IO.Path.Combine(strPath, fileName), IO.Path.Combine(pathDone, fileName))
                IO.File.Move(IO.Path.Combine(pathDone, fileName), IO.Path.Combine(pathDone, newFilename))
                '----------------------------
                ' Update IMP_FILES_STAGE 取込ファイルワークテーブル
                '----------------------------
                'STAGED_FOLDER_PATH  (ワークフォルダパス)
                'STAGED_FILE_NAME(ワークファイル名)
                'UPDATED_AT(更新日時) 
                Dim impFileStageId = impFIlesStage.ImpFileStageId
                Dim updateAt = DateTime.Now
                impFIlesStage.StagedFolderPath = pathDone
                impFIlesStage.StagedFileName = newFilename
                impFIlesStage.UpdatedAt = updateAt
                errors.Add(repis.Update(conn, tran, kImpFileStageId:=impFileStageId, stagedFolderPath:=pathDone, stagedFileName:=newFilename, updatedAt:=updateAt))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If

                'IMP_FILES（取込ファイルテーブル）データ挿入
                'IMP_FILES（取込ファイルテーブル）に、IMP_FILES_STAGE（取込ファイルワークテーブル）のレコードを追加する。
                '追加時に自動作成される IMP_FILE_ID（取込ファイルID） を保持する。
                '追加したデータはIMP_FILES_STAGE から物理削除する。
                '----------------------------
                ' Insert 取込ファイル
                '----------------------------
                ' impFIlesStage のレコードを 追加する
                Dim impFIles = ImpFilesStageRow.ToImpFilesRow(impFIlesStage)
                'impFIlesStage.Status = "PARSED"
                errors.Add(repif.Insert(conn, tran, impFIles))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                ' 追加した レコードを読み直し(imp_file_id を取得するため
                impFIles = repif.GetImpFilesFilename(conn, tran, fileName:=fileName)
                Dim impFileId = impFIles.ImpFileId
                ' 削除
                Dim impFIlesStageId = impFIlesStage.ImpFileStageId
                Dim status = impFIlesStage.Status
                errors.Add(repis.Delete(conn, tran, impFileStageId:=impFIlesStageId, status:=status))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                '----------------------------
                ' Update 生産計画ワーク
                '----------------------------
                'ORDERS_STAGE（生産計画ワークテーブル）データ更新
                'IMP_FILE_STAGE_ID(一時取込ファイルID)(取込ファイルワークテーブル)
                'IMP_FILE_ID(取込ファイルID)(取込ファイルテーブル)
                ''IMP_RUN_ID(取込実行ID)                  取込実行テーブル
                'status(ステータス)                      "PLAN_SET"
                'ACTIVE_FLAG(有効フラグ)                 "Y"
                'CREATED_AT(登録日時)                    imp_run.startedAt
                'CREATED_USER_ID(登録ユーザーID)         ログイン情報
                'CREATED_PG_ID(登録プログラムID)         "ProductionPlanningImport(Execute)"
                'UPDATED_AT(更新日時)                    imp_run.startedAt
                'UPDATED_USER_ID(更新ユーザーID)         ログイン情報
                'UPDATED_PG_ID(更新プログラムID)         "ProductionPlanningImport(Execute)"
                ' 変数
                'ordersRows
                'impRun
                'impFIlesStage
                'impFIles
                errors.Add(reps.UpdateByOrderId(conn, tran, OrderStageRepository.OrdersTable.ProductPlan,
                                     OrdersRow.ToOrdersStageRow(ordersRows),
                                     impFileStageId:=impFileStageId,
                                     impFileId:=impFileId,
                                     impRunId:=impRunId,
                                     status:="PLAN_SET",
                                     activeFlag:="Y",
                                     createdAt:=startedAt,
                                     createdUserId:=loginUserId,
                                     createdPgId:="ProductionPlanningImport(Execute)",
                                     updatedAt:=startedAt,
                                     updatedUserId:=loginUserId,
                                     updatedPgId:="ProductionPlanningImport(Execute)"))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If

                'PROD_PLAN_STAGE（生産計画ワークテーブル）のCUSTOMER_SETTING_ID（取引先設定ID）と一致するレコード
                '----------------------------
                ' Delete 生産計画
                '----------------------------
                Dim idList As New List(Of Long)
                ' 処理対象となる CustomerSettingId リスト
                ordersRows.ForEach(Sub(x) idList.Add(x.CustomerSettingId))
                ' Customer Setting ID リスト重複をまとめる
                Dim idSelectList = idList.Distinct().ToList()
                errors.Add(repo.DeleteRegisteredOrders(conn, tran, OrderRepository.OrdersTable.ProductPlan, idSelectList))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                '----------------------------
                ' Insert orders 生産計画
                '----------------------------
                'PROD_PLAN_STAGE（生産計画ワークテーブル）の全レコード
                Dim orders = OrdersStageRow.ToOrdersRow(reps.ToClass(reps.GetOrders(conn, tran, OrderStageRepository.OrdersTable.ProductPlan)))
                errors.Add(repo.InsertRange(conn, tran, OrderRepository.OrdersTable.ProductPlan, orders))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                '----------------------------
                ' Insert orders_history 生産計画履歴
                '----------------------------
                ' 履歴を登録する
                errors.Add(reph.InsertRange(conn, tran, OrderHistoryRepository.OrdersTable.ProductPlan, orders))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
                '----------------------------
                ' Update imp_run
                '----------------------------
                ' 更新条件
                'IMP_RUN_ID(取込実行ID)	処理中の IMP_RUN_ID
                'status(ステータス)		"RUNNING"
                'ENDED_AT(実行終了日時)      DateTIme.Now
                'status(ステータス)          OK:"COMPLETED" NG:"FAILED"
                'FILE_COUNT(ファイル件数)    1
                'ROW_COUNT(データ件数)
                'ERROR_COUNT(エラー件数)

                Dim endedAt = DateTime.Now
                Dim satus = If(errorCount = 0, "COMPLETED", "FAILED")
                Dim fileCount = 1
                errors.Add(repir.Update(conn, tran, kImpRunId:=impRunId, kStatus:="RUNNING", endedAt:=endedAt, status:=status, fileCount:=fileCount, rowCount:=importCount, errorCount:=errorCount))
                If (CheckError(errors)) Then
                    ' エラー 
                    Throw New Exception()
                End If
                '#If DEBUG Then
                '            ' #### DEBUG
                '            tran.Commit()
                '            tran = conn.BeginTransaction()
                '            ' #### DEBUG
                '#End If
            Catch ex As Exception
            Finally
                If (errors.Count > 0) Then
                    'lblError.Text = errors(0)
                    lblError.Text = String.Join(vbCrLf, errors)
                Else
                    If (importCount = 0) Then
                    Else
                        lblResult.Text = $"{fileName}ファイルの取り込みを行いました。"
                    End If
                End If
            End Try

        End Sub
#End Region
        ''' <summary>
        ''' 指定日付の 月曜日 休日の場合は 翌営業日 を取得する
        ''' </summary>
        ''' <param name="calender"></param>
        ''' <param name="startDate">月曜 指定 週先頭日付</param>
        ''' <returns>Error : DateTime.MinValue</returns>
        Public Function GetDateWeekBusinessDay(calender As List(Of CalenderRow), startDate As Date) As Date

            ' その週で営業日のリストを抽出
            Dim thisMonday = GetDateMonday(startDate)
            Dim businessDays = calender.Where(Function(x) x.DefDate >= DateSerial(thisMonday.Year, thisMonday.Month, thisMonday.Day) And x.DefDate <= DateSerial(startDate.Year, startDate.Month, startDate.Day + 7) And x.HolidayFlag = "W").OrderBy(Function(x) x.DefDate).ToList()
            'Dim businessDays = calender.Where(Function(x) x.DefDate <= DateSerial(startDate.Year, startDate.Month, startDate.Day + 7) And x.HolidayFlag = "W").OrderBy(Function(x) x.DefDate).ToList()
            Dim dt As Date = DateTime.MinValue
            If (businessDays.Count <> 0) Then
                ' 最も早い営業日を返す
                dt = businessDays(0).DefDate
            End If
            Return dt
        End Function

        ''' <summary>
        ''' 指定日付の 月曜日 日付を取得する
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Public Function GetDateMonday(tgDate As Date) As Date

            ' 月曜日を基準とした差分を計算
            ' today.DayOfWeek: Sunday(0), Monday(1), ..., Saturday(6)
            ' 月曜日(1)を基準にするため、月曜日を起点(0)とした差分を出す
            Dim diff = (7 + tgDate.DayOfWeek - DayOfWeek.Monday) Mod 7

            ' 今日の日付から差分を引く
            Dim thisMonday = tgDate.AddDays(-diff)

            'Console.WriteLine($"今日の曜日: {Today.DayOfWeek}")
            'Console.WriteLine($"今週の月曜日: {thisMonday:yyyy/MM/dd}")
            Return thisMonday

        End Function

        ''' <summary>
        ''' 4週前の日付を取得する
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Public Function Get4WeeksBefore(tgDate As Date) As Date
            Return GetDaysBefore(tgDate, 7 * 4)
        End Function

        ''' <summary>
        ''' 1週前の日付を取得する
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Public Function Get1WeeksBefore(tgDate As Date) As Date
            Return GetDaysBefore(tgDate, 7)
        End Function

        ''' <summary>
        ''' 指定日数前の 日付を取得
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <param name="bday"></param>
        ''' <returns></returns>
        Public Function GetDaysBefore(tgDate As Date, bday As Integer) As Date

            Dim daysBefore As Date

            daysBefore = tgDate.AddDays(-bday)

            Return daysBefore
        End Function

        ''' <summary>
        ''' 先月 第4週の月曜日を取得する
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Public Function Get4WeekOfLastMonth(tgDate As Date) As Date
            Dim firstOfThisMonth As Date
            Dim firstOfLastMonth As Date
            Dim fourthMonday As Date

            ' 1. 今月1日の日付を取得
            firstOfThisMonth = DateSerial(tgDate.Year, tgDate.Month, 1)

            ' 2. 先月1日の日付を取得
            firstOfLastMonth = DateAdd("m", -1, firstOfThisMonth)

            ' 3. 先月1日を基準に、第4月曜日を計算
            ' Weekday(日付, vbMonday) は月曜を1とする数値を返す
            fourthMonday = firstOfLastMonth.AddDays(((8 - Weekday(firstOfLastMonth, vbMonday)) Mod 7) + 21)

            'Debug.Print($"先月の第4月曜日は: {fourthMonday}")

            ' 第4週（月〜日）の日付を表示する場合
            'Dim i As Integer
            'For i = 0 To 6
            'Debug.Print($"第4週の日付 { i + 1 } 日目: {fourthMonday.AddDays(i)}")
            'Next i
            Return fourthMonday
        End Function

        ''' <summary>
        ''' 指定 index の 日付を取得する
        ''' </summary>
        ''' <param name="calender"></param>
        ''' <param name="holidayFlag"></param>
        ''' <param name="index"></param>
        ''' <returns></returns>
        Public Function GetDesignationDay(calender As List(Of CalenderRow), holidayFlag As String, index As Integer) As Date

            Dim calenderList As List(Of CalenderRow) = calender.FindAll(Function(x) x.HolidayFlag = holidayFlag)
            'If (index <= 0) Then
            'index = 1
            'End If
            'If (calenderList.Count < index - 1) Then
            If (calenderList.Count < index + 1) Then
                ' 範囲外エラー
            End If
            Return calenderList(index).DefDate

            'Return calenderList(index - 1).DefDate

        End Function

        ''' <summary>
        ''' 当月第3週 週末の日付取得
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Public Function Get3WeekOfThisMonth(tgDate As Date) As Date
            Dim firstOfThisMonth As Date

            ' 当月の1日を取得
            firstOfThisMonth = DateSerial(tgDate.Year, tgDate.Month, 1)

            ' 第1週の土曜日を求める (DayOfWeek.Saturday = 6)
            Dim daysUntilSaturday As Integer = (DayOfWeek.Saturday - firstOfThisMonth.DayOfWeek + 7) Mod 7
            Dim firstSaturday As DateTime = firstOfThisMonth.AddDays(daysUntilSaturday)

            ' 第3週の終わりの日（第1土曜日の2週間後）
            Dim thirdWeekEnd As DateTime = firstSaturday.AddDays(14)
            'Console.WriteLine("今月の第3週の終わり（土曜日）は: " & thirdWeekEnd.ToShortDateString())
            Return thirdWeekEnd

        End Function
        ''' <summary>
        ''' 指定月の週数を取得 
        ''' </summary>
        ''' <param name="tgDate"></param>
        ''' <returns></returns>
        Function GetWeeksInMonth(tgDate As Date) As Integer
            Return GetWeeksInMonth(tgDate.Year, tgDate.Month)
        End Function

        ''' <summary>
        ''' 指定月の週数を取得 
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <returns></returns>
        Function GetWeeksInMonth(ByVal year As Integer, ByVal month As Integer) As Integer
            Dim firstDay As Date
            Dim lastDay As Date
            Dim firstWeekNum As Integer
            Dim lastWeekNum As Integer

            ' 月の初日
            firstDay = DateSerial(year, month, 1)
            ' 月の最終日（翌月1日から1日引く） [8, 11]
            lastDay = DateSerial(year, month + 1, 0)

            ' 第1週と最終週の週番号を取得（月曜始まり: vbMonday） [2, 5, 10]
            firstWeekNum = DatePart("ww", firstDay, vbMonday)
            lastWeekNum = DatePart("ww", lastDay, vbMonday)

            ' 週数を計算
            Return lastWeekNum - firstWeekNum + 1

        End Function
        ''' <summary>
        ''' 指定期間で 指定数を丸め単位で格納する 計算(3分割)
        ''' </summary>
        ''' <param name="total"></param>
        ''' <param name="tdays"></param>
        ''' <param name="tunit"></param>
        ''' <returns>
        '''  roundedUp 丸め単位
        '''  days 展開日数
        '''  fraction 端数
        ''' </returns>
        Public Function GetRoundCalc3D(total As Integer, tdays As Integer, tunit As Integer) As (roundedUp As Integer, days As Integer, fraction As Integer)
            Dim totalQuantity As Double = total
            Dim daysValue As Double = 3 'tdays
            Dim unit As Double = tunit

            ' 1日あたりの平均
            Dim dailyAvg As Double = totalQuantity / daysValue

            ' tunit単位で切り上げ計算
            'Dim roundedUpValue As Double = Math.Ceiling(dailyAvg / 3) * 3
            Dim roundedUpValue As Double = Math.Ceiling(dailyAvg / unit) * unit

            ' 割当日数と端数の計算
            Dim dcnt As Integer = total \ CInt(roundedUpValue) ' 「\」は整数除算（余り切り捨て）
            Dim fraction As Integer = total - (CInt(roundedUpValue) * dcnt)

            ' 端数がある場合は日数を加算                            
            If fraction <> 0 Then
                dcnt += 1
            End If
            Return (CInt(roundedUpValue), dcnt, fraction)

        End Function

        ''' <summary>
        ''' 指定期間で 指定数を丸め単位で格納する 計算
        ''' </summary>
        ''' <param name="total">総需要数</param>
        ''' <param name="tdays">分割数</param>
        ''' <param name="tunit">まるめ数</param>
        ''' <param name="splitRationType">分割比 1:全数-分割割合 2:先頭もしくは後端x2</param>
        ''' <returns>
        '''  roundedUp 丸め単位
        '''  days 展開日数
        '''  fraction 端数
        ''' </returns>
        Public Function GetRoundCalc(total As Integer, tdays As Integer, tunit As Integer, splitRationType As Int16) As (roundedUp As Integer, days As Integer, fraction As Integer)
            ' splitRationType条件設定
            Dim daysValue As Double = tdays + If(splitRationType = 1, 0, 1)
            Dim totalQuantity As Double = total
            Dim unit As Double = tunit
            Dim modCount = total Mod daysValue
            'If (modCount <> 0) Then
            'daysValue -= 1
            'End If

            ' 1日あたりの平均
            Dim dailyAvg As Double = totalQuantity / daysValue

            ' tunit単位で切り上げ計算
            Dim roundedUpValue As Double = Math.Ceiling(dailyAvg / tunit) * tunit

            ' 割当日数と端数の計算
            Dim dcnt As Integer = total \ CInt(roundedUpValue) ' 「\」は整数除算（余り切り捨て）
            Dim fraction As Integer = total - (CInt(roundedUpValue) * dcnt)

            ' 端数がある場合は日数を加算
            If fraction <> 0 Then
                dcnt += 1
            End If

            If (modCount <> 0 And splitRationType = 2) Then
                dcnt -= 1
            End If

            Return (CInt(roundedUpValue), dcnt, fraction)
        End Function

        Protected Sub btnConfirm_Click(sender As Object, e As EventArgs)
            ' HiddenField の値を取得
            'Dim userResponse As String = hfResult.Value

            'If userResponse = "true" Then
            '    ' OK が押された時の処理
            '    'lblStatus.Text = "ユーザーは「OK」を選択しました。処理を実行します。"
            '    'lblStatus.ForeColor = System.Drawing.Color.Blue
            'Else
            '    ' キャンセル が押された時の処理
            '    'lblStatus.Text = "ユーザーは「キャンセル」を選択しました。処理を中断しました。"
            '    'lblStatus.ForeColor = System.Drawing.Color.Red
            'End If

            ' 次回の操作のために値をリセット（必要に応じて）
            'hfResult.Value = ""
        End Sub

        ''' <summary>
        ''' Work path 取得 (ない場合は作成する)
        ''' </summary>
        ''' <returns></returns>
        Public Function GetWorkPath() As String

            Dim rawWork As String = GetWorkFolderRoot()
            Dim workFolder = rawWork & "FileTempFolder"
            If (Not IO.File.Exists(workFolder)) Then
                EnsureDirectory(workFolder)
            End If
            Return workFolder
        End Function

        ''' <summary>
        ''' 完了 path 取得 (ない場合は作成する)
        ''' </summary>
        ''' <returns></returns>
        Public Function GetDonePath() As String

            Dim rawWork As String = GetDoneFolderRoot()
            Dim workFolder = rawWork & "FileTempFolder"
            If (Not IO.File.Exists(workFolder)) Then
                EnsureDirectory(workFolder)
            End If
            Return workFolder
        End Function

    End Class
End Namespace