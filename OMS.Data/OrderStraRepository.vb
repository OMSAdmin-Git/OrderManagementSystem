

Imports System.Configuration
Imports System.Data
Imports System.Text
Imports System.Web
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.SqlServer
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports DataTable = System.Data.DataTable

Namespace OMS.Data

    'OrderStraRepository

    ''' <summary>
    ''' 生産計画出力一覧 Excel ファイル出力
    ''' </summary>
    Public Class OrderStraRepository

        Implements IDisposable
        Private ReadOnly _connectionString As String
        Public Sub New()

        End Sub

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub


        Public strErrMsg As String

        ''' <summary>
        ''' IDisposable 実装
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            'Throw New NotImplementedException()
        End Sub

        'Public Shared Function ShippingStatusErrorListExcelFile(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Long) As String

        '    Dim filename = GetShippingStatusErrorListExcelFilename()
        '    Dim index = 1
        '    Using excelfile As New ShippingStatusErrorListExcelFile()
        '        If (excelfile.ErrorListExcelFile(conn, tran, filename, customerSettingId)) Then
        '        Else
        '            filename = ""
        '        End If
        '    End Using

        '    Return filename

        'End Function

        'Private Shared Function GetShippingStatusErrorListExcelFilename() As String
        '    Dim filename = ""
        '    Return filename
        'End Function

        '''' <summary>
        '''' PROD_PLAN_STRA_VIEW（生産計画出力一覧）をExcel出力する。
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="filename"></param>
        '''' <param name="customerSettingId"></param>
        '''' <returns></returns>
        'Public Function ErrorListExcelFile(conn As OracleConnection, tran As OracleTransaction, filename As String, customerSettingId As Long) As String




        '    Dim rows = GetErrorList(conn, tran, customerSettingId:=customerSettingId, status:="customerSettingId", activeFlag:="N")
        '    Dim orderStraRows = ToClass(rows)



        'End Function

        Public Function GetErrorList(conn As OracleConnection, tran As OracleTransaction,
                                        Optional ByVal customerSettingId As Long? = Nothing,
                                        Optional ByVal status As String = Nothing,
                                        Optional ByVal activeFlag As String = Nothing) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            Dim prm As New List(Of OracleParameter)()
            sb.AppendLine("SELECT * ")
            sb.AppendLine($"FROM prod_plan_stra_view ")
            sb.AppendLine("WHERE 1=1 ")

            If customerSettingId IsNot Nothing Then
                sb.AppendLine("AND customer_setting_id = :p_csid ")
                prm.Add(New OracleParameter(":p_customerSettingId", OracleDbType.Int64) With {.Value = customerSettingId})
            End If

            If status IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If Not String.IsNullOrEmpty(activeFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = :p_active ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = activeFlag})
            End If

            Using cmd As New OracleCommand(sb.ToString(), conn)
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt

        End Function
        Public Function GetOrderStras(conn As OracleConnection, tran As OracleTransaction,
                                        Optional ByVal customerSettingId As Long? = Nothing,
                                        Optional ByVal demandStatus As String = Nothing,
                                        Optional ByVal status As String = Nothing) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            Dim prm As New List(Of OracleParameter)()
            sb.AppendLine("SELECT * ")
            sb.AppendLine($"FROM prod_plan_stra_view ")
            'sb.AppendLine($"FROM prod_plan ")
            sb.AppendLine("WHERE 1=1 ")

            If customerSettingId IsNot Nothing Then
                sb.AppendLine("AND customer_setting_id = :p_csid ")
                prm.Add(New OracleParameter(":p_customerSettingId", OracleDbType.Int64) With {.Value = customerSettingId})
            End If

            If status IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If demandStatus IsNot Nothing Then
                sb.AppendLine("AND UPPER(demand_status) LIKE UPPER(:p_demandStatus) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_demandStatus", OracleDbType.Varchar2) With {.Value = demandStatus})
            End If

            'If Not String.IsNullOrEmpty(activeFlag) Then
            '    sb.AppendLine("AND UPPER(active_flag) = :p_active ")
            '    prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = activeFlag})
            'End If

            Using cmd As New OracleCommand(sb.ToString(), conn)
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt

        End Function

        Public Function GetOrderStage(conn As OracleConnection, tran As OracleTransaction,
                                        Optional ByVal customerSettingId As Long? = Nothing,
                                        Optional ByVal demandStatus As String = Nothing,
                                        Optional ByVal status As String = Nothing,
                                        Optional ByVal activeFlag As String = Nothing) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            Dim prm As New List(Of OracleParameter)()
            sb.AppendLine("SELECT * ")
            sb.AppendLine($"FROM prod_plan_stage_view ")
            sb.AppendLine("WHERE 1=1 ")

            If customerSettingId IsNot Nothing Then
                sb.AppendLine("AND customer_setting_id = :p_csid ")
                prm.Add(New OracleParameter(":p_customerSettingId", OracleDbType.Int64) With {.Value = customerSettingId})
            End If

            If status IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If demandStatus IsNot Nothing Then
                sb.AppendLine("AND UPPER(demand_status) LIKE UPPER(:p_demandStatus) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_demandStatus", OracleDbType.Varchar2) With {.Value = demandStatus})
            End If

            If Not String.IsNullOrEmpty(activeFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = :p_active ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = activeFlag})
            End If

            Using cmd As New OracleCommand(sb.ToString(), conn)
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt

        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As OrderStageViewRow

            If dt Is Nothing Then
                Return Nothing
            End If
            Dim osr = New OrderStageViewRow
            osr.StageID = If(dt.Field(Of Long?)("stage_id"), 0)
            'osr.StageID = dt.Field(Of Long?)("stage_id")
            osr.ProdPlanID = If(dt.Field(Of Long?)("prod_plan_id"), 0)
            'osr.ProdPlanID = dt.Field(Of Long?)("prod_plan_id")
            osr.CustomerSettingID = If(dt.Field(Of Long?)("customer_setting_id"), 0)
            'osr.CustomerSettingID = dt.Field(Of Long?)("customer_setting_id")
            osr.CustomerCode = dt.Field(Of String)("customer_code")
            osr.CustomerName = dt.Field(Of String)("customer_name")
            osr.ProfitCenter = dt.Field(Of String)("profit_center")
            osr.CustomerUnitID = If(dt.Field(Of Long?)("customer_unit_id"), 0)
            'osr.CustomerUnitID = dt.Field(Of Long?)("customer_unit_id")
            osr.CustomerUnitName = dt.Field(Of String)("customer_unit_name")
            osr.BillingTo = dt.Field(Of String)("billing_to")
            osr.CustomerOrderNo = dt.Field(Of String)("customer_order_no")
            osr.DemandStatus = dt.Field(Of String)("demand_status")
            osr.ShipTo = dt.Field(Of String)("ship_to")
            osr.OrderDate = dt.Field(Of Date?)("order_date")
            osr.DueDate = dt.Field(Of Date?)("due_date")
            osr.ShipScheduledDate = dt.Field(Of Date?)("ship_scheduled_date")
            osr.CustomerItemNo = dt.Field(Of String)("customer_item_no")
            osr.ItemNo = dt.Field(Of String)("item_no")
            osr.DemandQty = If(dt.Field(Of Long?)("demand_qty"), 0)
            'osr.DemandQty = dt.Field(Of Long?)("demand_qty")
            osr.DemandUnit = dt.Field(Of String)("demand_unit")
            osr.CurrencyCode = dt.Field(Of String)("currency_code")
            osr.ShipStockLocation = dt.Field(Of String)("ship_stock_location")
            osr.CompanyId = dt.Field(Of String)("company_id")
            osr.ProductCode = dt.Field(Of String)("product_code")
            osr.BillingStandard = dt.Field(Of String)("billing_standard")
            osr.ShipProcessType = dt.Field(Of String)("ship_process_type")
            osr.DeliveryInstrFlag = dt.Field(Of String)("delivery_instr_flag")
            osr.OrderNo = dt.Field(Of String)("order_no")
            osr.Remarks = dt.Field(Of String)("remarks")
            osr.DeliveryCode = dt.Field(Of String)("delivery_code")
            osr.TotalShipQty = If(dt.Field(Of Decimal?)("total_ship_qty"), 0)
            'osr.TotalShipQty = dt.Field(Of Decimal?)("total_ship_qty")
            osr.ShipDate = dt.Field(Of Date?)("ship_date")
            osr.TransportMethod = dt.Field(Of String)("transport_method")
            osr.ShipPlanDate = dt.Field(Of Date?)("ship_plan_date")
            osr.CustomerOrderLineNo = dt.Field(Of String)("customer_order_line_no")
            osr.PreDailyOrderQty = If(dt.Field(Of Decimal?)("pre_daily_order_qty"), 0)
            'osr.PreDailyOrderQty = dt.Field(Of Decimal?)("pre_daily_order_qty")
            osr.PreDailyDeliveryDate = dt.Field(Of Date?)("pre_daily_delivery_date")
            osr.ImpFileStageID = If(dt.Field(Of Long?)("imp_file_stage_id"), 0)
            'osr.ImpFileStageID = dt.Field(Of Long?)("imp_file_stage_id")
            osr.ImpFileID = If(dt.Field(Of Long?)("imp_file_id"), 0)
            'osr.ImpFileID = dt.Field(Of Long?)("imp_file_id")
            osr.OrderType = If(dt.Field(Of Int16?)("order_type"), 0)
            'osr.OrderType = dt.Field(Of Int16?)("order_type")
            osr.ProratedType = If(dt.Field(Of Int16?)("prorated_type"), 0)
            'osr.ProratedType = dt.Field(Of Int16?)("prorated_type")
            osr.CustomerInfoType = dt.Field(Of String)("customer_info_type")
            osr.InfoType = dt.Field(Of String)("info_type")
            osr.SelfFcstFlag = dt.Field(Of String)("self_fcst_flag")
            osr.SelfFcstDeleteFlag = dt.Field(Of String)("self_fcst_delete_flag")
            osr.ReconcilType = If(dt.Field(Of Int16?)("reconcile_type"), 0)
            'osr.ReconcilType = dt.Field(Of Int16?)("reconcile_type")
            osr.ImpRunID = If(dt.Field(Of Long?)("imp_run_id"), 0)
            'osr.ImpRunID = dt.Field(Of Long?)("imp_run_id")
            osr.Status = dt.Field(Of String)("status")
            osr.ActiveFlag = dt.Field(Of String)("active_flag")
            osr.CreatedAt = dt.Field(Of Date?)("created_at")
            osr.CreatedUserID = dt.Field(Of String)("created_user_id")
            osr.CreatedPGID = dt.Field(Of String)("created_pg_id")
            osr.UpdatedAt = dt.Field(Of Date?)("updated_at")
            osr.UpdatedUserID = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPGID = dt.Field(Of String)("updated_pg_id")
            osr.ProdMgmtUserID = dt.Field(Of String)("prod_mgmt_user_id")
#If False Then
            osr.CustomerCode = dt.Field(Of String)("customer_code")
            osr.BillingTo = dt.Field(Of String)("billing_to")
            osr.CustomerOrderNo = dt.Field(Of String)("customer_order_no")
            osr.DemandStatus = dt.Field(Of String)("demand_status")
            osr.ShipTo = dt.Field(Of String)("ship_to")
            osr.OrderDate = dt.Field(Of Date?)("order_date")
            osr.DueDate = dt.Field(Of Date?)("due_date")
            osr.ShipScheduledDate = dt.Field(Of Date?)("ship_scheduled_date")
            osr.CustomerItemNo = dt.Field(Of String)("customer_item_no")
            osr.ItemNo = dt.Field(Of String)("item_no")
            osr.DemandQty = dt.Field(Of Long?)("demand_qty")
            osr.DemandUnit = dt.Field(Of String)("demand_unit")
            osr.CurrencyCode = dt.Field(Of String)("currency_code")
            osr.ShipStockLocation = dt.Field(Of String)("ship_stock_location")
            osr.CompanyId = dt.Field(Of String)("company_id")
            osr.ProductCode = dt.Field(Of String)("product_code")
            osr.BillingStandard = dt.Field(Of String)("billing_standard")
            osr.ShipProcessType = dt.Field(Of String)("ship_process_type")
            osr.DeliveryInstrFlag = dt.Field(Of String)("delivery_instr_flag")
            osr.OrderNo = dt.Field(Of String)("order_no")
            osr.Remarks = dt.Field(Of String)("remarks")
            osr.DeliveryCode = dt.Field(Of String)("delivery_code")
            'osr.OrderTime = dt.Field(Of Decimal?)("order_time")
            osr.SalesUnitPrice = dt.Field(Of Decimal?)("sales_unit_price")
            osr.DeliveryTime = dt.Field(Of Decimal?)("delivery_time")
            osr.UsageLocation = dt.Field(Of String)("usage_location")
            osr.TotalShipQty = dt.Field(Of Decimal?)("total_ship_qty")
            osr.ProductionCategory = dt.Field(Of String)("production_category")
            osr.Char2 = dt.Field(Of String)("char_2")
            osr.ContainerNo = dt.Field(Of String)("container_no")
            osr.Char3 = dt.Field(Of String)("char_3")
            osr.Char4 = dt.Field(Of String)("char_4")
            osr.Char42 = dt.Field(Of String)("char_4_2")
            osr.Char5 = dt.Field(Of String)("char_5")
            osr.Char52 = dt.Field(Of String)("char_5_2")
            osr.Char6 = dt.Field(Of String)("char_6")
            osr.OrderReason = dt.Field(Of String)("order_reason")
            osr.ContainerCapacity = dt.Field(Of Decimal?)("container_capacity")
            osr.CustomerLotNo = dt.Field(Of String)("customer_lot_no")
            osr.InitialFlag = dt.Field(Of String)("initial_flag")
            osr.ShipDate = dt.Field(Of Date?)("ship_date")
            osr.Char50 = dt.Field(Of String)("char_50")
            osr.TransportMethod = dt.Field(Of String)("transport_method")
            osr.ShipPlanDate = dt.Field(Of Date?)("ship_plan_date")
            osr.CustomerOrderLineNo = dt.Field(Of String)("customer_order_line_no")
            osr.PreDailyOrderQty = dt.Field(Of Decimal?)("pre_daily_order_qty")
            osr.PreDailyDeliveryDate = dt.Field(Of Date?)("pre_daily_delivery_date")
#End If
            Return osr

        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As List(Of OrderStageViewRow)

            Dim osrs = New List(Of OrderStageViewRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

    End Class

    ''' <summary>
    ''' 出荷状況 エラー
    ''' 
    ''' </summary>
    Public Class OrderStageViewRow
        Public Property StageID As Long?                'ステージID NUMBER(10)
        Public Property ProdPlanID As Long?             '生産計画ID NUMBER(10)
        Public Property CustomerSettingID As Long?      '客先ID     NUMBER(10)
        Public Property CustomerCode As String          '取引先コード
        Public Property CustomerName As String          '取引先名
        Public Property ProfitCenter As String          'PC
        Public Property CustomerUnitID As Long?         '取引先ユニットID NUMBER(10)
        Public Property CustomerUnitName As String      '取引先ユニット
        Public Property BillingTo As String             '請求先
        Public Property CustomerOrderNo As String       '客先発注No
        Public Property DemandStatus As String          '需要ステイタス
        Public Property ShipTo As String                '出荷先
        Public Property OrderDate As Date?              '受注日
        Public Property DueDate As Date?                '希望納期
        Public Property ShipScheduledDate As Date?      '出荷予定日
        Public Property CustomerItemNo As String        '客先品目No
        Public Property ItemNo As String                '品目No
        Public Property DemandQty As Long?              '需要数
        Public Property DemandUnit As String            '需要単位
        Public Property CurrencyCode As String          '通貨コード
        Public Property ShipStockLocation As String     '出荷在庫場所
        Public Property CompanyId As String             '会社ID
        Public Property ProductCode As String           '製品コード
        Public Property BillingStandard As String       '請求基準
        Public Property ShipProcessType As String       '出荷プロセスタイプ
        Public Property DeliveryInstrFlag As String     '納入指示フラグ
        Public Property OrderNo As String               '受注番号
        Public Property Remarks As String               'コメント
        Public Property DeliveryCode As String          '納入先コード
        Public Property TotalShipQty As Decimal?        '累計出荷数
        Public Property ShipDate As Date?               '出荷日
        Public Property TransportMethod As String       '輸送方法
        Public Property ShipPlanDate As Date?           '出荷計画日
        Public Property CustomerOrderLineNo As String   '客先発注No行番号
        Public Property PreDailyOrderQty As Decimal?    '日割前受注数
        Public Property PreDailyDeliveryDate As Date?   '日割前納期
        Public Property ImpFileStageID As Long?         'ImpFileStageID
        Public Property ImpFileID As Long?              'ImpFileID
        Public Property OrderType As Int16?             'OrderType
        Public Property ProratedType As Int16?          'ProratedType
        Public Property CustomerInfoType As String      'CustomerInfoType
        Public Property InfoType As String              'InfoType
        Public Property SelfFcstFlag As String          'SelfFcstFlag
        Public Property SelfFcstDeleteFlag As String    'SelfFcstDeleteFlag
        Public Property ReconcilType As Int16?          'ReconcilType
        Public Property ImpRunID As Long?               'ImpRunID
        Public Property Status As String                'Status
        Public Property ActiveFlag As String            'ActiveFlag
        Public Property CreatedAt As Date?              'CreatedAt
        Public Property CreatedUserID As String         'CreatedUserID
        Public Property CreatedPGID As String           'CreatedPGID
        Public Property UpdatedAt As Date?              'UpdatedAt
        Public Property UpdatedUserID As String         'UpdatedUserID
        Public Property UpdatedPGID As String           'UpdatedPGID
        Public Property ProdMgmtUserID As String        'ProdMgmtUserID
#If False Then
        Public Property CustomerCode As String          '取引先コード
        Public Property BillingTo As String             '請求先
        Public Property CustomerOrderNo As String       '客先発注No
        Public Property DemandStatus As String          '需要ステイタス
        Public Property ShipTo As String                '出荷先
        Public Property OrderDate As Date?              '受注日
        Public Property DueDate As Date?                '希望納期
        Public Property ShipScheduledDate As Date?      '出荷予定日
        Public Property CustomerItemNo As String        '客先品目No
        Public Property ItemNo As String                '品目No
        Public Property DemandQty As Long?              '需要数
        Public Property DemandUnit As String            '需要単位
        Public Property CurrencyCode As String          '通貨コード
        Public Property ShipStockLocation As String     '出荷在庫場所
        Public Property CompanyId As String             '会社ID
        Public Property ProductCode As String           '製品コード
        Public Property BillingStandard As String       '請求基準
        Public Property ShipProcessType As String       '出荷プロセスタイプ
        Public Property DeliveryInstrFlag As String     '納入指示フラグ
        Public Property OrderNo As String               '受注番号
        Public Property Remarks As String               'コメント
        Public Property DeliveryCode As String          '納入先コード
        'Public Property OrderTime As Decimal?           '受注時刻
        Public Property SalesUnitPrice As Decimal?      '売上単価
        Public Property DeliveryTime As Decimal?        '納入時間
        Public Property UsageLocation As String         '使用先
        Public Property TotalShipQty As Decimal?        '累計出荷数
        Public Property ProductionCategory As String    '生産区分
        Public Property Char2 As String                 '文字 2ｹﾀ
        Public Property ContainerNo As String           '容器番号
        Public Property Char3 As String                 '文字 3ｹﾀ
        Public Property Char4 As String                 '文字 4ｹﾀ
        Public Property Char42 As String                '文字 4ｹﾀ
        Public Property Char5 As String                 '文字 5ｹﾀ
        Public Property Char52 As String                '文字 5ｹﾀ
        Public Property Char6 As String                 '文字 6ｹﾀ
        Public Property OrderReason As String           '発注理由
        Public Property ContainerCapacity As Decimal?   '容器収容数
        Public Property CustomerLotNo As String         '得意先 ﾛｯﾄNO
        Public Property InitialFlag As String           '初品区分
        Public Property ShipDate As Date?               '出荷日
        Public Property Char50 As String                '文字 50ｹﾀ
        Public Property TransportMethod As String       '輸送方法
        Public Property ShipPlanDate As Date?           '出荷計画日
        Public Property CustomerOrderLineNo As String   '客先発注No行番号
        Public Property PreDailyOrderQty As Decimal?    '日割前受注数
        Public Property PreDailyDeliveryDate As Date?   '日割前納期
#End If
    End Class

End Namespace

