
Imports System.Data
Imports System.Text
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Spreadsheet
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports SixLabors.Fonts

Namespace OMS.Data


    ' Order_Stage と　Order_History の違いは キーの名称と　Order_Stageに 一時取り込みファイルID があること


    Public Class OrderHistoryRepository

        Public Enum OrdersTable
            Orders = 1
            ProductPlan = 2
        End Enum

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
        ''' <summary>
        ''' type に 応じたTable 名称を返す
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Private Function GetTableName(type As OrdersTable)

            Dim tableNames As New List(Of (id As OrdersTable, name As String)) From {
                (OrdersTable.Orders, "orders_history"),
                (OrdersTable.ProductPlan, "prod_plan_history")
            }
            Dim name = tableNames.Find(Function(x) x.id = type).name
            Return name

        End Function
        ''' <summary>
        ''' type に 応じたTableId 名称を返す
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Private Function GetIdName(type As OrdersTable)

            Dim tableNames As New List(Of (id As OrdersTable, name As String)) From {
                (OrdersTable.Orders, "order_id"),
                (OrdersTable.ProductPlan, "prod_plan_id")
            }
            Dim name = tableNames.Find(Function(x) x.id = type).name
            Return name

        End Function
        '''' <summary>
        '''' OrderRow class を Order History DBに追加する
        '''' R.sagisaka Add
        '''' </summary>
        '''' <param name="row"></param>
        'Public Sub Insert(row As OrdersRow)
        '    Dim records = New List(Of OrdersRow)()
        '    records.Add(row)
        '    InsertRange(records)
        'End Sub
        '''' <summary>
        '''' OrderRow class を Order History DBに追加する
        '''' R.sagisaka Add
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="row"></param>
        '''' <returns>string: Number:xxxx Errormessage</returns>
        'Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As OrdersRow) As String

        '    Dim records = New List(Of OrdersRow)()
        '    records.Add(row)
        '    Return InsertRange(conn, tran, records)

        'End Function

        '''' <summary>
        '''' OrderRow class リストをDBに追加する (元のコード同等呼び出し
        '''' R.sagisaka Add
        '''' </summary>
        '''' <param name="records"></param>
        'Public Sub InsertRange(records As IEnumerable(Of OrdersRow))

        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            InsertRange(conn, tran, records)
        '            tran.Commit()
        '        End Using
        '    End Using
        'End Sub
        ''' <summary>
        ''' Order record 追加
        ''' R.sagisaka Modified
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As IEnumerable(Of OrdersRow)) As String

            Return InsertRange(conn, tran, OrdersTable.Orders, records)

        End Function
        ''' <summary>
        ''' Order record 追加
        ''' R.sagisaka Modified
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, records As List(Of OrdersRow)) As String
            Dim errorMessage As String = ""
            If records Is Nothing Then Return errorMessage
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO {GetTableName(type)} (")
                sb.AppendLine($" {GetIdName(type)}, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, ")
                sb.AppendLine("  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, ")
                sb.AppendLine("  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, ")
                sb.AppendLine("  product_code, billing_standard, ship_process_type, delivery_instr_flag, ")
                sb.AppendLine("  order_no, remarks, delivery_code, total_ship_qty, ship_date, ")
                sb.AppendLine("  transport_method, ship_plan_date, customer_order_line_no, ")
                sb.AppendLine("  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, ")
                sb.AppendLine("  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, ")
                sb.AppendLine("  reconcile_type, imp_run_id, status, active_flag, ")
                sb.AppendLine("  created_at, created_user_id, created_pg_id, ")
                sb.AppendLine("  updated_at, updated_user_id, updated_pg_id ")
                If (type = OrdersTable.Orders) Then
                    sb.AppendLine(",  stra_order_qty, stra_ship_qty, stra_order_backlog ")
                End If
                sb.AppendLine(" ) VALUES (")
                sb.AppendLine("  :p_order_id, :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, ")
                sb.AppendLine("  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, ")
                sb.AppendLine("  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, ")
                sb.AppendLine("  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, ")
                sb.AppendLine("  :p_order_no, :p_remarks, :p_delivery_code, :p_total_ship_qty, :p_ship_date, ")
                sb.AppendLine("  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, ")
                sb.AppendLine("  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, ")
                sb.AppendLine("  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag, ")
                sb.AppendLine("  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, ")
                sb.AppendLine("  :p_created_at, :p_created_user_id, :p_created_pg_id, ")
                sb.AppendLine("  :p_updated_at, :p_updated_user_id, :p_updated_pg_id ")
                If (type = OrdersTable.Orders) Then
                    sb.AppendLine(",  :p_stra_order_qty, :p_stra_ship_qty, :p_stra_order_backlog ")
                End If
                sb.AppendLine(")")
#If False Then
                Dim sql As String =
                    $"INSERT INTO {GetTableName(type)} (" &
                    $" {GetIdName(type)}, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                    "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                    "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                    "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                    "  order_no, remarks, delivery_code, total_ship_qty, ship_date, " &
                    "  transport_method, ship_plan_date, customer_order_line_no, " &
                    "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                    "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                    "  reconcile_type, imp_run_id, status, active_flag, " &
                    "  created_at, created_user_id, created_pg_id, " &
                    "  updated_at, updated_user_id, updated_pg_id, " &
                    "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                    " ) VALUES (" &
                    "  :p_order_id, :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, " &
                    "  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, " &
                    "  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, " &
                    "  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, " &
                    "  :p_order_no, :p_remarks, :p_delivery_code, :p_total_ship_qty, :p_ship_date, " &
                    "  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
                    "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, " &
                    "  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag, " &
                    "  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, " &
                    "  :p_created_at, :p_created_user_id, :p_created_pg_id, " &
                    "  :p_updated_at, :p_updated_user_id, :p_updated_pg_id, " &
                    "  :p_stra_order_qty, :p_stra_ship_qty, :p_stra_order_backlog " &
                    ")"
                Using cmd As New OracleCommand(sql, conn)
#End If
                Using cmd As New OracleCommand(sb.Tostring(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    For Each r In records

                        ' 仮テーブルがエラーになるため
                        If (r.ProductCode Is Nothing) Then
                            r.ProductCode = 0
                        End If

                        cmd.Parameters.Clear()

                        ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                        cmd.Parameters.Add(":p_order_id", OracleDbType.Int64).Value = r.OrderId ' NUMBER(10,0)

                        cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
                        cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerCode, 25)
                        cmd.Parameters.Add(":p_billing_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.BillingTo, 25)
                        cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 40).Value = SafeVarchar(r.CustomerOrderNo, 40)
                        cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = SafeVarchar(r.DemandStatus, 1) ' 1桁記号想定
                        cmd.Parameters.Add(":p_ship_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.ShipTo, 25)

                        cmd.Parameters.Add(":p_order_date", OracleDbType.Date).Value = r.OrderDate
                        cmd.Parameters.Add(":p_due_date", OracleDbType.Date).Value = r.DueDate
                        cmd.Parameters.Add(":p_ship_scheduled_date", OracleDbType.Date).Value = r.ShipScheduledDate
                        cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.CustomerItemNo, 45)
                        cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ItemNo, 45)

                        cmd.Parameters.Add(":p_demand_qty", OracleDbType.Int64).Value = r.DemandQty ' NUMBER(10,0)
                        cmd.Parameters.Add(":p_demand_unit", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.DemandUnit, 4)
                        cmd.Parameters.Add(":p_currency_code", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.CurrencyCode, 3)
                        cmd.Parameters.Add(":p_ship_stock_location", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.ShipStockLocation, 25)
                        cmd.Parameters.Add(":p_company_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CompanyId, 25)

                        cmd.Parameters.Add(":p_product_code", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ProductCode, 45)
                        cmd.Parameters.Add(":p_billing_standard", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.BillingStandard, 3)
                        cmd.Parameters.Add(":p_ship_process_type", OracleDbType.Char, 1).Value = SafeVarchar(r.ShipProcessType, 1)
                        cmd.Parameters.Add(":p_delivery_instr_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.DeliveryInstrFlag)

                        cmd.Parameters.Add(":p_order_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderNo, 45)
                        cmd.Parameters.Add(":p_remarks", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Remarks, 45)
                        cmd.Parameters.Add(":p_delivery_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.DeliveryCode, 25)
                        cmd.Parameters.Add(":p_total_ship_qty", OracleDbType.Decimal).Value = r.TotalShipQty
                        cmd.Parameters.Add(":p_ship_date", OracleDbType.Date).Value = r.ShipDate

                        cmd.Parameters.Add(":p_transport_method", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.TransportMethod, 3)
                        cmd.Parameters.Add(":p_ship_plan_date", OracleDbType.Date).Value = r.ShipPlanDate
                        cmd.Parameters.Add(":p_customer_order_line_no", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CustomerOrderLineNo, 2)

                        cmd.Parameters.Add(":p_pre_daily_order_qty", OracleDbType.Decimal).Value = r.PreDailyOrderQty
                        cmd.Parameters.Add(":p_pre_daily_delivery_date", OracleDbType.Date).Value = r.PreDailyDeliveryDate
                        cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = r.ImpFileId ' NUMBER(10,0)

                        cmd.Parameters.Add(":p_order_type", OracleDbType.Int16).Value = r.OrderType       ' NUMBER(1,0)
                        cmd.Parameters.Add(":p_prorated_type", OracleDbType.Int16).Value = r.ProratedType    ' NUMBER(1,0)
                        cmd.Parameters.Add(":p_customer_info_type", OracleDbType.Varchar2, 50).Value = SafeVarchar(r.CustomerInfoType, 50)
                        cmd.Parameters.Add(":p_info_type", OracleDbType.Char, 1).Value = SafeVarchar(r.InfoType, 1)
                        cmd.Parameters.Add(":p_self_fcst_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.SelfFcstFlag)
                        cmd.Parameters.Add(":p_self_fcst_delete_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.SelfFcstDeleteFlag)

                        cmd.Parameters.Add(":p_reconcile_type", OracleDbType.Int16).Value = r.ReconcileType   ' NUMBER(1,0)
                        cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Long).Value = r.ImpRunId
                        cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.Status, 20)
                        cmd.Parameters.Add(":p_active_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.ActiveFlag)

                        ' 監査
                        cmd.Parameters.Add(":p_created_at", OracleDbType.Date).Value = r.CreatedAt
                        cmd.Parameters.Add(":p_created_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.CreatedUserId, 9)
                        cmd.Parameters.Add(":p_created_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.CreatedPgId, 150)

                        cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = r.UpdatedAt
                        cmd.Parameters.Add(":p_updated_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.UpdatedUserId, 9)
                        cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.UpdatedPgId, 150)
                        ' Pharse2
                        If (type = OrdersTable.Orders) Then
                            cmd.Parameters.Add(":p_stra_order_qty", OracleDbType.Decimal).Value = r.StraOrderQty
                            cmd.Parameters.Add(":p_stra_ship_qty", OracleDbType.Decimal).Value = r.StraShipQty
                            cmd.Parameters.Add(":p_stra_order_backlog", OracleDbType.Decimal).Value = r.StraOrderBacklog
                        End If

                        cmd.ExecuteNonQuery()
                    Next

                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function

        ''' <summary>
        ''' OrdersStage Table から 条件でレコード抽出
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="status"></param>
        ''' <param name="activeFlag"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function GetOrders(conn As OracleConnection, tran As OracleTransaction,
            Optional ByVal status As String = Nothing,
            Optional ByVal activeFlag As String = Nothing,
            Optional ByVal customerSettingId As Long? = Nothing,
            Optional ByVal orderId As Long? = Nothing,
            Optional ByVal additionalConditions As String = Nothing
        ) As DataTable
            Return GetOrders(conn, tran, OrdersTable.Orders, status, activeFlag, customerSettingId, orderId, additionalConditions)
        End Function

        ''' <summary>
        ''' OrdersStage Table から 条件でレコード抽出
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="status"></param>
        ''' <param name="activeFlag"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function GetOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
            Optional ByVal status As String = Nothing,
            Optional ByVal activeFlag As String = Nothing,
            Optional ByVal customerSettingId As Long? = Nothing,
            Optional ByVal orderId As Long? = Nothing,
            Optional ByVal additionalConditions As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            Dim prm As New List(Of OracleParameter)()

            sb.AppendLine("SELECT * ")
            sb.AppendLine($"FROM {GetTableName(type)} ")
            sb.AppendLine("WHERE 1=1 ")

            If customerSettingId IsNot Nothing Then
                sb.AppendLine("AND customer_setting_id = :p_csid ")
                prm.Add(New OracleParameter(":p_csid", OracleDbType.Int64) With {.Value = customerSettingId})
            End If

            If orderId IsNot Nothing Then
                sb.AppendLine($"AND {GetIdName(type)} = :p_orderid ")
                prm.Add(New OracleParameter(":p_orderid", OracleDbType.Int64) With {.Value = orderId})
            End If

            If status IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If Not String.IsNullOrEmpty(activeFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = :p_active ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = activeFlag})
            End If

            If Not String.IsNullOrEmpty(additionalConditions) Then
                sb.AppendLine(additionalConditions)
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

            'Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
            '    cmd.CommandText = "
            '            SELECT *
            '            FROM orders
            '            WHERE status = :p_status 
            '            AND active_flag    = :p_activeFlag 
            '            AND customer_Setting_Id    = :p_customerSettingId "
            '    cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(status, 20)
            '    cmd.Parameters.Add(":p_activeFlag", OracleDbType.Char).Value = activeFlag
            '    cmd.Parameters.Add(":p_customerSettingId", OracleDbType.Int64).Value = customerSettingId

            '    Using reader As OracleDataReader = cmd.ExecuteReader()
            '        dt.Load(reader)
            '    End Using
            'End Using
            'Return dt

        End Function

        '''' <summary>
        '''' customerSettingId と prodMgmtUserId で OrderStage レコードを削除する
        '''' R.sagisaka create
        '''' </summary>
        '''' <param name="customerSettingId"></param>
        '''' <param name="prodMgmtUserId"></param>
        'Public Sub Delete(customerSettingId As Integer, prodMgmtUserId As String)
        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            Delete(customerSettingId, prodMgmtUserId)
        '            tran.Commit()
        '        End Using
        '    End Using
        'End Sub

        ''' <summary>
        ''' customerSettingId と prodMgmtUserId で OrderStage レコードを削除する
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="prodMgmtUserId"></param>
        ''' <returns></returns>
        Public Function Delete(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Integer, prodMgmtUserId As String) As String

            Return Delete(conn, tran, OrdersTable.Orders, customerSettingId, prodMgmtUserId)

        End Function


        ''' <summary>
        ''' customerSettingId と prodMgmtUserId で OrderStage レコードを削除する
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="prodMgmtUserId"></param>
        ''' <returns></returns>
        Public Function Delete(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingId As Integer, prodMgmtUserId As String) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                sb.AppendLine("DELETE ")
                sb.AppendLine($"FROM {GetTableName(type)} ")
                sb.AppendLine("WHERE CustomerSettingId = :customerSettingId ")

                Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
                If Not isAdmin Then
                    sb.AppendLine("AND prod_mgmt_user_id = :p_user ")
                End If

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    Dim pId = cmd.Parameters.Add(":customerSettingId", OracleDbType.Varchar2).Value = prodMgmtUserId
                    cmd.BindByName = True
                    If Not isAdmin Then
                        cmd.Parameters.Add(":p_user", OracleDbType.Varchar2).Value = prodMgmtUserId
                    End If
                    conn.Open()
                    Dim cnt = cmd.ExecuteNonQuery()
                End Using

            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function

        ''' <summary>
        ''' 受注後 内示差分 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterOrderUnofficialNoticeDifference(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Integer) As DataTable

            Return AfterOrderUnofficialNoticeDifference(conn, tran, OrdersTable.Orders, customerSettingId)

        End Function

        ''' <summary>
        ''' 受注後 内示差分 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterOrderUnofficialNoticeDifference(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingId As Integer) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            ' 2026/3/26 版
#If True Then
            ' -- 受注差異取得（内示：order_type = 1）
            sb.AppendLine("WITH ranked_ts AS ( ")
            sb.AppendLine("    SELECT DISTINCT ")
            sb.AppendLine("        created_at, ")
            sb.AppendLine("        ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine("    FROM orders_history_view ")
            sb.AppendLine("    WHERE status = 'DUE_SET' ")
            sb.AppendLine("      AND order_type = 1 ")
            sb.AppendLine("      AND customer_setting_id = :customer_setting_id ")
            sb.AppendLine("), ")
            sb.AppendLine("ts AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("        MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("    FROM ranked_ts ")
            sb.AppendLine("), ")
            sb.AppendLine("prev AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty, ")
            sb.AppendLine("        COUNT(*) AS previous_data ")
            sb.AppendLine("    FROM orders_history_view ohv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND ohv.created_at = ts.previous_ts ")
            sb.AppendLine("      AND ohv.status = 'DUE_SET' ")
            sb.AppendLine("      AND ohv.order_type = 1 ")
            sb.AppendLine("      AND ohv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty ")
            sb.AppendLine("), ")
            sb.AppendLine("cur AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty, ")
            sb.AppendLine("        COUNT(*) AS current_data ")
            sb.AppendLine("    FROM orders_history_view ohv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ohv.created_at = ts.current_ts ")
            sb.AppendLine("      AND ohv.status = 'DUE_SET' ")
            sb.AppendLine("      AND ohv.order_type = 1 ")
            sb.AppendLine("      AND ohv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty ")
            sb.AppendLine(") ")
            sb.AppendLine("SELECT ")
            sb.AppendLine("    NVL(p.customer_code,        c.customer_code)           AS ""取引先コード"", ")
            sb.AppendLine("    NVL(p.profit_center,        c.profit_center)           AS ""PC"", ")
            sb.AppendLine("    NVL(p.customer_unit_name,   c.customer_unit_name)      AS ""注文工場／担当者名"", ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no)                 AS ""品目No"", ")
            sb.AppendLine("    NVL(p.pre_daily_delivery_date, ")
            sb.AppendLine("        c.pre_daily_delivery_date)                          AS ""希望納期"", ")
            sb.AppendLine("    NVL(p.pre_daily_order_qty, ")
            sb.AppendLine("        c.pre_daily_order_qty)                              AS ""数量"", ")
            sb.AppendLine("    NVL(p.previous_data, 0)                                 AS ""前回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0)                                 AS ""今回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0) - NVL(p.previous_data, 0)       AS ""前回比"" ")
            sb.AppendLine("FROM prev p ")
            sb.AppendLine("FULL OUTER JOIN cur c ")
            sb.AppendLine("  ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("  AND p.item_no                  = c.item_no ")
            sb.AppendLine("  AND p.pre_daily_delivery_date  = c.pre_daily_delivery_date ")
            sb.AppendLine("  AND p.pre_daily_order_qty      = c.pre_daily_order_qty ")
            sb.AppendLine("ORDER BY ")
            sb.AppendLine("    NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("    NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date), ")
            sb.AppendLine("    NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty) ")
#End If

            ' 2026/3/16 版
#If False Then
            sb.AppendLine("WITH ranked_ts AS (  ")
            sb.AppendLine("   SELECT DISTINCT  ")
            sb.AppendLine("          created_at,  ")
            sb.AppendLine("          ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn  ")
            sb.AppendLine($"     FROM {GetTableName(type)}  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" ts AS (  ")
            sb.AppendLine("   SELECT MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts,  ")
            sb.AppendLine("          MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts  ")
            sb.AppendLine("     FROM ranked_ts  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" prev AS (  ")
            sb.AppendLine("   SELECT  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.pre_daily_delivery_date,  ")
            sb.AppendLine("       oh.pre_daily_order_qty,  ")
            sb.AppendLine("       COUNT(*) AS previous_data  ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh  ")
            sb.AppendLine("     CROSS JOIN ts  ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL  ")
            sb.AppendLine("      AND oh.created_at = ts.previous_ts  ")
            sb.AppendLine("      AND oh.status = 'DUE_SET'  ")
            sb.AppendLine("      AND oh.order_type = 1   ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id  ")
            sb.AppendLine("    GROUP BY  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.pre_daily_delivery_date,  ")
            sb.AppendLine("       oh.pre_daily_order_qty  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" cur AS (  ")
            sb.AppendLine("   SELECT  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.pre_daily_delivery_date,  ")
            sb.AppendLine("       oh.pre_daily_order_qty,  ")
            sb.AppendLine("       COUNT(*) AS current_data  ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh  ")
            sb.AppendLine("     CROSS JOIN ts  ")
            sb.AppendLine("    WHERE oh.created_at = ts.current_ts  ")
            sb.AppendLine("      AND oh.status = 'DUE_SET'  ")
            sb.AppendLine("      AND oh.order_type = 1  ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id  ")
            sb.AppendLine("    GROUP BY  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.pre_daily_delivery_date,  ")
            sb.AppendLine("       oh.pre_daily_order_qty  ")
            sb.AppendLine(" )  ")
            sb.AppendLine(" SELECT  ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id)   AS customer_setting_id,  ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no) AS item_no,  ")
            sb.AppendLine("   NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date) AS pre_daily_delivery_date,  ")
            sb.AppendLine("   NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty)   AS pre_daily_order_qty,  ")
            sb.AppendLine("   NVL(p.previous_data, 0)  AS previous_data,  ")
            sb.AppendLine("   NVL(c.current_data,  0) AS current_data,  ")
            sb.AppendLine("   NVL(c.current_data,  0) - NVL(p.previous_data, 0)    AS diff_from_prev  ")
            sb.AppendLine(" FROM prev p  ")
            sb.AppendLine(" FULL OUTER JOIN cur c  ")
            sb.AppendLine("   ON  p.customer_setting_id      = c.customer_setting_id  ")
            sb.AppendLine("   And p.item_no    = c.item_no  ")
            sb.AppendLine("   AND p.pre_daily_delivery_date  = c.pre_daily_delivery_date  ")
            sb.AppendLine("   AND p.pre_daily_order_qty      = c.pre_daily_order_qty  ")
            sb.AppendLine(" ORDER BY  ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id),  ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no),  ")
            sb.AppendLine("   NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date),  ")
            sb.AppendLine("   NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty) ")
#End If
            Try
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId ' NUMBER(10,0))
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try
            Return dt
        End Function

        ''' <summary>
        ''' 受注後 受注差異取得 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterOrderDeliveryInstructionDifference(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Integer) As DataTable

            Return AfterOrderDeliveryInstructionDifference(conn, tran, OrdersTable.Orders, customerSettingId)

        End Function

        ''' <summary>
        ''' 受注後 受注差異取得 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterOrderDeliveryInstructionDifference(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingId As Integer) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            ' 2026/3/26 版
#If True Then
            '-- 受注差異取得（確定/納入指示：order_type <> 1）
            sb.AppendLine("WITH ranked_ts AS ( ")
            sb.AppendLine("    SELECT DISTINCT ")
            sb.AppendLine("        created_at, ")
            sb.AppendLine("        ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine("    FROM orders_history_view ")
            sb.AppendLine("    WHERE status = 'DUE_SET' ")
            sb.AppendLine("      AND order_type <> 1 ")
            sb.AppendLine("      AND customer_setting_id = :customer_setting_id ")
            sb.AppendLine("), ")
            sb.AppendLine("ts AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("        MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("    FROM ranked_ts ")
            sb.AppendLine("), ")
            sb.AppendLine("prev AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.customer_order_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty, ")
            sb.AppendLine("        COUNT(*) AS previous_data ")
            sb.AppendLine("    FROM orders_history_view ohv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND ohv.created_at = ts.previous_ts ")
            sb.AppendLine("      AND ohv.status = 'DUE_SET' ")
            sb.AppendLine("      AND ohv.order_type <> 1 ")
            sb.AppendLine("      AND ohv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.customer_order_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty ")
            sb.AppendLine("), ")
            sb.AppendLine("cur AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.customer_order_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty, ")
            sb.AppendLine("        COUNT(*) AS current_data ")
            sb.AppendLine("    FROM orders_history_view ohv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ohv.created_at = ts.current_ts ")
            sb.AppendLine("      AND ohv.status = 'DUE_SET' ")
            sb.AppendLine("      AND ohv.order_type <> 1 ")
            sb.AppendLine("      AND ohv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        ohv.customer_setting_id, ")
            sb.AppendLine("        ohv.customer_code, ")
            sb.AppendLine("        ohv.profit_center, ")
            sb.AppendLine("        ohv.customer_unit_id, ")
            sb.AppendLine("        ohv.customer_unit_name, ")
            sb.AppendLine("        ohv.item_no, ")
            sb.AppendLine("        ohv.customer_order_no, ")
            sb.AppendLine("        ohv.pre_daily_delivery_date, ")
            sb.AppendLine("        ohv.pre_daily_order_qty ")
            sb.AppendLine(") ")
            sb.AppendLine("SELECT ")
            sb.AppendLine("    NVL(p.customer_code,        c.customer_code)           AS ""取引先コード"", ")
            sb.AppendLine("    NVL(p.profit_center,        c.profit_center)           AS ""PC"", ")
            sb.AppendLine("    NVL(p.customer_unit_name,   c.customer_unit_name)      AS ""注文工場／担当者名"", ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no)                 AS ""品目No"", ")
            sb.AppendLine("    NVL(p.customer_order_no, c.customer_order_no)          AS ""客先発注No"", ")
            sb.AppendLine("    NVL(p.pre_daily_delivery_date, ")
            sb.AppendLine("        c.pre_daily_delivery_date)                          AS ""希望納期"", ")
            sb.AppendLine("    NVL(p.pre_daily_order_qty, ")
            sb.AppendLine("        c.pre_daily_order_qty)                              AS ""数量"", ")
            sb.AppendLine("    NVL(p.previous_data, 0)                                 AS ""前回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0)                                 AS ""今回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0) - NVL(p.previous_data, 0)       AS ""前回比"" ")
            sb.AppendLine("FROM prev p ")
            sb.AppendLine("FULL OUTER JOIN cur c ")
            sb.AppendLine("  ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("  AND p.item_no                  = c.item_no ")
            sb.AppendLine("  AND p.customer_order_no  = c.customer_order_no ")
            sb.AppendLine("  AND p.pre_daily_delivery_date  = c.pre_daily_delivery_date ")
            sb.AppendLine("  AND p.pre_daily_order_qty      = c.pre_daily_order_qty ")
            sb.AppendLine("ORDER BY ")
            sb.AppendLine("    NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("    NVL(p.customer_order_no,  c.customer_order_no), ")
            sb.AppendLine("    NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date), ")
            sb.AppendLine("    NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty); ")
#End If
            ' 2026/3/16 版
#If False Then
            sb.AppendLine(" WITH ranked_ts AS ( ")
            sb.AppendLine("   SELECT DISTINCT ")
            sb.AppendLine("          created_at, ")
            sb.AppendLine("          ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine($"     FROM {GetTableName(type)} ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" ts AS ( ")
            sb.AppendLine("   SELECT MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("          MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("     FROM ranked_ts ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" prev AS ( ")
            sb.AppendLine("   SELECT ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.pre_daily_delivery_date, ")
            sb.AppendLine("       oh.pre_daily_order_qty, ")
            sb.AppendLine("       COUNT(*) AS previous_data ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh ")
            sb.AppendLine("     CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND oh.created_at = ts.previous_ts ")
            sb.AppendLine("      AND oh.status = 'DUE_SET' ")
            sb.AppendLine("      AND oh.order_type <> 1 ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.pre_daily_delivery_date, ")
            sb.AppendLine("       oh.pre_daily_order_qty ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" cur AS ( ")
            sb.AppendLine("   SELECT ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.pre_daily_delivery_date, ")
            sb.AppendLine("       oh.pre_daily_order_qty, ")
            sb.AppendLine("       COUNT(*) AS current_data ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh ")
            sb.AppendLine("     CROSS JOIN ts ")
            sb.AppendLine("    WHERE oh.created_at = ts.current_ts ")
            sb.AppendLine("      AND oh.status = 'DUE_SET' ")
            sb.AppendLine("      AND oh.order_type <> 1 ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.pre_daily_delivery_date, ")
            sb.AppendLine("       oh.pre_daily_order_qty ")
            sb.AppendLine(" ) ")
            sb.AppendLine(" SELECT ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id)        AS customer_setting_id, ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no)                    AS item_no, ")
            sb.AppendLine("   NVL(p.customer_order_no,    c.customer_order_no)          AS customer_order_no, ")
            sb.AppendLine("   NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date) AS pre_daily_delivery_date, ")
            sb.AppendLine("   NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty)        AS pre_daily_order_qty, ")
            sb.AppendLine("   NVL(p.previous_data, 0)                                   AS previous_data, ")
            sb.AppendLine("   NVL(c.current_data,  0)                                   AS current_data, ")
            sb.AppendLine("   NVL(c.current_data,  0) - NVL(p.previous_data, 0)         AS diff_from_prev ")
            sb.AppendLine(" FROM prev p ")
            sb.AppendLine(" FULL OUTER JOIN cur c ")
            sb.AppendLine("   ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("   AND p.item_no                  = c.item_no ")
            sb.AppendLine("   AND p.customer_order_no        = c.customer_order_no ")
            sb.AppendLine("   AND p.pre_daily_delivery_date  = c.pre_daily_delivery_date ")
            sb.AppendLine("   AND p.pre_daily_order_qty      = c.pre_daily_order_qty ")
            sb.AppendLine(" ORDER BY ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("   NVL(p.customer_order_no,    c.customer_order_no), ")
            sb.AppendLine("   NVL(p.pre_daily_delivery_date, c.pre_daily_delivery_date), ")
            sb.AppendLine("   NVL(p.pre_daily_order_qty,  c.pre_daily_order_qty) ")
#End If

            Try
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId ' NUMBER(10,0))
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try

            Return dt
        End Function

        ''' <summary>
        ''' 生産計画後 内示差分 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterProductionPlanningUnofficialNoticeDifference(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Integer) As DataTable

            Return AfterProductionPlanningUnofficialNoticeDifference(conn, tran, OrdersTable.Orders, customerSettingId)

        End Function

        ''' <summary>
        ''' 生産計画後 内示差分 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterProductionPlanningUnofficialNoticeDifference(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingId As Integer) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()

            ' 2026/3/26 版
#If True Then
            '-- 生産計画差異取得（内示：order_type = 1） ")
            sb.AppendLine("WITH ranked_ts AS ( ")
            sb.AppendLine("    SELECT DISTINCT ")
            sb.AppendLine("        created_at, ")
            sb.AppendLine("        ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine("    FROM orders_history_view ")
            sb.AppendLine("    WHERE status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND order_type = 1 ")
            sb.AppendLine("      AND customer_setting_id = :customer_setting_id ")
            sb.AppendLine("), ")
            sb.AppendLine("ts AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("        MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("    FROM ranked_ts ")
            sb.AppendLine("), ")
            sb.AppendLine("prev AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty, ")
            sb.AppendLine("        COUNT(*) AS previous_data ")
            sb.AppendLine("    FROM prod_plan_history_view phv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND phv.created_at = ts.previous_ts ")
            sb.AppendLine("      AND phv.status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND phv.order_type = 1 ")
            sb.AppendLine("      AND phv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("   GROUP BY ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty ")
            sb.AppendLine("), ")
            sb.AppendLine("cur AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty, ")
            sb.AppendLine("        COUNT(*) AS current_data ")
            sb.AppendLine("    FROM prod_plan_history_view phv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE phv.created_at = ts.current_ts ")
            sb.AppendLine("      AND phv.status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND phv.order_type = 1 ")
            sb.AppendLine("      AND phv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty ")
            sb.AppendLine(") ")
            sb.AppendLine("SELECT ")
            sb.AppendLine("    NVL(p.customer_code,        c.customer_code)        AS ""取引先コード"", ")
            sb.AppendLine("    NVL(p.profit_center,        c.profit_center)        AS ""PC"", ")
            sb.AppendLine("    NVL(p.customer_unit_name,   c.customer_unit_name)   AS ""注文工場／担当者名"", ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no)              AS ""品目No"", ")
            sb.AppendLine("    NVL(p.due_date,             c.due_date)             AS ""希望納期"", ")
            sb.AppendLine("    NVL(p.demand_qty,           c.demand_qty)           AS ""数量"", ")
            sb.AppendLine("    NVL(p.previous_data, 0)                             AS ""前回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0)                             AS ""今回件数"", ")
            sb.AppendLine("    NVL(c.current_data,  0) - NVL(p.previous_data, 0)   AS ""前回比"" ")
            sb.AppendLine("FROM prev p ")
            sb.AppendLine("FULL OUTER JOIN cur c ")
            sb.AppendLine("    ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("    AND p.item_no                  = c.item_no ")
            sb.AppendLine("    AND p.due_date  = c.due_date ")
            sb.AppendLine("    AND p.demand_qty      = c.demand_qty ")
            sb.AppendLine("ORDER BY ")
            sb.AppendLine("    NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("    NVL(p.due_date,             c.due_date), ")
            sb.AppendLine("    NVL(p.demand_qty,           c.demand_qty) ")
#End If
            ' 2026/3/16 版
#If False Then
            sb.AppendLine("WITH ranked_ts AS (  ")
            sb.AppendLine("   SELECT DISTINCT  ")
            sb.AppendLine("          created_at,  ")
            sb.AppendLine("          ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn  ")
            sb.AppendLine($"     FROM {GetTableName(type)}  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" ts AS (  ")
            sb.AppendLine("   SELECT MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts,  ")
            sb.AppendLine("          MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts  ")
            sb.AppendLine("     FROM ranked_ts  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" prev AS (  ")
            sb.AppendLine("   SELECT  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.due_date,  ")
            sb.AppendLine("       oh.demand_qty,  ")
            sb.AppendLine("       COUNT(*) AS previous_data  ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh  ")
            sb.AppendLine("     CROSS JOIN ts  ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL  ")
            sb.AppendLine("      AND oh.created_at = ts.previous_ts  ")
            sb.AppendLine("      AND oh.status = 'DUE_SET'  ")
            sb.AppendLine("      AND oh.order_type = 1   ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id  ")
            sb.AppendLine("    GROUP BY  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.due_date,  ")
            sb.AppendLine("       oh.demand_qty  ")
            sb.AppendLine(" ),  ")
            sb.AppendLine(" cur AS (  ")
            sb.AppendLine("   SELECT  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.due_date,  ")
            sb.AppendLine("       oh.demand_qty,  ")
            sb.AppendLine("       COUNT(*) AS current_data  ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh  ")
            sb.AppendLine("     CROSS JOIN ts  ")
            sb.AppendLine("    WHERE oh.created_at = ts.current_ts  ")
            sb.AppendLine("      AND oh.status = 'DUE_SET'  ")
            sb.AppendLine("      AND oh.order_type = 1  ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id  ")
            sb.AppendLine("    GROUP BY  ")
            sb.AppendLine("       oh.customer_setting_id,  ")
            sb.AppendLine("       oh.item_no,  ")
            sb.AppendLine("       oh.due_date,  ")
            sb.AppendLine("       oh.demand_qty  ")
            sb.AppendLine(" )  ")
            sb.AppendLine(" SELECT  ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id)   AS customer_setting_id,  ")  ' 取引先設定ID
            sb.AppendLine("   NVL(p.item_no,              c.item_no) AS item_no,  ")                            ' 品目No
            sb.AppendLine("   NVL(p.due_date,             c.due_date) AS due_date,  ")                          ' 希望納期
            sb.AppendLine("   NVL(p.demand_qty,  c.demand_qty)   AS demand_qty,  ")                             ' 数量
            sb.AppendLine("   NVL(p.previous_data, 0)  AS previous_data,  ")                                    ' 前回_件数
            sb.AppendLine("   NVL(c.current_data,  0) AS current_data,  ")                                      ' 今回_件数
            sb.AppendLine("   NVL(c.current_data,  0) - NVL(p.previous_data, 0)    AS diff_from_prev  ")        ' 前回比
            sb.AppendLine(" FROM prev p  ")
            sb.AppendLine(" FULL OUTER JOIN cur c  ")
            sb.AppendLine("   ON  p.customer_setting_id      = c.customer_setting_id  ")
            sb.AppendLine("   And p.item_no    = c.item_no  ")
            sb.AppendLine("   AND p.due_date  = c.due_date  ")
            sb.AppendLine("   AND p.demand_qty      = c.demand_qty  ")
            sb.AppendLine(" ORDER BY  ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id),  ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no),  ")
            sb.AppendLine("   NVL(p.due_date, c.due_date),  ")
            sb.AppendLine("   NVL(p.demand_qty,  c.demand_qty) ")
#End If
            Try
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId ' NUMBER(10,0))
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try
            Return dt
        End Function
        ''' <summary>
        ''' 生産計画後 受注差異取得 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterProductionPlanningDeliveryInstructionDifference(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Integer) As DataTable

            Return AfterProductionPlanningDeliveryInstructionDifference(conn, tran, OrdersTable.Orders, customerSettingId)

        End Function

        ''' <summary>
        ''' 生産計画後 受注差異取得 取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function AfterProductionPlanningDeliveryInstructionDifference(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingId As Integer) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            ' 2026/3/26 版
            '-- 生産計画差異取得（確定・納入指示：order_type <> 1） ")
            sb.AppendLine("WITH ranked_ts AS ( ")
            sb.AppendLine("    SELECT DISTINCT ")
            sb.AppendLine("        created_at, ")
            sb.AppendLine("        ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine("    FROM prod_plan_history_view ")
            sb.AppendLine("    WHERE status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND order_type <> 1 ")
            sb.AppendLine("      AND customer_setting_id = :customer_setting_id ")
            sb.AppendLine("), ")
            sb.AppendLine("ts AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("        MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("    FROM ranked_ts ")
            sb.AppendLine("), ")
            sb.AppendLine("prev AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.customer_order_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty, ")
            sb.AppendLine("        COUNT(*) AS previous_data ")
            sb.AppendLine("    FROM prod_plan_history_view phv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND phv.created_at = ts.previous_ts ")
            sb.AppendLine("      AND phv.status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND phv.order_type <> 1 ")
            sb.AppendLine("      AND phv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.customer_order_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty ")
            sb.AppendLine("), ")
            sb.AppendLine("cur AS ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.customer_order_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty, ")
            sb.AppendLine("        COUNT(*) AS current_data ")
            sb.AppendLine("    FROM prod_plan_history_view phv ")
            sb.AppendLine("    CROSS JOIN ts ")
            sb.AppendLine("    WHERE phv.created_at = ts.current_ts ")
            sb.AppendLine("      AND phv.status = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("      AND phv.order_type <> 1 ")
            sb.AppendLine("      AND phv.customer_setting_id = :customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("        phv.customer_setting_id, ")
            sb.AppendLine("        phv.customer_code, ")
            sb.AppendLine("        phv.profit_center, ")
            sb.AppendLine("        phv.customer_unit_id, ")
            sb.AppendLine("        phv.customer_unit_name, ")
            sb.AppendLine("        phv.item_no, ")
            sb.AppendLine("        phv.customer_order_no, ")
            sb.AppendLine("        phv.due_date, ")
            sb.AppendLine("        phv.demand_qty ")
            sb.AppendLine(") ")
            sb.AppendLine("SELECT ")
            sb.AppendLine("    NVL(p.customer_code,        c.customer_code)           AS 取引先コード, ")
            sb.AppendLine("    NVL(p.profit_center,        c.profit_center)           AS PC, ")
            sb.AppendLine("    NVL(p.customer_unit_name,   c.customer_unit_name)      AS 注文工場／担当者名, ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no)                 AS 品目No, ")
            sb.AppendLine("    NVL(p.customer_order_no, c.customer_order_no)          AS 客先発注No, ")
            sb.AppendLine("    NVL(p.due_date,             c.due_date)                AS 希望納期, ")
            sb.AppendLine("    NVL(p.demand_qty,           c.demand_qty)              AS 数量, ")
            sb.AppendLine("    NVL(p.previous_data, 0)                                AS 前回件数, ")
            sb.AppendLine("    NVL(c.current_data,  0)                                AS 今回件数, ")
            sb.AppendLine("    NVL(c.current_data,  0) - NVL(p.previous_data, 0)      AS 前回比 ")
            sb.AppendLine("FROM prev p ")
            sb.AppendLine("FULL OUTER JOIN cur c ")
            sb.AppendLine("    ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("    AND p.item_no                  = c.item_no ")
            sb.AppendLine("    AND p.customer_order_no  = c.customer_order_no ")
            sb.AppendLine("    AND p.due_date            = c.due_date ")
            sb.AppendLine("    AND p.demand_qty          = c.demand_qty ")
            sb.AppendLine("ORDER BY ")
            sb.AppendLine("    NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("    NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("    NVL(p.customer_order_no,  c.customer_order_no), ")
            sb.AppendLine("    NVL(p.due_date,             c.due_date), ")
            sb.AppendLine("    NVL(p.demand_qty,           c.demand_qty) ")

            ' 2026/3/16 版
#If False Then
            sb.AppendLine(" WITH ranked_ts AS ( ")
            sb.AppendLine("   SELECT DISTINCT ")
            sb.AppendLine("          created_at, ")
            sb.AppendLine("          ROW_NUMBER() OVER (ORDER BY created_at DESC) AS rn ")
            sb.AppendLine($"     FROM {GetTableName(type)} ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" ts AS ( ")
            sb.AppendLine("   SELECT MAX(CASE WHEN rn = 1 THEN created_at END) AS current_ts, ")
            sb.AppendLine("          MAX(CASE WHEN rn = 2 THEN created_at END) AS previous_ts ")
            sb.AppendLine("     FROM ranked_ts ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" prev AS ( ")
            sb.AppendLine("   SELECT ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.due_date, ")
            sb.AppendLine("       oh.demand_qty, ")
            sb.AppendLine("       COUNT(*) AS previous_data ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh ")
            sb.AppendLine("     CROSS JOIN ts ")
            sb.AppendLine("    WHERE ts.previous_ts IS NOT NULL ")
            sb.AppendLine("      AND oh.created_at = ts.previous_ts ")
            sb.AppendLine("      AND oh.status = 'DUE_SET' ")
            sb.AppendLine("      AND oh.order_type <> 1 ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.due_date, ")
            sb.AppendLine("       oh.demand_qty ")
            sb.AppendLine(" ), ")
            sb.AppendLine(" cur AS ( ")
            sb.AppendLine("   SELECT ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.due_date, ")
            sb.AppendLine("       oh.demand_qty, ")
            sb.AppendLine("       COUNT(*) AS current_data ")
            sb.AppendLine($"     FROM {GetTableName(type)} oh ")
            sb.AppendLine("     CROSS JOIN ts ")
            sb.AppendLine("    WHERE oh.created_at = ts.current_ts ")
            sb.AppendLine("      AND oh.status = 'DUE_SET' ")
            sb.AppendLine("      AND oh.order_type <> 1 ")
            sb.AppendLine("      AND oh.customer_setting_id = :p_customer_setting_id ")
            sb.AppendLine("    GROUP BY ")
            sb.AppendLine("       oh.customer_setting_id, ")
            sb.AppendLine("       oh.item_no, ")
            sb.AppendLine("       oh.customer_order_no, ")
            sb.AppendLine("       oh.due_date, ")
            sb.AppendLine("       oh.demand_qty ")
            sb.AppendLine(" ) ")
            sb.AppendLine(" SELECT ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id)        AS customer_setting_id, ")  '取引先設定ID
            sb.AppendLine("   NVL(p.item_no,              c.item_no)                    AS item_no, ")              '品目No
            sb.AppendLine("   NVL(p.customer_order_no,    c.customer_order_no)          AS customer_order_no, ")    '客先発注No
            sb.AppendLine("   NVL(p.due_date, c.due_date)                               AS due_date, ")             '希望納期
            sb.AppendLine("   NVL(p.demand_qty,  c.demand_qty)                          AS demand_qty, ")           '数量
            sb.AppendLine("   NVL(p.previous_data, 0)                                   AS previous_data, ")        '前回_件数
            sb.AppendLine("   NVL(c.current_data,  0)                                   AS current_data, ")         '今回_件数
            sb.AppendLine("   NVL(c.current_data,  0) - NVL(p.previous_data, 0)         AS diff_from_prev ")        '前回比
            sb.AppendLine(" FROM prev p ")
            sb.AppendLine(" FULL OUTER JOIN cur c ")
            sb.AppendLine("   ON  p.customer_setting_id      = c.customer_setting_id ")
            sb.AppendLine("   AND p.item_no                  = c.item_no ")
            sb.AppendLine("   AND p.customer_order_no        = c.customer_order_no ")
            sb.AppendLine("   AND p.due_date  = c.due_date ")
            sb.AppendLine("   AND p.demand_qty      = c.demand_qty ")
            sb.AppendLine(" ORDER BY ")
            sb.AppendLine("   NVL(p.customer_setting_id,  c.customer_setting_id), ")
            sb.AppendLine("   NVL(p.item_no,              c.item_no), ")
            sb.AppendLine("   NVL(p.customer_order_no,    c.customer_order_no), ")
            sb.AppendLine("   NVL(p.due_date, c.due_date), ")
            sb.AppendLine("   NVL(p.demand_qty,  c.demand_qty) ")
#End If
            Try
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId ' NUMBER(10,0))
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try

            Return dt
        End Function

        Public Function ToClass(dt As DataRow) As OrdersRowHistolyRow

            Dim osr = New OrdersRowHistolyRow
            If (dt.Table.Columns.Contains("order_id")) Then
                osr.OrderId = If(dt.Field(Of Long?)("order_id"), 0)
                'osr.OrderId = dt.Field(Of Long)("order_id")
            ElseIf (dt.Table.Columns.Contains("prod_plan_id")) Then
                osr.OrderId = If(dt.Field(Of Long?)("prod_plan_id"), 0)
                'osr.OrderId = dt.Field(Of Long)("prod_plan_id")
            End If
            'osr.OrderId = dt.Field(Of Long)("order_id")
            osr.CustomerSettingId = dt.Field(Of String)("customer_setting_id")
            osr.CustomerCode = dt.Field(Of String)("customer_code")
            osr.BillingTo = dt.Field(Of String)("billing_to")
            osr.CustomerOrderNo = dt.Field(Of String)("customer_order_no")
            osr.DemandStatus = dt.Field(Of String)("demand_status")
            osr.ShipTo = dt.Field(Of String)("ship_to")
            osr.CustomerItemNo = dt.Field(Of String)("customer_item_no")
            osr.ItemNo = dt.Field(Of String)("item_no")
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
            osr.TransportMethod = dt.Field(Of String)("transport_method")
            osr.CustomerOrderLineNo = dt.Field(Of String)("customer_order_line_no")
            osr.CustomerInfoType = dt.Field(Of String)("customer_info_type")
            osr.InfoType = dt.Field(Of String)("info_type")
            osr.SelfFcstFlag = dt.Field(Of String)("self_fcst_flag")
            osr.SelfFcstDeleteFlag = dt.Field(Of String)("self_fcst_delete_flag")
            osr.ImpRunId = dt.Field(Of Long)("imp_run_id")
            osr.Status = dt.Field(Of String)("status")
            osr.ActiveFlag = dt.Field(Of String)("active_flag")
            ' ====== 数値系 ======
            osr.DemandQty = dt.Field(Of Long?)("demand_qty")
            osr.TotalShipQty = If(dt.Field(Of Decimal?)("total_ship_qty"), 0)
            'osr.TotalShipQty = dt.Field(Of Decimal?)("total_ship_qty")
            osr.PreDailyOrderQty = If(dt.Field(Of Decimal?)("pre_daily_order_qty"), 0)
            'osr.PreDailyOrderQty = dt.Field(Of Decimal?)("pre_daily_order_qty")
            osr.ImpFileId = If(dt.Field(Of Long?)("imp_file_id"), 0)
            'osr.ImpFileId = dt.Field(Of Long?)("imp_file_id")
            osr.OrderType = dt.Field(Of Int16?)("order_type")
            osr.ProratedType = dt.Field(Of Int16?)("prorated_type")
            osr.ReconcileType = If(dt.Field(Of Int16?)("reconcile_type"), 0)
            'osr.ReconcileType = dt.Field(Of Int16?)("reconcile_type")
            ' Pharse2
            If (dt.Table.Columns.Contains("order_id")) Then
                osr.StraOrderQty = If(dt.Field(Of Decimal?)("stra_order_qty"), 0)
                'osr.StraOrderQty = dt.Field(Of Decimal?)("stra_order_qty")
                osr.StraShipQty = If(dt.Field(Of Decimal?)("stra_ship_qty"), 0)
                'osr.StraShipQty = dt.Field(Of Decimal?)("stra_ship_qty")
                osr.StraOrderBacklog = If(dt.Field(Of Decimal?)("stra_order_backlog"),0)
                'osr.StraOrderBacklog = dt.Field(Of Decimal?)("stra_order_backlog")
            End If
            ' ====== 日付系 ======
            osr.OrderDate = dt.Field(Of Date?)("order_date")
            osr.DueDate = dt.Field(Of Date?)("due_date")
            osr.ShipScheduledDate = dt.Field(Of Date?)("ship_scheduled_date")
            osr.ShipDate = dt.Field(Of Date?)("ship_date")
            osr.ShipPlanDate = dt.Field(Of Date?)("ship_plan_date")
            osr.PreDailyDeliveryDate = dt.Field(Of Date?)("pre_daily_delivery_date")
            ' 監査系
            osr.CreatedAt = dt.Field(Of DateTime?)("created_at")
            osr.CreatedUserId = dt.Field(Of String)("created_user_id")
            osr.CreatedPgId = dt.Field(Of String)("created_pg_id")
            osr.UpdatedAt = dt.Field(Of DateTime?)("updated_at")
            osr.UpdatedUserId = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPgId = dt.Field(Of String)("updated_pg_id")
            ' Stage 固有
            osr.HistoryId = If(dt.Field(Of Long?)("stage_id"), 0)
            'osr.HistoryId = dt.Field(Of Long?)("stage_id")

            Return osr

        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As List(Of OrdersRowHistolyRow)

            Dim osrs = New List(Of OrdersRowHistolyRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function


        '''' <summary>
        '''' 受注差異取得 差異リスト 内示 DataRow to class
        '''' </summary>
        '''' <param name="dt"></param>
        '''' <returns></returns>
        'Public Function AfterOrderUnNoticeDifferenceToClass(dt As DataTable) As List(Of AfterOrderUnNoticeDifferenceRow)

        '    Dim osrs = New List(Of AfterOrderUnNoticeDifferenceRow)()

        '    For Each dtRow In dt.Rows
        '        osrs.Add(AfterOrderUnNoticeDifferenceToClass(dtRow))
        '    Next

        '    Return osrs

        'End Function

        '''' <summary>
        '''' 受注差異取得 差異リスト  確定納入指示 DataRow to class
        '''' </summary>
        '''' <param name="dt"></param>
        '''' <returns></returns>
        'Public Function AfterOrderInstructionDifferenceToClass(dt As DataTable) As List(Of AfterOrderInstructionDifferenceRow)

        '    Dim osrs = New List(Of AfterOrderInstructionDifferenceRow)()

        '    For Each dtRow In dt.Rows
        '        osrs.Add(AfterOrderInstructionDifferenceToClass(dtRow))
        '    Next
        '    Return osrs

        'End Function


        '''' <summary>
        '''' 差異 内示 DataRow to class
        '''' </summary>
        '''' <param name="dt"></param>
        '''' <returns></returns>
        'Public Function AfterOrderUnNoticeDifferenceToClass(dt As DataRow) As AfterOrderUnNoticeDifferenceRow

        '    Dim osr = New AfterOrderUnNoticeDifferenceRow()
        '    osr.CustomerCode = dt.Field(Of String)("取引先コード")
        '    osr.ProfitCenter = dt.Field(Of String)("PC")
        '    osr.CustomerUnitName = dt.Field(Of String)("注文工場/ 担当者名")
        '    osr.PreDailyDeliveryDate = dt.Field(Of Date?)("希望納期")
        '    osr.PreDailyOrderQty = dt.Field(Of Long?)("数量")
        '    osr.PreviousData = dt.Field(Of Long?)("前回件数")
        '    osr.CurrentData = dt.Field(Of Long?)("今回件数")
        '    osr.DiffFromPrev = dt.Field(Of Long?)("前回比")
        '    Return osr
        'End Function

        '''' <summary>
        '''' 差異 内示 DataRow to class
        '''' </summary>
        '''' <param name="dt"></param>
        '''' <returns></returns>
        'Public Function AfterOrderInstructionDifferenceToClass(dt As DataRow) As AfterOrderInstructionDifferenceRow

        '    Dim osr = New AfterOrderInstructionDifferenceRow()
        '    osr.CustomerSettingId = dt.Field(Of String)("customer_setting_id")
        '    osr.ItemNo = dt.Field(Of String)("item_no")
        '    osr.CustomerOrderNo = dt.Field(Of String)("customer_order_no")
        '    osr.PreDailyDeliveryDate = dt.Field(Of Date?)("pre_daily_delivery_date")
        '    osr.PreDailyOrderQty = dt.Field(Of Long?)("pre_daily_order_qty")
        '    osr.PreviousData = dt.Field(Of Long?)("previous_data")
        '    osr.CurrentData = dt.Field(Of Long?)("current_data")
        '    osr.DiffFromPrev = dt.Field(Of Long?)("diff_from_prev")
        '    Return osr

        'End Function

        ''' <summary>
        ''' 差異リスト 内示 DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function UnNoticeDifferenceToClass(dt As DataTable) As List(Of UnNoticeDifferenceRow)

            Dim osrs = New List(Of UnNoticeDifferenceRow)()

            For Each dtRow In dt.Rows
                osrs.Add(UnNoticeDifferenceToClass(dtRow))
            Next
            'osrs.Add(New UnNoticeDifferenceRow())
            Return osrs

        End Function

        ''' <summary>
        ''' 差異リスト  確定納入指示 DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function InstructionDifferenceToClass(dt As DataTable) As List(Of InstructionDifferenceRow)

            Dim osrs = New List(Of InstructionDifferenceRow)()

            For Each dtRow In dt.Rows
                osrs.Add(InstructionDifferenceToClass(dtRow))
            Next
            Return osrs

        End Function

        ''' <summary>
        ''' 差異 内示 DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function UnNoticeDifferenceToClass(dt As DataRow) As UnNoticeDifferenceRow

            Dim osr = New UnNoticeDifferenceRow()
            'osr.CustomerSettingId = dt.Field(Of String)("取引先設定ID")
            osr.CustomerCode = dt.Field(Of String)("取引先コード")
            osr.ProfitCenter = dt.Field(Of String)("PC")
            osr.CustomerUnitName = dt.Field(Of String)("注文工場／担当者名")
            osr.ItemNo = dt.Field(Of String)("品目No")
            osr.DueDate = dt.Field(Of String)("希望納期")
            osr.DemandQty = If(dt.Field(Of Long?)("数量"), 0)
            osr.PreviousData = If(dt.Field(Of Long?)("前回_件数"), 0)
            osr.CurrentData = If(dt.Field(Of Long?)("今回_件数"), 0)
            osr.DiffFromPrev = If(dt.Field(Of Long?)("前回比"), 0)
            Return osr

        End Function
        ''' <summary>
        ''' 差異  確定納入指示 DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function InstructionDifferenceToClass(dt As DataRow) As InstructionDifferenceRow

            Dim osr = New InstructionDifferenceRow()
            'osr.CustomerSettingId = dt.Field(Of String)("取引先設定ID")
            osr.CustomerCode = dt.Field(Of String)("取引先コード")
            osr.ProfitCenter = dt.Field(Of String)("PC")
            osr.CustomerUnitName = dt.Field(Of String)("注文工場／担当者名")
            osr.ItemNo = dt.Field(Of String)("品目No")
            osr.CustomerOrderNo = dt.Field(Of String)("客先発注No")
            osr.DueDate = dt.Field(Of String)("希望納期")
            osr.DemandQty = If(dt.Field(Of Long?)("数量"), 0)
            osr.PreviousData = If(dt.Field(Of Long?)("前回_件数"), 0)
            osr.CurrentData = If(dt.Field(Of Long?)("今回_件数"), 0)
            osr.DiffFromPrev = If(dt.Field(Of Long?)("前回比"), 0)
            Return osr

        End Function
    End Class
#If False Then
    '受注差異取得-内示
    Public Class AfterOrderUnNoticeDifferenceRow
        Public Property CustomerCode As String          '取引先コード		
        Public Property ProfitCenter As String          'PC					
        Public Property CustomerUnitName As String      '注文工場／担当者名	
        Public Property ItemNo As String                '品目No				
        Public Property PreDailyDeliveryDate As Date?   '希望納期			
        Public Property PreDailyOrderQty As Long?       '数量				
        Public Property PreviousData As Long?           '前回件数
        Public Property CurrentData As Long?            '今回件数			
        Public Property DiffFromPrev As Long?           '前回比				

        Public Sub New()

        End Sub

        Public Sub New(src As AfterOrderUnNoticeDifferenceRow)
            CustomerCode = src.CustomerCode
            ProfitCenter = src.ProfitCenter
            CustomerUnitName = src.CustomerUnitName
            ItemNo = src.ItemNo
            PreDailyDeliveryDate = src.PreDailyDeliveryDate
            PreDailyOrderQty = src.PreDailyOrderQty
            PreviousData = src.PreviousData
            CurrentData = src.CurrentData
            DiffFromPrev = src.DiffFromPrev
        End Sub

    End Class
    '受注差異取得-確定・納入指示
    Public Class AfterOrderInstructionDifferenceRow
        Public Property CustomerSettingId As Long?      '取引先設定ID		
        Public Property ItemNo As String                '品目No
        Public Property CustomerOrderNo As String       '客先発注No
        Public Property PreDailyDeliveryDate As Date?   '希望納期			
        Public Property PreDailyOrderQty As Long?       '数量				
        Public Property PreviousData As Long?           '前回件数
        Public Property CurrentData As Long?            '今回件数			
        Public Property DiffFromPrev As Long?           '前回比				

        Public Sub New()

        End Sub

        Public Sub New(src As AfterOrderInstructionDifferenceRow)
            CustomerSettingId = src.CustomerSettingId
            ItemNo = src.ItemNo
            CustomerOrderNo = src.CustomerOrderNo
            PreDailyDeliveryDate = src.PreDailyDeliveryDate
            PreDailyOrderQty = src.PreDailyOrderQty
            PreviousData = src.PreviousData
            CurrentData = src.CurrentData
            DiffFromPrev = src.DiffFromPrev
        End Sub
    End Class
#End If

    '差異取得-内示
    Public Class UnNoticeDifferenceRow
        'Public Property CustomerSettingId As Long?      '取引先設定ID
        Public Property CustomerCode As String          '取引先コード
        Public Property ProfitCenter As String          'PC
        Public Property CustomerUnitName As String      '注文工場／担当者名
        Public Property ItemNo As String                '品目No
        Public Property DueDate As Date?                '希望納期			
        Public Property DemandQty As Long?              '数量				
        Public Property PreviousData As Long?           '前回件数
        Public Property CurrentData As Long?            '今回件数			
        Public Property DiffFromPrev As Long?           '前回比				
        Public Sub New()

        End Sub
        Public Sub New(src As UnNoticeDifferenceRow)
            'CustomerSettingId = src.CustomerSettingId
            CustomerCode = src.CustomerCode
            ProfitCenter = src.ProfitCenter
            CustomerUnitName = src.CustomerUnitName
            ItemNo = src.ItemNo
            DueDate = src.DueDate
            DemandQty = src.DemandQty
            PreviousData = src.PreviousData
            CurrentData = src.CurrentData
            DiffFromPrev = src.DiffFromPrev
        End Sub
    End Class
    ' 差異取得-確定・納入指示
    Public Class InstructionDifferenceRow
        'Public Property CustomerSettingId As Long?      '取引先設定ID                       
        Public Property CustomerCode As String          '取引先コード
        Public Property ProfitCenter As String          'PC
        Public Property CustomerUnitName As String      '注文工場／担当者名
        Public Property ItemNo As String                '品目No
        Public Property CustomerOrderNo As String       '客先発注No
        Public Property DueDate As Date?                '希望納期
        Public Property DemandQty As Long?              '数量
        Public Property PreviousData As Long?           '前回件数
        Public Property CurrentData As Long?            '今回件数
        Public Property DiffFromPrev As Long?           '前回比
        Public Sub New()

        End Sub
        Public Sub New(src As InstructionDifferenceRow)
            'CustomerSettingId = src.CustomerSettingId
            CustomerCode = src.CustomerCode
            ProfitCenter = src.ProfitCenter
            CustomerUnitName = src.CustomerUnitName
            ItemNo = src.ItemNo
            CustomerOrderNo = src.CustomerOrderNo
            DueDate = src.DueDate
            DemandQty = src.DemandQty
            PreviousData = src.PreviousData
            CurrentData = src.CurrentData
            DiffFromPrev = src.DiffFromPrev
        End Sub

    End Class


    ''' <summary>
    ''' ORDERS 受け渡し用の行DTO（Repository向け）
    ''' </summary>
    Public Class OrdersRowHistolyRow

        Public Property OrderId As Long?                        ' ORDER_ID NUMBER(10,0)
        'Public Property ProdPlanId As Long?                        ' ORDER_ID NUMBER(10,0)

        ' ====== 取引先・品目等 ======
        Public Property CustomerSettingId As Long               ' CUSTOMER_SETTING_ID NUMBER(10,0)
        'Public Property CustomerSettingId As String             ' CUSTOMER_SETTING_ID VARCHAR2(25)
        Public Property CustomerCode As String                  ' CUSTOMER_CODE VARCHAR2(25)
        Public Property BillingTo As String                     ' BILLING_TO VARCHAR2(25)
        Public Property CustomerOrderNo As String               ' CUSTOMER_ORDER_NO CHAR(40)
        Public Property DemandStatus As String                  ' DEMAND_STATUS CHAR(1)
        Public Property ShipTo As String                        ' SHIP_TO VARCHAR2(25)
        Public Property CustomerItemNo As String                ' CUSTOMER_ITEM_NO VARCHAR2(45)
        Public Property ItemNo As String                        ' ITEM_NO VARCHAR2(45)
        Public Property DemandUnit As String                    ' DEMAND_UNIT CHAR(4)
        Public Property CurrencyCode As String                  ' CURRENCY_CODE CHAR(3)
        Public Property ShipStockLocation As String             ' SHIP_STOCK_LOCATION VARCHAR2(25)
        Public Property CompanyId As String                     ' COMPANY_ID VARCHAR2(25)
        Public Property ProductCode As String                   ' PRODUCT_CODE VARCHAR2(45)
        Public Property BillingStandard As String               ' BILLING_STANDARD VARCHAR2(3)
        Public Property ShipProcessType As String               ' SHIP_PROCESS_TYPE CHAR(1)
        Public Property DeliveryInstrFlag As String             ' DELIVERY_INSTR_FLAG CHAR(1)
        Public Property OrderNo As String                       ' ORDER_NO VARCHAR2(45)
        Public Property Remarks As String                       ' REMARKS VARCHAR2(45)
        Public Property DeliveryCode As String                  ' DELIVERY_CODE VARCHAR2(25)

        Public Property TransportMethod As String               ' TRANSPORT_METHOD VARCHAR2(3)
        Public Property CustomerOrderLineNo As String           ' CUSTOMER_ORDER_LINE_NO VARCHAR2(2)
        Public Property CustomerInfoType As String              ' CUSTOMER_INFO_TYPE VARCHAR2(50)
        Public Property InfoType As String                      ' INFO_TYPE CHAR(1)
        Public Property SelfFcstFlag As String                  ' SELF_FCST_FLAG CHAR(1)
        Public Property SelfFcstDeleteFlag As String            ' SELF_FCST_DELETE_FLAG CHAR(1)
        Public Property ImpRunId As Long                      ' IMP_RUN_ID NUMBER(10,0)
        Public Property Status As String                        ' STATUS VARCHAR2(20)
        Public Property ActiveFlag As String                    ' ACTIVE_FLAG CHAR(1)

        ' ====== 数値系 ======
        Public Property DemandQty As Long?                      ' DEMAND_QTY NUMBER(10,0)
        Public Property TotalShipQty As Decimal?                ' TOTAL_SHIP_QTY NUMBER(18,6)
        Public Property PreDailyOrderQty As Decimal?            ' PRE_DAILY_ORDER_QTY NUMBER(18,6)
        Public Property ImpFileId As Long?                      ' IMP_FILE_ID NUMBER(10,0)
        Public Property OrderType As Int16?                     ' ORDER_TYPE NUMBER(1,0)
        Public Property ProratedType As Int16?                  ' PRORATED_TYPE NUMBER(1,0)
        Public Property ReconcileType As Int16?                 ' RECONCILE_TYPE NUMBER(1,0)
        ' Pharse2
        Public Property StraOrderQty As Decimal?                ' STRA_ORDER_QTY NUMBER(18, 6)
        Public Property StraShipQty As Decimal?                 ' STRA_SHIP_QTY NUMBER(18, 6)
        Public Property StraOrderBacklog As Decimal?            ' STRA_ORDER_BACKLOG NUMBER(18, 6)

        ' ====== 日付系 ======
        Public Property OrderDate As Date?                      ' ORDER_DATE DATE
        Public Property DueDate As Date?                        ' DUE_DATE DATE
        Public Property ShipScheduledDate As Date?              ' SHIP_SCHEDULED_DATE DATE
        Public Property ShipDate As Date?                       ' SHIP_DATE DATE
        Public Property ShipPlanDate As Date?                   ' SHIP_PLAN_DATE DATE
        Public Property PreDailyDeliveryDate As Date           ' PRE_DAILY_DELIVERY_DATE DATE

        ' 監査系
        Public Property CreatedAt As DateTime           ' CREATED_AT
        Public Property CreatedUserId As String         ' CREATED_USER_ID
        Public Property CreatedPgId As String           ' CREATED_PG_ID
        Public Property UpdatedAt As DateTime           ' UPDATED_AT
        Public Property UpdatedUserId As String         ' UPDATED_USER_ID
        Public Property UpdatedPgId As String           ' UPDATED_PG_ID

        ' History
        ' ====== 主キー／識別系 ======
        Public Property HistoryId As Long?              ' HISTORY_ID

    End Class

End Namespace
