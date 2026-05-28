
Imports System.Configuration
Imports System.Data
Imports System.Data.Odbc
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports System.Threading
Imports DocumentFormat.OpenXml.Drawing.Diagrams
Imports DocumentFormat.OpenXml.Spreadsheet
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data
    Public Class OrderRepository
        Public Enum OrdersTable
            Orders = 1
            ProductPlan = 2
        End Enum

#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

        ''' <summary>
        ''' type に 応じたTable 名称を返す
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Private Function GetTableName(type As OrdersTable)

            Dim tableNames As New List(Of (id As OrdersTable, name As String)) From {
                (OrdersTable.Orders, "orders"),
                (OrdersTable.ProductPlan, "prod_plan")
            }
            Dim name = tableNames.Find(Function(x) x.id = type).name
            Return name

        End Function

        Private Function GetStageTableName(type As OrdersTable)

            Dim tableNames As New List(Of (id As OrdersTable, name As String)) From {
                (OrdersTable.Orders, "orders_stage"),
                (OrdersTable.ProductPlan, "prod_plan_stage")
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

        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="row"></param>
        Public Sub Insert(row As OrdersRow)

            Insert(OrdersTable.Orders, row)

        End Sub
        ''' <summary>
        ''' Orders Table Type 指定
        ''' </summary>
        ''' <param name="type"></param>
        ''' <param name="row"></param>
        Public Sub Insert(type As OrdersTable, row As OrdersRow)

            Dim records = New List(Of OrdersRow)()
            records.Add(row)
            InsertRange(type, records)

        End Sub
        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As OrdersRow) As String

            Return Insert(conn, tran, OrdersTable.Orders, row)

        End Function
        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, row As OrdersRow) As String

            Dim records = New List(Of OrdersRow)()
            records.Add(row)
            Return InsertRange(conn, tran, type, records)

        End Function

        ''' <summary>
        ''' OrderRow class リストをDBに追加する (元のコード同等呼び出し
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of OrdersRow))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, OrdersTable.Orders, records)
                    tran.Commit()
                End Using
            End Using
        End Sub
        ''' <summary>
        ''' OrderRow class リストをDBに追加する (元のコード同等呼び出し
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(type As OrdersTable, records As IEnumerable(Of OrdersRow))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, type, records)
                    tran.Commit()
                End Using
            End Using
        End Sub
#If False Then
        ''' <summary>
        ''' OrderRow class リストをDBに追加する
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of OrdersRow))
            If records Is Nothing Then Return

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Const sql As String =
                        "INSERT INTO orders (" &
                        "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  order_no, remarks, delivery_code, order_time, sales_unit_price, delivery_time, " &
                        "  usage_location, total_ship_qty, production_category, char_2, container_no, " &
                        "  char_3, char_4, char_4_2, char_5, char_5_2, char_6, " &
                        "  order_reason, container_capacity, customer_lot_no, initial_flag, ship_date, " &
                        "  char_50, transport_method, ship_plan_date, customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                        "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                        "  reconcile_type, imp_run_id, status, active_flag, " &
                        "  created_at, created_user_id, created_pg_id, " &
                        "  updated_at, updated_user_id, updated_pg_id" &
                        ") VALUES (" &
                        "  :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, " &
                        "  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, " &
                        "  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, " &
                        "  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, " &
                        "  :p_order_no, :p_remarks, :p_delivery_code, :p_order_time, :p_sales_unit_price, :p_delivery_time, " &
                        "  :p_usage_location, :p_total_ship_qty, :p_production_category, :p_char_2, :p_container_no, " &
                        "  :p_char_3, :p_char_4, :p_char_4_2, :p_char_5, :p_char_5_2, :p_char_6, " &
                        "  :p_order_reason, :p_container_capacity, :p_customer_lot_no, :p_initial_flag, :p_ship_date, " &
                        "  :p_char_50, :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
                        "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, " &
                        "  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag, " &
                        "  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, " &
                        "  :p_created_at, :p_created_user_id, :p_created_pg_id, " &
                        "  :p_updated_at, :p_updated_user_id, :p_updated_pg_id" &
                        ")"

                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Transaction = tran
                        cmd.BindByName = True
                        cmd.CommandType = CommandType.Text

                        For Each r In records
                            cmd.Parameters.Clear()

                            ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                            cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
                            cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerCode, 25)
                            cmd.Parameters.Add(":p_billing_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.BillingTo, 25)
                            cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 40).Value = SafeVarchar(r.CustomerOrderNo, 40)
                            cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = NormalizeYN(r.DemandStatus) ' 1桁記号想定
                            cmd.Parameters.Add(":p_ship_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.ShipTo, 25)

                            cmd.Parameters.Add(":p_order_date", OracleDbType.Date).Value = r.OrderDate
                            cmd.Parameters.Add(":p_due_date", OracleDbType.Date).Value = r.DueDate
                            cmd.Parameters.Add(":p_ship_scheduled_date", OracleDbType.Date).Value = r.ShipScheduledDate

                            cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.CustomerItemNo, 20)
                            cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.ItemNo, 20)

                            cmd.Parameters.Add(":p_demand_qty", OracleDbType.Int64).Value = r.DemandQty ' NUMBER(10,0)
                            cmd.Parameters.Add(":p_demand_unit", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.DemandUnit, 4)
                            cmd.Parameters.Add(":p_currency_code", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.CurrencyCode, 3)
                            cmd.Parameters.Add(":p_ship_stock_location", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.ShipStockLocation, 25)
                            cmd.Parameters.Add(":p_company_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CompanyId, 25)

                            cmd.Parameters.Add(":p_product_code", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.ProductCode, 20)
                            cmd.Parameters.Add(":p_billing_standard", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.BillingStandard, 3)
                            cmd.Parameters.Add(":p_ship_process_type", OracleDbType.Char, 1).Value = NormalizeYN(r.ShipProcessType)
                            cmd.Parameters.Add(":p_delivery_instr_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.DeliveryInstrFlag)

                            cmd.Parameters.Add(":p_order_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderNo, 45)
                            cmd.Parameters.Add(":p_remarks", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Remarks, 45)
                            cmd.Parameters.Add(":p_delivery_code", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.DeliveryCode, 20)

                            cmd.Parameters.Add(":p_order_time", OracleDbType.Decimal).Value = r.OrderTime       ' NUMBER(18,6)
                            cmd.Parameters.Add(":p_sales_unit_price", OracleDbType.Decimal).Value = r.SalesUnitPrice  ' NUMBER(18,6)
                            cmd.Parameters.Add(":p_delivery_time", OracleDbType.Decimal).Value = r.DeliveryTime    ' NUMBER(18,6)

                            cmd.Parameters.Add(":p_usage_location", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.UsageLocation, 45)
                            cmd.Parameters.Add(":p_total_ship_qty", OracleDbType.Decimal).Value = r.TotalShipQty
                            cmd.Parameters.Add(":p_production_category", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ProductionCategory, 45)
                            cmd.Parameters.Add(":p_char_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char2, 45)
                            cmd.Parameters.Add(":p_container_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ContainerNo, 45)

                            cmd.Parameters.Add(":p_char_3", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char3, 45)
                            cmd.Parameters.Add(":p_char_4", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4, 45)
                            cmd.Parameters.Add(":p_char_4_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4_2, 45)
                            cmd.Parameters.Add(":p_char_5", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5, 45)
                            cmd.Parameters.Add(":p_char_5_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5_2, 45)
                            cmd.Parameters.Add(":p_char_6", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char6, 45)

                            cmd.Parameters.Add(":p_order_reason", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderReason, 45)
                            cmd.Parameters.Add(":p_container_capacity", OracleDbType.Decimal).Value = r.ContainerCapacity
                            cmd.Parameters.Add(":p_customer_lot_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.CustomerLotNo, 45)
                            cmd.Parameters.Add(":p_initial_flag", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.InitialFlag, 45)

                            cmd.Parameters.Add(":p_ship_date", OracleDbType.Date).Value = r.ShipDate
                            cmd.Parameters.Add(":p_char_50", OracleDbType.Varchar2, 60).Value = SafeVarchar(r.Char50, 60)
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
                            cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Varchar2, 36).Value = SafeVarchar(r.ImpRunId, 36)
                            cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.Status, 20)
                            cmd.Parameters.Add(":p_active_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.ActiveFlag)

                            ' 監査
                            cmd.Parameters.Add(":p_created_at", OracleDbType.Date).Value = r.CreatedAt
                            cmd.Parameters.Add(":p_created_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.CreatedUserId, 9)
                            cmd.Parameters.Add(":p_created_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.CreatedPgId, 150)
                            cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = r.UpdatedAt
                            cmd.Parameters.Add(":p_updated_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.UpdatedUserId, 9)
                            cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.UpdatedPgId, 150)

                            cmd.ExecuteNonQuery()
                        Next

                        tran.Commit()
                    End Using
                End Using
            End Using
        End Sub
#Else
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
            If records Is Nothing Then Return "OrdersRow InsertRange no record data."
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO {GetTableName(type)} (")
                sb.AppendLine("  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, ")
                sb.AppendLine("  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, ")
                sb.AppendLine("  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, ")
                sb.AppendLine("  product_code, billing_standard, ship_process_type, delivery_instr_flag, ")
                sb.AppendLine("  order_no, remarks, delivery_code, order_time, sales_unit_price, delivery_time, ")
                sb.AppendLine("  usage_location, total_ship_qty, production_category, char_2, container_no, ")
                sb.AppendLine("  char_3, char_4, char_4_2, char_5, char_5_2, char_6, ")
                sb.AppendLine("  order_reason, container_capacity, customer_lot_no, initial_flag, ship_date, ")
                sb.AppendLine("  char_50, transport_method, ship_plan_date, customer_order_line_no, ")
                sb.AppendLine("  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, ")
                sb.AppendLine("  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, ")
                sb.AppendLine("  reconcile_type, imp_run_id, status, active_flag, ")
                sb.AppendLine("  created_at, created_user_id, created_pg_id, ")
                sb.AppendLine("  updated_at, updated_user_id, updated_pg_id ")
                If (type = OrdersTable.Orders) Then
                    sb.AppendLine(",  stra_order_qty, stra_ship_qty, stra_order_backlog ")
                End If
                sb.AppendLine(") VALUES (")
                sb.AppendLine("  :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, ")
                sb.AppendLine("  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, ")
                sb.AppendLine("  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, ")
                sb.AppendLine("  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, ")
                sb.AppendLine("  :p_order_no, :p_remarks, :p_delivery_code, :p_order_time, :p_sales_unit_price, :p_delivery_time, ")
                sb.AppendLine("  :p_usage_location, :p_total_ship_qty, :p_production_category, :p_char_2, :p_container_no, ")
                sb.AppendLine("  :p_char_3, :p_char_4, :p_char_4_2, :p_char_5, :p_char_5_2, :p_char_6, ")
                sb.AppendLine("  :p_order_reason, :p_container_capacity, :p_customer_lot_no, :p_initial_flag, :p_ship_date, ")
                sb.AppendLine("  :p_char_50, :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, ")
                sb.AppendLine("  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, ")
                sb.AppendLine("  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag,")
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
                "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                "  order_no, remarks, delivery_code, order_time, sales_unit_price, delivery_time, " &
                "  usage_location, total_ship_qty, production_category, char_2, container_no, " &
                "  char_3, char_4, char_4_2, char_5, char_5_2, char_6, " &
                "  order_reason, container_capacity, customer_lot_no, initial_flag, ship_date, " &
                "  char_50, transport_method, ship_plan_date, customer_order_line_no, " &
                "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                "  reconcile_type, imp_run_id, status, active_flag, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id, " &
                "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                ") VALUES (" &
                "  :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, " &
                "  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, " &
                "  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, " &
                "  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, " &
                "  :p_order_no, :p_remarks, :p_delivery_code, :p_order_time, :p_sales_unit_price, :p_delivery_time, " &
                "  :p_usage_location, :p_total_ship_qty, :p_production_category, :p_char_2, :p_container_no, " &
                "  :p_char_3, :p_char_4, :p_char_4_2, :p_char_5, :p_char_5_2, :p_char_6, " &
                "  :p_order_reason, :p_container_capacity, :p_customer_lot_no, :p_initial_flag, :p_ship_date, " &
                "  :p_char_50, :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
                "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, " &
                "  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag," &
                "  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, " &
                "  :p_created_at, :p_created_user_id, :p_created_pg_id, " &
                "  :p_updated_at, :p_updated_user_id, :p_updated_pg_id, " &
                "  :p_stra_order_qty, :p_stra_ship_qty, :p_stra_order_backlog " &
                ")"
                Using cmd As New OracleCommand(sql, conn)
#End If
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    For Each r In records
                        cmd.Parameters.Clear()

                        ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                        cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = r.CustomerSettingId ' NUMBER(10,0))
                        'cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
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

                        cmd.Parameters.Add(":p_order_time", OracleDbType.Decimal).Value = r.OrderTime       ' NUMBER(18,6)
                        cmd.Parameters.Add(":p_sales_unit_price", OracleDbType.Decimal).Value = r.SalesUnitPrice  ' NUMBER(18,6)
                        cmd.Parameters.Add(":p_delivery_time", OracleDbType.Decimal).Value = r.DeliveryTime    ' NUMBER(18,6)

                        cmd.Parameters.Add(":p_usage_location", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.UsageLocation, 45)
                        cmd.Parameters.Add(":p_total_ship_qty", OracleDbType.Decimal).Value = r.TotalShipQty
                        cmd.Parameters.Add(":p_production_category", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ProductionCategory, 45)
                        cmd.Parameters.Add(":p_char_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char2, 45)
                        cmd.Parameters.Add(":p_container_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ContainerNo, 45)

                        cmd.Parameters.Add(":p_char_3", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char3, 45)
                        cmd.Parameters.Add(":p_char_4", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4, 45)
                        cmd.Parameters.Add(":p_char_4_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4_2, 45)
                        cmd.Parameters.Add(":p_char_5", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5, 45)
                        cmd.Parameters.Add(":p_char_5_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5_2, 45)
                        cmd.Parameters.Add(":p_char_6", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char6, 45)

                        cmd.Parameters.Add(":p_order_reason", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderReason, 45)
                        cmd.Parameters.Add(":p_container_capacity", OracleDbType.Decimal).Value = r.ContainerCapacity
                        cmd.Parameters.Add(":p_customer_lot_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.CustomerLotNo, 45)
                        cmd.Parameters.Add(":p_initial_flag", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.InitialFlag, 45)

                        cmd.Parameters.Add(":p_ship_date", OracleDbType.Date).Value = r.ShipDate
                        cmd.Parameters.Add(":p_char_50", OracleDbType.Varchar2, 60).Value = SafeVarchar(r.Char50, 60)
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
#End If
        ' 受注一覧取得
        ' 2026/02/16 R.Sagisaka
        'Public Function GetOrders(
        '    Optional ByVal customerCode As String = Nothing,
        '    Optional ByVal customerName As String = Nothing,
        '    Optional ByVal profitCenter As String = Nothing,
        '    Optional ByVal customerUnitName As String = Nothing,
        '    Optional ByVal status As String = Nothing,
        '    Optional ByVal prodMgmtUserId As String = Nothing,
        '    Optional ByVal activeFlag As String = Nothing
        ') As DataTable

        '    Return GetOrders(customerCode, customerName, profitCenter, customerUnitName, status, prodMgmtUserId, activeFlag, Nothing)

        'End Function

        '''' <summary>
        '''' orderId から Record 取得
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="orderId"></param>
        '''' <returns></returns>
        'Public Function GetOrder(conn As OracleConnection, tran As OracleTransaction, orderId As Long) As DataRow
        '    Dim errorMessage = ""
        '    Dim dt As New DataTable()
        '    Try
        '        Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
        '            cmd.CommandText = "
        '                SELECT *
        '                FROM orders 
        '                WHERE oder_id = :p_oder_id "
        '            AddVarchar(cmd, ":p_oder_id", orderId)
        '            Using reader As OracleDataReader = cmd.ExecuteReader()
        '                dt.Load(reader)
        '            End Using
        '        End Using
        '    Catch e As OracleException
        '        errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
        '    Finally
        '    End Try

        '    If dt.Rows.Count = 1 Then
        '        Return dt.Rows(0)
        '    Else
        '        Return Nothing
        '    End If

        'End Function


        ''' <summary>
        ''' Orders Table から 条件でレコード抽出
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
            Optional ByVal prodMgmtUserId As String = Nothing,
            Optional ByVal additionalConditions As String = Nothing
        ) As DataTable

            Return GetOrders(conn, tran, OrdersTable.Orders, status, activeFlag, customerSettingId, orderId, prodMgmtUserId, additionalConditions)

        End Function
        ''' <summary>
        ''' Orders Table から 条件でレコード抽出
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
            Optional ByVal prodMgmtUserId As String = Nothing,
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

            If prodMgmtUserId IsNot Nothing Then
                Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
                If Not isAdmin Then
                    sb.AppendLine("AND UPPER(prod_mgmt_user_id) = :p_user ")
                    prm.Add(New OracleParameter(":p_user", OracleDbType.Varchar2) With {.Value = prodMgmtUserId})
                End If
            End If

            If Not String.IsNullOrEmpty(additionalConditions) Then
                sb.AppendLine(additionalConditions)
            End If
            Try

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try

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

        Public Function GetOrders(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal status As String = Nothing,
            Optional ByVal prodMgmtUserId As String = Nothing,
            Optional ByVal activeFlag As String = Nothing,
            Optional ByVal customerSettingId As Long = -1
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            'sb.AppendLine("  ORDER_ID                   AS ""OrderId"", ")
            sb.AppendLine("  customer_setting_id        AS ""CustomerSettingId"", ")
            sb.AppendLine("  customer_code              AS ""CustomerCode"", ")
            sb.AppendLine("  customer_name              AS ""CustomerName"", ")
            sb.AppendLine("  profit_center              AS ""ProfitCenter"", ")
            sb.AppendLine("  customer_unit_id           AS ""CustomerUnitId"", ")
            sb.AppendLine("  customer_unit_name         AS ""CustomerUnitName"", ")
            sb.AppendLine("  billing_to                 AS ""BillingTo"", ")
            sb.AppendLine("  customer_order_no          AS ""CustomerOrderNo"", ")
            sb.AppendLine("  demand_status              AS ""DemandStatus"", ")
            sb.AppendLine("  ship_to                    AS ""ShipTo"", ")
            sb.AppendLine("  order_date                 AS ""OrderDate"", ")
            sb.AppendLine("  due_date                   AS ""DueDate"", ")
            sb.AppendLine("  ship_scheduled_date        AS ""ShipScheduledDate"", ")
            sb.AppendLine("  customer_item_no           AS ""CustomerItemNo"", ")
            sb.AppendLine("  item_no                    AS ""ItemNo"", ")
            sb.AppendLine("  demand_qty                 AS ""DemandQty"", ")
            sb.AppendLine("  demand_unit                AS ""DemandUnit"", ")
            sb.AppendLine("  currency_code              AS ""CurrencyCode"", ")
            sb.AppendLine("  ship_stock_location        AS ""ShipStockLocation"", ")
            sb.AppendLine("  company_id                 AS ""CompanyId"", ")
            sb.AppendLine("  product_code               AS ""ProductCode"", ")
            sb.AppendLine("  billing_standard           AS ""BillingStandard"", ")
            sb.AppendLine("  ship_process_type          AS ""ship_process_type"", ")
            sb.AppendLine("  delivery_instr_flag        AS ""DeliveryInstrFlag"", ")
            sb.AppendLine("  order_no                   AS ""OrderNo"", ")
            sb.AppendLine("  remarks                    AS ""Remarks"", ")
            sb.AppendLine("  delivery_code              AS ""DeliveryCode"", ")
            sb.AppendLine("  order_time                 AS ""OrderTime"", ")
            sb.AppendLine("  sales_unit_price           AS ""SalesUnitPrice"", ")
            sb.AppendLine("  delivery_time              AS ""DeliveryTime"", ")
            sb.AppendLine("  usage_location             AS ""UsageLocation"", ")
            sb.AppendLine("  total_ship_qty             AS ""TotalShipQty"", ")
            sb.AppendLine("  production_category        AS ""ProductionCategory"", ")
            sb.AppendLine("  ship_date                  AS ""ShipDate"", ")
            sb.AppendLine("  transport_method           AS ""TransportMethod"", ")
            sb.AppendLine("  ship_plan_date             AS ""ShipPlanDate"", ")
            sb.AppendLine("  customer_order_line_no     AS ""CustomerOrderLineNo"", ")
            sb.AppendLine("  pre_daily_order_qty        AS ""PreDailyOrderQty"", ")
            sb.AppendLine("  pre_daily_delivery_date    AS ""PreDailyDeliveryDate"", ")
            sb.AppendLine("  imp_file_id                AS ""ImpFileId"", ")
            sb.AppendLine("  order_type                 AS ""OrderType"", ")
            sb.AppendLine("  prorated_type              AS ""ProratedType"", ")
            sb.AppendLine("  customer_info_type         AS ""CustomerInfoType"", ")
            sb.AppendLine("  info_type                  AS ""InfoType"", ")
            sb.AppendLine("  self_fcst_flag             AS ""SalfFcstFlag"", ")
            sb.AppendLine("  self_fcst_delete_flag      AS ""SalfFcstDeleteFlag"", ")
            sb.AppendLine("  reconcile_type             AS ""ReconcileType"", ")
            sb.AppendLine("  imp_run_id                 AS ""ImpRunId"", ")
            sb.AppendLine("  status                     AS ""Status"", ")
            sb.AppendLine("  active_flag                AS ""ActiveFlag"", ")
            sb.AppendLine("  created_at                 AS ""CreatedAt"", ")
            sb.AppendLine("  created_user_id            AS ""CreatedUserId"", ")
            sb.AppendLine("  created_user_id            AS ""CreatedUserId"", ")
            sb.AppendLine("  created_pg_id              AS ""CreatedPgId"", ")
            sb.AppendLine("  updated_at                 AS ""UpdatedAt"", ")
            sb.AppendLine("  updated_user_id            AS ""UpdatedUserId"", ")
            sb.AppendLine("  updated_pg_id              AS ""UpdatedPgId"",")
            sb.AppendLine("  prod_mgmt_user_id          AS ""ProdMgmtUserId"" ")
            ' Pharse2
            sb.AppendLine("  stra_order_qty             AS ""StraOrderQty"" ")
            sb.AppendLine("  stra_ship_qty              AS ""StraShipQty"" ")
            sb.AppendLine("  stra_order_backlog         AS ""StraOrderBacklog"" ")

            sb.AppendLine("FROM orders_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pCustomerCode As String = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName As String = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter As String = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName As String = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId As String = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)
            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())
            ' 2026/02/16 R.Sagisaka
            'Dim pCustomerSettingId As String = Utils.BuildLikePattern(customerSettingId, LikeMode.Contains)
            ' 2026/02/16 R.Sagisaka Add end

            If pCustomerCode IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_code) LIKE UPPER(:p_ccode) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_ccode", OracleDbType.Varchar2) With {.Value = pCustomerCode})
            End If

            If pCustomerName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_name) LIKE UPPER(:p_cname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cname", OracleDbType.Varchar2) With {.Value = pCustomerCode})
            End If

            If pProfitCenter IsNot Nothing Then
                sb.AppendLine("AND UPPER(profit_center) LIKE UPPER(:p_pc) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_pc", OracleDbType.Varchar2) With {.Value = pProfitCenter})
            End If

            If pCustomerUnitName IsNot Nothing Then
                sb.AppendLine("AND UPPER(customer_unit_name) LIKE UPPER(:p_cuname) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_cuname", OracleDbType.Varchar2) With {.Value = pCustomerUnitName})
            End If

            If status IsNot Nothing Then
                sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If pProdMgmtUserId IsNot Nothing Then
                Dim isAdmin As Boolean = String.Equals(prodMgmtUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
                If Not isAdmin Then
                    'sb.AppendLine("AND UPPER(prod_mgmt_user_id) = :p_user ")
                    sb.AppendLine("AND UPPER(prod_mgmt_user_id) LIKE UPPER(:p_user) ")
                    prm.Add(New OracleParameter(":p_user", OracleDbType.Varchar2) With {.Value = pProdMgmtUserId})
                End If
            End If

            If Not String.IsNullOrEmpty(pActiveFlag) Then
                'sb.AppendLine("AND UPPER(active_flag) = :p_active ")
                sb.AppendLine("AND UPPER(active_flag) = UPPER(:p_active) ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = pActiveFlag})
            End If

            ' 2026/02/16 R.Sagisaka
            If customerSettingId <> -1 Then
                sb.AppendLine("AND customer_setting_id = :p_csid ")
                prm.Add(New OracleParameter(":p_csid", OracleDbType.Int64) With {.Value = customerSettingId})
            End If
            ' 2026/02/16 R.Sagisaka Add end

            sb.AppendLine("ORDER BY created_at, customer_setting_id ")

            Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' 指定の customer setting id list のレコードを抽出
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <param name="customerSettingIdList"></param>
        ''' <returns></returns>
        Public Function GetRegisteredOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, status As String, customerSettingIdList As List(Of Long)) As DataTable

            Dim indexList = ""
            customerSettingIdList.ForEach(Sub(x) indexList &= $"{x}, ")
            Dim additionalParam As String = $"customer_setting_id IN ({indexList.TrimEnd(","c)}) "
            Return GetOrders(conn, tran, type, status:=status, additionalConditions:=additionalParam)

        End Function
        ''' <summary>
        ''' 一括 Update できる場合
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="kOrderId"></param>
        ''' <param name="kCustomerSettingId"></param>
        ''' <param name="kDemandQty"></param>
        ''' <param name="kStatus"></param>
        ''' <param name="orderId"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="demandQty"></param>
        ''' <param name="status"></param>
        ''' <returns></returns>
        Public Function Update(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                Optional ByVal kOrderStageId As Long? = Nothing,
                                Optional ByVal kOrderId As Long? = Nothing,
                                Optional ByVal kCustomerSettingId As Integer? = Nothing,
                                Optional ByVal kDemandQty As Long? = Nothing,
                                Optional ByVal kStatus As String = Nothing,
                                Optional ByVal kActiveFlag As String = Nothing,
                                Optional ByVal kShipScheduledDate As Date? = Nothing,
                                Optional ByVal orderId As Long? = Nothing,
                                Optional ByVal customerSettingId As Integer? = Nothing,
                                Optional ByVal demandQty As Long? = Nothing,
                                Optional ByVal status As String = Nothing,
                                Optional ByVal impFilesStageId As Long? = Nothing,
                                Optional ByVal impFileId As Long? = Nothing,
                                Optional ByVal impRunId As Long? = Nothing,
                                Optional ByVal activeFlag As String = Nothing,
                                Optional ByVal shipScheduledDate As Date? = Nothing,
                                Optional ByVal orderNo As String = Nothing,
                                Optional ByVal shipPlanDate As Date? = Nothing,
                                Optional ByVal createdAt As Date? = Nothing,
                                Optional ByVal createdUserId As String = Nothing,
                                Optional ByVal createdPgId As String = Nothing,
                                Optional ByVal updatedAt As Date? = Nothing,
                                Optional ByVal updatedUserId As String = Nothing,
                                Optional ByVal updatedPgId As String = Nothing
                                ) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                Dim prm As New List(Of OracleParameter)()
                sb.AppendLine($"UPDATE {GetTableName(type)} ")
                ' セット
                sb.AppendLine("SET ")
                If orderId IsNot Nothing Then
                    sb.AppendLine($" {GetIdName(type)} = :p_orderId, ")
                    prm.Add(New OracleParameter(":p_orderId", OracleDbType.Int64) With {.Value = orderId})
                End If
                If customerSettingId IsNot Nothing Then
                    sb.AppendLine(" customer_setting_id = :p_customerSettingId, ")
                    prm.Add(New OracleParameter(":p_customerSettingId", OracleDbType.Int64) With {.Value = customerSettingId})
                End If
                If demandQty IsNot Nothing Then
                    sb.AppendLine(" demand_qty = :p_demandQty, ")
                    prm.Add(New OracleParameter(":p_demandQty", OracleDbType.Int64) With {.Value = demandQty})
                End If

                If status IsNot Nothing Then
                    sb.AppendLine(" status = :p_status, ")
                    prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
                End If

                If impFilesStageId IsNot Nothing Then
                    sb.AppendLine(" imp_files_stage_id = :p_impFilesStageId, ")
                    prm.Add(New OracleParameter(":p_impFilesStageId", OracleDbType.Int64) With {.Value = impFilesStageId})
                End If

                If impFileId IsNot Nothing Then
                    sb.AppendLine(" imp_file_id = :p_impFileId, ")
                    prm.Add(New OracleParameter(":p_impFileId", OracleDbType.Int64) With {.Value = impFileId})
                End If

                If impRunId IsNot Nothing Then
                    sb.AppendLine(" imp_run_id = :p_impRunId, ")
                    prm.Add(New OracleParameter(":p_impRunId", OracleDbType.Long) With {.Value = impRunId})
                End If

                If activeFlag IsNot Nothing Then
                    sb.AppendLine(" active_flag = :p_activeFlag, ")
                    prm.Add(New OracleParameter(":p_activeFlag", OracleDbType.Varchar2) With {.Value = activeFlag})
                End If

                If createdAt IsNot Nothing Then
                    sb.AppendLine(" created_at = :p_createdAt, ")
                    prm.Add(New OracleParameter(":p_createdAt", OracleDbType.Date) With {.Value = createdAt})
                End If

                If createdUserId IsNot Nothing Then
                    sb.AppendLine(" created_user_id = :p_createdUserId, ")
                    prm.Add(New OracleParameter(":p_createdUserId", OracleDbType.Varchar2) With {.Value = createdUserId})
                End If

                If createdPgId IsNot Nothing Then
                    sb.AppendLine(" created_pg_id = :p_createdPgId, ")
                    prm.Add(New OracleParameter(":p_createdPgId", OracleDbType.Varchar2) With {.Value = createdPgId})
                End If

                If updatedAt IsNot Nothing Then
                    sb.AppendLine(" updated_at = :p_updatedAt, ")
                    prm.Add(New OracleParameter(":p_updatedAt", OracleDbType.Date) With {.Value = updatedAt})
                End If

                If updatedUserId IsNot Nothing Then
                    sb.AppendLine(" updated_user_id = :p_updatedUserId, ")
                    prm.Add(New OracleParameter(":p_updatedUserId", OracleDbType.Varchar2) With {.Value = updatedUserId})
                End If

                If updatedPgId IsNot Nothing Then
                    sb.AppendLine(" updated_pg_id = :p_updatedPgId, ")
                    prm.Add(New OracleParameter(":p_updatedPgId", OracleDbType.Varchar2) With {.Value = updatedPgId})
                End If
                If shipScheduledDate IsNot Nothing Then
                    sb.AppendLine(" ship_scheduled_date = :p_shipScheduledDate, ")
                    prm.Add(New OracleParameter(":p_shipScheduledDate", OracleDbType.Date) With {.Value = shipScheduledDate})
                End If

                If shipPlanDate IsNot Nothing Then
                    sb.AppendLine(" ship_plan_date = :p_shipPlanDate, ")
                    prm.Add(New OracleParameter(":p_shipPlanDate", OracleDbType.Date) With {.Value = shipPlanDate})
                End If

                ' 最後のカンマ削除
                If sb.Length > 0 Then
                    Dim startPos As Integer = Math.Max(0, sb.Length - 5)
                    sb.Replace(",", " ", startPos, sb.Length - startPos)
                End If

                sb.AppendLine(" WHERE 1=1 ")

                ' 絞り込み
                If kOrderStageId IsNot Nothing Then
                    sb.AppendLine("AND stage_id = :p_kOrderStageId ")
                    prm.Add(New OracleParameter(":p_kOrderStageId", OracleDbType.Int64) With {.Value = kOrderStageId})
                End If
                If kOrderId IsNot Nothing Then
                    sb.AppendLine($"AND {GetIdName(type)} = :p_kOrderId ")
                    prm.Add(New OracleParameter(":p_kOrderId", OracleDbType.Int64) With {.Value = kOrderId})
                End If

                If kCustomerSettingId IsNot Nothing Then
                    sb.AppendLine("AND customer_setting_id = :p_kCustomerSettingId ")
                    prm.Add(New OracleParameter(":p_kCustomerSettingId", OracleDbType.Int64) With {.Value = kCustomerSettingId})
                End If

                If kDemandQty IsNot Nothing Then
                    sb.AppendLine("AND demand_qty = :p_kDemandQty ")
                    prm.Add(New OracleParameter(":p_kDemandQty", OracleDbType.Int64) With {.Value = kDemandQty})
                End If

                If kStatus IsNot Nothing Then
                    sb.AppendLine("AND UPPER(status) = :p_kStatus ")
                    prm.Add(New OracleParameter(":p_kStatus", OracleDbType.Varchar2) With {.Value = kStatus})
                End If
                If kStatus IsNot Nothing Then
                    sb.AppendLine("AND UPPER(activeFlag) = :p_kActiveFlag ")
                    prm.Add(New OracleParameter(":p_kActiveFlag", OracleDbType.Varchar2) With {.Value = kActiveFlag})
                End If
                If kActiveFlag IsNot Nothing Then
                    sb.AppendLine("AND UPPER(active_Flag) = :p_kActiveFlag ")
                    prm.Add(New OracleParameter(":p_kActiveFlag", OracleDbType.Varchar2) With {.Value = kActiveFlag})
                End If
                If kShipScheduledDate IsNot Nothing Then
                    sb.AppendLine("AND ship_scheduled_date = :p_kShipScheduledDate ")
                    prm.Add(New OracleParameter(":p_kShipScheduledDate", OracleDbType.Date) With {.Value = kShipScheduledDate})
                End If

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    cmd.ExecuteNonQuery()
                End Using

            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function


        ''' <summary>
        ''' 納期設定
        ''' Order の id より
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="orderId"></param>
        ''' <param name="shipScheduledDate"></param>
        ''' <param name="shipDate"></param>
        ''' <param name="status"></param>
        ''' <param name="updateAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction,
                                        orderId As Integer,
                                        shipScheduledDate As Date,
                                        shipDate As Date,
                                        status As String,
                                        updateAt As Date,
                                        updatedUserId As String,
                                        updatedPgId As String
                                        ) As String

            Return UpdateDeadline(conn, tran, OrdersTable.Orders, orderId, shipScheduledDate, shipDate, status, updateAt, updatedUserId, updatedPgId)

        End Function

        ''' <summary>
        ''' Order の id より
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <param name="orderId"></param>
        ''' <param name="shipScheduledDate"></param>
        ''' <param name="shipDate"></param>
        ''' <param name="status"></param>
        ''' <param name="updatedAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                        orderId As Integer,
                                        shipScheduledDate As Date,
                                        shipDate As Date,
                                        status As String,
                                        updatedAt As Date,
                                        updatedUserId As String,
                                        updatedPgId As String
                                        ) As String

            ' SHIP_SCHEDULED_DATE(出荷予定日)
            ' SHIP_DATE(出荷日)
            ' STATUS(ステータス)
            ' UPDATED_AT(更新日時)
            ' UPDATED_USER_ID(更新ユーザーID)
            ' UPDATED_PG_ID(更新プログラムID)
            ' CUSTOMER_SETTING_ID(取引先設定ID)
            Dim errorMessage = ""
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}

                    cmd.CommandText = $"
                        UPDATE {GetTableName(type)} 
                         SET 
                              ship_scheduled_date = :p_ship_scheduled_date, 
                              ship_date           = :p_ship_date, 
                              status              = :p_status, 
                              updated_at          = :p_updated_at, 
                              updated_user_id     = :p_updated_user_id, 
                              updated_pg_id       = :p_updated_pg_id 
                          WHERE {GetIdName(type)} = :p_orderid 
                    "
                    AddDate(cmd, ":p_ship_scheduled_date", shipScheduledDate)
                    AddDate(cmd, ":p_ship_date", shipDate)
                    AddVarchar(cmd, ":p_status", status)
                    AddDate(cmd, ":p_updated_at", updatedAt)
                    AddVarchar(cmd, ":p_updated_user_id", updatedUserId)
                    AddVarchar(cmd, ":p_updated_pg_id", updatedPgId)
                    AddIntOrNull(cmd, ":p_orderid", orderId)

                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage
        End Function

#If False Then
        ''' <summary>
        ''' 納期設定
        ''' OrderStage の id より
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="stageId"></param>
        ''' <param name="due"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                        stageId As Integer,
                                        shipScheduledDate As Date,
                                        shipDate As Date,
                                        status As Integer,
                                        updateAt As Date,
                                        updatedUserId As String,
                                        updatedPgId As String,
                                        due As Boolean) As String

            ' SHIP_SCHEDULED_DATE(出荷予定日)
            ' SHIP_DATE(出荷日)
            ' STATUS(ステータス)
            ' UPDATED_AT(更新日時)
            ' UPDATED_USER_ID(更新ユーザーID)
            ' UPDATED_PG_ID(更新プログラムID)
            ' CUSTOMER_SETTING_ID(取引先設定ID)
            Dim errorMessage = ""
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                    cmd.CommandText = $"
                        UPDATE {GetTableName(type)} 
                         SET ship_scheduled_date = :p_ship_scheduled_date,
                              ship_date          = :p_ship_date,
                              status             = :p_status,
                              updated_at         = :p_update_at,
                              updated_user_id    = :p_updated_user_id,
                              updated_pg_id      = :p_updated_pg_id,
                          WHERE stage_id         = :p_stageid
                    "
                    AddDate(cmd, ":p_ship_scheduled_date", shipScheduledDate)
                    AddDate(cmd, ":p_ship_date", shipDate)
                    AddVarchar(cmd, ":p_status", status)
                    AddDate(cmd, ":p_updated_at", updateAt)
                    AddVarchar(cmd, ":p_updated_user_id", updatedUserId)
                    AddVarchar(cmd, ":p_updated_pg_id", updatedPgId)
                    AddIntOrNull(cmd, ":p_stageid", stageId)
                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage
        End Function
        ''' <summary>
        ''' 納期設定
        ''' Order の id より
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="orderId"></param>
        ''' <param name="shipScheduledDate"></param>
        ''' <param name="shipDate"></param>
        ''' <param name="status"></param>
        ''' <param name="updateAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                        orderId As Integer,
                                        shipScheduledDate As Date,
                                        shipDate As Date,
                                        status As String,
                                        updateAt As Date,
                                        updatedUserId As String,
                                        updatedPgId As String
                                        ) As String

            ' SHIP_SCHEDULED_DATE(出荷予定日)
            ' SHIP_DATE(出荷日)
            ' STATUS(ステータス)
            ' UPDATED_AT(更新日時)
            ' UPDATED_USER_ID(更新ユーザーID)
            ' UPDATED_PG_ID(更新プログラムID)
            ' CUSTOMER_SETTING_ID(取引先設定ID)
            Dim errorMessage = ""
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                    cmd.CommandText = $"
                        UPDATE {GetTableName(type)} 
                         SET ship_scheduled_date = :p_ship_scheduled_date,
                              ship_date          = :p_ship_date,
                              status             = :p_status,
                              updated_at         = :p_update_at,
                              updated_user_id    = :p_updated_user_id,
                              updated_pg_id      = :p_updated_pg_id,
                          WHERE order_id         = :p_orderid
                    "
                    AddDate(cmd, ":p_ship_scheduled_date", shipScheduledDate)
                    AddDate(cmd, ":p_ship_date", shipDate)
                    AddVarchar(cmd, ":p_status", status)
                    AddDate(cmd, ":p_updated_at", updateAt)
                    AddVarchar(cmd, ":p_updated_user_id", updatedUserId)
                    AddVarchar(cmd, ":p_updated_pg_id", updatedPgId)
                    AddIntOrNull(cmd, ":p_orderid", orderId)
                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage
        End Function
        ''' <summary>
        ''' 登録済みの Orders レコードを OrdersStage のstatus から検索し該当レコードを返す
        ''' ORDERS_STAGE（受注ワークテーブル）のSTATUS（ステータス）が[PLAN_SET]で登録されているレコードの
        ''' CUSTOMER_SETTING_ID（取引先設定ID）で結合し、ORDERS（受注テーブル）のSTATUS（ステータス）が[DUE_SET]で登録されているレコード
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <returns></returns>
        Public Function GetRegisteredOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable) As DataTable

            Dim errorMessage = ""
            Dim dt As New DataTable()
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                    cmd.CommandText = $"
                            SELECT 
                                o.* 
                            FROM 
                                {GetTableName(type)} o 
                            INNER JOIN 
                                order_stage os 
                            ON 
                                o.customer_setting_id = OS.customer_setting_id 
                            WHERE 
                                os.status = 'PLAN_SET' 
                                And o.status = 'DUE_SET'; 
                             "
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try

            If dt.Rows.Count > 1 Then
                Return dt
            Else
                Return Nothing
            End If

        End Function
#End If
        ''' <summary>
        ''' customerSettingId と status で OrderStage レコードを削除する
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="status"></param>
        ''' <returns></returns>
        Public Function Delete(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                Optional ByVal customerSettingId As Integer? = Nothing,
                                Optional ByVal demandQty As Long? = Nothing,
                                Optional ByVal status As String = Nothing,
                                Optional ByVal additionalParam As String = Nothing) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                sb.AppendLine("DELETE ")
                sb.AppendLine($"FROM {GetTableName(type)} ")
                sb.AppendLine("WHERE 1=1 ")
                Dim prm As New List(Of OracleParameter)()

                If customerSettingId IsNot Nothing Then
                    sb.AppendLine("AND customer_setting_id = :p_customerSettingId ")
                    prm.Add(New OracleParameter(":p_customerSettingId", OracleDbType.Int64) With {.Value = customerSettingId})
                End If
                If demandQty IsNot Nothing Then
                    sb.AppendLine("AND demand_Qty = :p_demandQty ")
                    prm.Add(New OracleParameter(":p_demandQty", OracleDbType.Int64) With {.Value = demandQty})
                End If
                If status IsNot Nothing Then
                    sb.AppendLine("AND UPPER(status) = :p_status ")
                    prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
                End If
                If (additionalParam IsNot Nothing) Then
                    sb.AppendLine(additionalParam)
                End If

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    cmd.ExecuteNonQuery()
                End Using

            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage

        End Function
        '''' <summary>
        '''' 登録済みの Orders レコードを OrdersStage のstatus から検索し該当レコードを削除する
        '''' ORDERS_STAGE（受注ワークテーブル）のSTATUS（ステータス）が[PLAN_SET]で登録されているレコードの
        '''' CUSTOMER_SETTING_ID（取引先設定ID）で結合し、ORDERS（受注テーブル）のSTATUS（ステータス）が[DUE_SET]で登録されているレコード
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <returns></returns>
        'Public Function DeleteRegisteredOrders(conn As OracleConnection, tran As OracleTransaction) As String

        '    Return DeleteRegisteredOrders(conn, tran, OrdersTable.Orders)

        'End Function
        ''' <summary>
        ''' Orders レコードをCUSTOMER_SETTING_ID（取引先設定ID）で検索し該当レコードを削除する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <returns></returns>
        Public Function DeleteRegisteredOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingIdList As List(Of Long)) As String

            Dim indexList = ""
            customerSettingIdList.ForEach(Sub(x) indexList &= $"{x},")
            Dim additionalParam As String = $"AND customer_setting_id IN ({indexList.TrimEnd(","c)}) "
            Return Delete(conn, tran, type, additionalParam:=additionalParam)

        End Function

        ''' <summary>
        ''' 登録済みの Orders レコードを OrdersStage のstatus から検索し該当レコードを削除する
        ''' ORDERS_STAGE（受注ワークテーブル）のSTATUS（ステータス）が[PLAN_SET]で登録されているレコードの
        ''' CUSTOMER_SETTING_ID（取引先設定ID）で結合し、ORDERS（受注テーブル）のSTATUS（ステータス）が[DUE_SET]で登録されているレコード
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <returns></returns>
        Public Function DeleteRegisteredOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable) As String

            Dim errorMessage = ""
            Dim dt As New DataTable()
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}

                    'Select Case count(*) FROM orders o
                    'WHERE EXISTS (
                    '                    SELECT 1 FROM orders_stage s 
                    '    WHERE s.customer_setting_id = o.customer_setting_id
                    ');

                    cmd.CommandText =
                    $"DELETE From {GetTableName(type)}
                        Where customer_setting_id IN (
                        Select Case customer_setting_id
                        FROM {GetStageTableName(type)}
                    ) "

                    'cmd.CommandText = $"
                    '        DELETE 
                    '            o
                    '        FROM 
                    '            {GetTableName(type)} o 
                    '        INNER JOIN 
                    '            order_stage os 
                    '        ON 
                    '            o.customer_setting_id = OS.customer_setting_id 
                    '        WHERE 
                    '            os.status = 'PLAN_SET' 
                    '            And o.status = 'DUE_SET'; 
                    '         "
                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try

            Return errorMessage

        End Function
        ''' <summary>
        ''' orders_stage の生産計画済みのレコードと同じOrders 側のレコードを削除する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <returns></returns>
        Public Function DeleteOrderStagePlanSetRecordByCustomerSettingId(conn As OracleConnection, tran As OracleTransaction)

            Return DeleteOrderStagePlanSetRecordByCustomerSettingId(conn, tran, OrdersTable.Orders)

        End Function
        ''' <summary>
        ''' orders_stage の生産計画済みのレコードと同じOrders 側のレコードを削除する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <returns></returns>
        Public Function DeleteOrderStagePlanSetRecordByCustomerSettingId(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable)
            Dim errorMessage = ""
            Dim dt As New DataTable()
            Try
                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                    cmd.CommandText = $"
                            Delete 
                            From 
                                {GetTableName(type)} 
                            Where 
                                status = 'PLAN_SET' 
                            And customer_setting_id 
                            In ( 
                                Select 
                                    customer_setting_id 
                                From 
                                    orders_stage 
                                Where 
                                    status = 'PLAN_SET' 
                            ); "
                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try

            Return errorMessage
        End Function

        ''' <summary>
        ''' DataRow to class
        ''' 名称
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As OrdersRow
            If dt Is Nothing Then
                Return Nothing
            End If

            Dim osr = New OrdersRow
            If (dt.Table.Columns.Contains("order_id")) Then
                osr.OrderId = dt.Field(Of Long)("order_id")
            ElseIf (dt.Table.Columns.Contains("prod_plan_id")) Then
                osr.OrderId = dt.Field(Of Long)("prod_plan_id")
            End If
            'osr.OrderId = dt.Field(Of Long)("order_id")
            osr.CustomerSettingId = dt.Field(Of Long)("customer_setting_id")
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
            ' Order のみ
            osr.UsageLocation = dt.Field(Of String)("usage_location")
            osr.ProductionCategory = dt.Field(Of String)("production_category")
            osr.Char2 = dt.Field(Of String)("char_2")
            osr.ContainerNo = dt.Field(Of String)("container_no")
            osr.Char3 = dt.Field(Of String)("char_3")
            osr.Char4 = dt.Field(Of String)("char_4")
            osr.Char4_2 = dt.Field(Of String)("char_4_2")
            osr.Char5 = dt.Field(Of String)("char_5")
            osr.Char5_2 = dt.Field(Of String)("char_5_2")
            osr.Char6 = dt.Field(Of String)("char_6")
            osr.OrderReason = dt.Field(Of String)("order_reason")
            osr.CustomerLotNo = dt.Field(Of String)("customer_lot_no")
            osr.InitialFlag = dt.Field(Of String)("initial_flag")
            osr.Char50 = dt.Field(Of String)("char_50")
            ' ====== 数値系 ======
            osr.DemandQty = dt.Field(Of Long?)("demand_qty")
            osr.TotalShipQty = dt.Field(Of Decimal?)("total_ship_qty")
            osr.PreDailyOrderQty = dt.Field(Of Decimal?)("pre_daily_order_qty")
            osr.ImpFileId = dt.Field(Of Long?)("imp_file_id")
            osr.OrderType = dt.Field(Of Int16?)("order_type")
            osr.ProratedType = dt.Field(Of Int16?)("prorated_type")
            osr.ReconcileType = dt.Field(Of Int16?)("reconcile_type")
            osr.ContainerCapacity = dt.Field(Of Decimal?)("container_capacity")
            ' Pharse2
            If (dt.Table.Columns.Contains("order_id")) Then
                osr.StraOrderQty = dt.Field(Of Decimal?)("stra_order_qty")
                osr.StraShipQty = dt.Field(Of Decimal?)("stra_ship_qty")
                osr.StraOrderBacklog = dt.Field(Of Decimal?)("stra_order_backlog")
            End If
            ' ====== 日付系 ======
            osr.OrderDate = dt.Field(Of Date?)("order_date")
            osr.DueDate = dt.Field(Of Date?)("due_date")
            osr.ShipScheduledDate = dt.Field(Of Date?)("ship_scheduled_date")
            osr.ShipDate = dt.Field(Of Date?)("ship_date")
            osr.ShipPlanDate = dt.Field(Of Date?)("ship_plan_date")
            osr.PreDailyDeliveryDate = dt.Field(Of Date?)("pre_daily_delivery_date")
            ' 監査系
            osr.CreatedAt = dt.Field(Of Date?)("created_at")
            osr.UpdatedAt = dt.Field(Of Date?)("updated_at")
            osr.CreatedUserId = dt.Field(Of String)("created_user_id")
            osr.CreatedPgId = dt.Field(Of String)("created_pg_id")
            osr.UpdatedUserId = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPgId = dt.Field(Of String)("updated_pg_id")

            Return osr

        End Function
        Private Sub FieldCheck(Row As DataRow, fieldName As String)
            Dim rawValue = Row(fieldName)
            Dim tt = rawValue.GetType().FullName
        End Sub
        ''' <summary>
        ''' DataRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As List(Of OrdersRow)
            Dim osrs = New List(Of OrdersRow)()
            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next
            Return osrs
        End Function


        ' 以下 新しく作らずに CustomerRepository で探すこと


        ''' <summary>
        ''' Get Customer data 
        ''' </summary>
        ''' <param name="customerSettingId"></param>
        ''' <returns></returns>
        Public Function GetCustomer(customerSettingId As Long) As Customer

            Dim custmer = New Customer

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT * ")
            sb.AppendLine("FROM customer_list_view ")
            'sb.AppendLine("WHERE customer_setting_id = :p_customerSettingId ")
            sb.AppendLine($"WHERE customer_setting_id = {customerSettingId} ")
            Try
                Using conn As New OracleConnection(_connectionString)
                    Using cmd As New OracleCommand(sb.ToString(), conn)
                        'Dim pId = cmd.Parameters.Add(":p_customerSettingId", OracleDbType.Int64).Value = customerSettingId
                        'cmd.BindByName = True
                        conn.Open()
                        Using reader As OracleDataReader = cmd.ExecuteReader()
                            dt.Load(reader)
                        End Using
                    End Using
                End Using

                Dim dr As DataRow = dt.Rows(0)
                custmer.Id = dr.Field(Of Long)("customer_setting_id")
                custmer.CustomerCode = dr.Field(Of String)("customer_code")
                custmer.CustomerUnitName = dr.Field(Of String)("customer_unit_name")
            Catch
            End Try
            Return custmer

        End Function
        ''' <summary>
        ''' 受注ファイル出力 ProdPlanStraView csv ファイル名フルパス取得
        ''' </summary>
        ''' <param name="fileBaseName"></param>
        ''' <param name="processDate"></param>
        ''' <returns></returns>
        Public Function GetProdPlanStraViewCsvFilename(path As String, fileBaseName As String, processDate As DateTime)
            Return IO.Path.Combine(path, GetProdPlanStraViewCsvFilename(fileBaseName, processDate))
        End Function
        ''' <summary>
        ''' 受注ファイル出力 ProdPlanStraView csv ファイル名取得
        ''' </summary>
        ''' <param name="fileBaseName"></param>
        ''' <param name="processDate"></param>
        ''' <returns></returns>
        Public Function GetProdPlanStraViewCsvFilename(fileBaseName As String, processDate As DateTime)
            Dim filename As String = ""
            filename = $"{fileBaseName}_{processDate:yyyyMMdd_HHmmss}.csv"
            Return filename
        End Function
        ''' <summary>
        ''' 受注ファイル出力 Zip ファイル名取得
        ''' </summary>
        ''' <param name="fileBaseName"></param>
        ''' <param name="processDate"></param>
        ''' <returns></returns>
        Public Function GeOrderZipFilename(fileBaseName As String, processDate As DateTime)
            Dim filename As String = ""
            filename = $"{fileBaseName}({processDate:yyyyMMddHHmmss}).zip"
            Return filename
        End Function
        ''' <summary>
        ''' 内示/確定 設定数を取得する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Public Function ProdPlanCount(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable) As (unofficialNotice As Integer, confirmed As Integer)

            Dim dt As New DataTable()
            Dim uc = 0
            Dim cc = 0
            ' 受注ファイル出力時 Prd_Plan で 内示処理と 確定処理を 行ったレコード数を 返す SQL
            Dim sql = "SELECT 
                         COUNT(CASE WHEN active_flag = 'Y' AND status = 'POST_PLAN_DUE_SET' AND demand_status = 'F' THEN 1 ELSE NULL END) AS UnofficialNotice,
                         COUNT(CASE WHEN active_flag = 'Y' AND status = 'POST_PLAN_DUE_SET' AND demand_status = 'O' THEN 1 ELSE NULL END) AS Confirmed
                        FROM 
                         prod_plan "
            Try
                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using

                    If (dt.Rows.Count = 1) Then
                        uc = dt.Rows(0).Field(Of Decimal)("UnofficialNotice")
                        cc = dt.Rows(0).Field(Of Decimal)("Confirmed")
                    End If
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try
            Return (uc, cc)
        End Function

        Public Function ProdPlanStraViewCsvFile(
            ByVal sql As String,
            ByVal delimiterCode As String,
            ByVal enclosureCode As String,
            ByVal headerYN As String,
            ByVal lineEndingCode As String,
            ByVal charsetCode As String,
            ByVal processDate As DateTime,
            Optional ByVal filename As String = Nothing,
            Optional ByVal connectionStringName As String = "OMSConnection",
            Optional ByVal parameters As IEnumerable(Of OdbcParameter) = Nothing,
            Optional ByVal endResponse As Boolean = True,
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        )
            ' コード → 実値へ解決
            Dim delimiter As String = Utils.MapDelimiter(delimiterCode)
            Dim enclosure As String = Utils.MapQuote(enclosureCode)
            Dim newline As String = Utils.MapNewline(lineEndingCode)
            Dim enc As Encoding = Utils.MapEncoding(charsetCode)

            Return ProdPlanStraViewCsvFile(sql, delimiter, enclosure, includeHeader:=String.Equals(headerYN, "Y", StringComparison.OrdinalIgnoreCase), newline, enc, processDate, filename, connectionStringName, parameters, formatEx)

        End Function
        ''' <summary>
        ''' ProdPlanStraView から csv ファイルを作成する
        ''' stream でDownload出力していた内容だが複数ファイル
        ''' のDownload 対応するためファイルを先に生成することになった
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <param name="delimiter"></param>
        ''' <param name="enclosure"></param>
        ''' <param name="includeHeader"></param>
        ''' <param name="newline"></param>
        ''' <param name="enc"></param>
        ''' <param name="processDate"></param>
        ''' <param name="filename"></param>
        ''' <param name="connectionStringName"></param>
        ''' <param name="parameters"></param>
        ''' <param name="formatEx"></param>
        ''' <returns></returns>
        Public Function ProdPlanStraViewCsvFile(ByVal sql As String,
            ByVal delimiter As String,
            ByVal enclosure As String,
            ByVal includeHeader As Boolean,
            ByVal newline As String,
            ByVal enc As Encoding,
            ByVal processDate As DateTime,
            ByVal filename As String,
            ByVal connectionStringName As String,
            ByVal parameters As IEnumerable(Of OracleParameter),
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        ) As String
            Dim rt = ""
            ' --- 検証（SELECTのみ許可）
            If String.IsNullOrWhiteSpace(sql) OrElse Not sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) Then
                Throw New ArgumentException("SELECT文のみ許可されます。", NameOf(sql))
            End If

            'Dim fileName As String = GetProdPlanStraViewCsvFilename(fileBaseName, processDate)
            'Dim fileName As String = $"{fileBaseName}_{processDate:yyyyMMdd_HHmmss}.csv"
            Try
                ' --- DB 接続・実行（先に例外を出させる）
                Dim connStr As String = ConfigurationManager.ConnectionStrings(connectionStringName).ConnectionString
                Using conn As New OracleConnection(connStr)
                    conn.Open()
                    Using cmd As New OracleCommand(sql, conn)
                        If parameters IsNot Nothing Then
                            For Each p In parameters
                                cmd.Parameters.Add(p)
                            Next
                        End If

                        Using rdr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)

                            Using fs As New FileStream(filename, FileMode.Create, FileAccess.Write)

                                ' UTF-8（BOM付き）の場合はPreambleを書き込み
                                Dim preamble As Byte() = enc.GetPreamble()
                                If preamble IsNot Nothing AndAlso preamble.Length > 0 Then
                                    fs.Write(preamble, 0, preamble.Length)
                                End If

                                ' --- ヘッダー行
                                If includeHeader Then
                                    Dim headerCols As New List(Of String)(rdr.FieldCount)
                                    For i As Integer = 0 To rdr.FieldCount - 1
                                        headerCols.Add(CsvFormat(rdr.GetName(i), delimiter, enclosure))
                                    Next

                                    Dim str = String.Join(delimiter, headerCols) & vbCrLf
                                    fs.Write(enc.GetBytes(str), 0, enc.GetByteCount(str))
                                End If

                                ' --- データ本体（逐次）
                                Dim rowCount As Integer = 0
                                While rdr.Read()
                                    Dim cols As New List(Of String)(rdr.FieldCount)
                                    For i As Integer = 0 To rdr.FieldCount - 1
                                        If rdr.IsDBNull(i) Then
                                            cols.Add("")
                                        Else
                                            Dim n = rdr.GetName(i)
                                            Dim t = rdr.GetFieldType(i)
                                            Dim v As Object = rdr.GetValue(i)
                                            Dim s As String
                                            Dim f As String = Nothing
                                            ' フィールド名 同じもののフォーマット情報を取得する
                                            If (formatEx IsNot Nothing) Then
                                                f = formatEx.Find(Function(x) x.name.ToUpper() = n.ToUpper()).format
                                            End If
                                            If t Is GetType(DateTime) Then
                                                f = If(f Is Nothing, "yyyy/MM/dd HH:mm:ss", f)
                                                s = CType(v, DateTime).ToString(f, CultureInfo.InvariantCulture)
                                            ElseIf GetType(IFormattable).IsAssignableFrom(t) Then
                                                f = If(f Is Nothing, "0.##########", f)
                                                s = DirectCast(v, IFormattable).ToString(f, CultureInfo.InvariantCulture)
                                            Else
                                                s = Convert.ToString(v, CultureInfo.InvariantCulture)
                                            End If
                                            cols.Add(CsvFormat(s, delimiter, enclosure))
                                        End If
                                    Next
                                    Dim str = String.Join(delimiter, cols) & vbCrLf

                                    fs.Write(enc.GetBytes(str), 0, enc.GetByteCount(str))

                                    rowCount += 1
                                    If (rowCount Mod 200) = 0 Then
                                        fs.Flush()
                                    End If
                                End While

                                fs.Flush()
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                ' ThreadAbortException は 無視する

                Dim m = ex.Message

            End Try

            Return rt

        End Function

        ''' <summary>
        ''' 生産計画 update
        ''' </summary>
        ''' <returns></returns>
        Public Function ProductionPlanUpdate(conn As OracleConnection, tran As OracleTransaction, updateDate As Date, userId As String) As String

            ' PROD_PLAN update
            ' PROD_PLAN_STAGE は PROD_PLAN のコピーですが、PROD_PLAN_ID フィールドを持ち PROD_PLANのレコードを特定できます。
            ' 1) PROD_PLAN_STAGE のフィールドが 下記の状態の時
            ' STATUS(ステータス)='EXPORTED'、ACTIVE_FLAG(有効フラグ)='Y'
            ' 2)該当する PROD_PLAN のレコードを更新します。
            ' STATUS(ステータス)='EXPORTED'
            ' UPDATED_AT(更新日時)=[現在時刻]
            ' UPDATED_USER_ID(更新ユーザーID)=外部から与える値
            ' UPDATED_PG_ID(更新プログラムID)='OrderExport'

            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("MERGE INTO PROD_PLAN target ")
            sb.AppendLine("USING ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        PROD_PLAN_ID ")
            sb.AppendLine("    FROM ")
            sb.AppendLine("        PROD_PLAN_STAGE ")
            sb.AppendLine("    WHERE ")
            sb.AppendLine("        STATUS = 'EXPORTED' ")
            sb.AppendLine("        AND ACTIVE_FLAG = 'Y' ")
            sb.AppendLine(") source ")
            sb.AppendLine("ON (target.PROD_PLAN_ID = source.PROD_PLAN_ID) ")
            sb.AppendLine("WHEN MATCHED THEN ")
            sb.AppendLine("UPDATE SET ")
            sb.AppendLine("    target.STATUS = 'EXPORTED', ")
            sb.AppendLine("    target.UPDATED_AT = :p_date, ")
            sb.AppendLine("    target.UPDATED_USER_ID = :p_user_id, ")
            sb.AppendLine("    target.UPDATED_PG_ID = 'OrderExport' ")

            Try
                'Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Parameters.Add(":p_date", OracleDbType.Date).Value = updateDate
                    cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = userId
                    'conn.Open()
                    'Using tran As OracleTransaction = conn.BeginTransaction()
                    Dim cnt = cmd.ExecuteNonQuery()
                    'tran.Commit()
                End Using
                'conn.Close()
                'End Using
                'End Using

            Catch ex As Exception
                errors = ex.Message
            End Try

            Return errors

        End Function

        ''' <summary>
        ''' 受注 Update Pharse-2
        ''' </summary>
        ''' <returns></returns>
        Public Function OrderUpdate(conn As OracleConnection, tran As OracleTransaction, updateDate As Date, userId As String) As String

            ' CUSTOMER_ORDER_NO（客先発注No）をキーとして、ORDERS（受注テーブル）をPROD_PLAN_STAGEの値に更新する。
            ' 1) ORDERS テーブルがターゲットで PROD_PLAN_STAGE がソーステーブルです。
            ' 2) PROD_PLAN_STAGE の フィールドが 次の時
            '  STATUS(ステータス) = 'EXPORTED'、ACTIVE_FLAG(有効フラグ) = 'Y'
            '  そのレコードのフィールド CUSTOMER_ORDER_NO（客先発注No）と
            '   ORDERS テーブルの CUSTOMER_ORDER_NO が同じレコードを更新します。
            ' 3) 更新内容は下記のとおりです。
            ' STATUS(ステータス)='EXPORTED'
            ' UPDATED_AT(更新日時)=[現在時刻]
            ' UPDATED_USER_ID(更新ユーザーID)=外部から渡す値
            ' UPDATED_PG_ID(更新プログラムID)='OrderExport'
            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("MERGE INTO ORDERS target ")
            sb.AppendLine("USING ( ")
            sb.AppendLine("    SELECT ")
            sb.AppendLine("        CUSTOMER_ORDER_NO ")
            sb.AppendLine("    FROM ")
            sb.AppendLine("        PROD_PLAN_STAGE ")
            sb.AppendLine("    WHERE ")
            sb.AppendLine("        STATUS = 'EXPORTED' ")
            sb.AppendLine("        AND ACTIVE_FLAG = 'Y' ")
            sb.AppendLine(") source ")
            sb.AppendLine("ON (target.CUSTOMER_ORDER_NO = source.CUSTOMER_ORDER_NO) ")
            sb.AppendLine("WHEN MATCHED THEN ")
            sb.AppendLine("UPDATE SET ")
            sb.AppendLine("    target.STATUS = 'EXPORTED', ")
            sb.AppendLine("    target.UPDATED_AT = ;p_date, ")
            sb.AppendLine("    target.UPDATED_USER_ID = :p_user_id, ")
            sb.AppendLine("    target.UPDATED_PG_ID = 'OrderExport' ")

            Try
                'Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Parameters.Add(":p_date", OracleDbType.Date).Value = updateDate
                    cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = userId
                    'conn.Open()
                    'Using tran As OracleTransaction = conn.BeginTransaction()
                    Dim cnt = cmd.ExecuteNonQuery()
                    'tran.Commit()
                End Using
                'conn.Close()
                'End Using
                'End Using

            Catch ex As Exception
                errors = ex.Message
            End Try

            Return errors


        End Function

    End Class
    ''' <summary>
    ''' Customer data class
    ''' </summary>
    Public Class Customer

        Public Property Id As Long?                         ' CUSTOMER_SETTING_ID	NUMBER(10)
        Public Property CustomerCode As String              ' CUSTOMER_CODE	VARCHAR2(25)
        Public Property CustomerUnitName As String          ' CUSTOMER_UNIT_NAME	VARCHAR2(50)

    End Class


    Public Class OrderGroup
        'CUSTOMER_SETTING_ID（取引先設定ID）
        'CUSTOMER_ITEM_NO（客先品目No）
        'ORDER_TYPE（受注区分）
        'PRORATED_TYPE（分割区分）
        'SHIP_SCHEDULED_DATE（出荷予定日）のyyyy/mm
        Public Property CustomerSettingId As Long               ' CUSTOMER_SETTING_ID NUMBER(10,0)
        Public Property CustomerItemNo As String                ' CUSTOMER_ITEM_NO VARCHAR2(20)
        Public Property OrderType As Integer?                   ' ORDER_TYPE NUMBER(1,0)
        Public Property ProratedType As Integer?                ' PRORATED_TYPE NUMBER(1,0)
        Public Property ShipScheduledDate As Date?              ' SHIP_SCHEDULED_DATE DATE (選別された yyyy/mm/01)
        ''' <summary>
        ''' constructor
        ''' </summary>
        Public Sub New()
        End Sub
        ''' <summary>
        ''' copy constructor
        ''' </summary>
        ''' <param name="customerSettingId"></param>
        ''' <param name="customerItemNo"></param>
        ''' <param name="orderType"></param>
        ''' <param name="proratedType"></param>
        Public Sub New(customerSettingId As Long, customerItemNo As String, orderType As Integer, proratedType As Integer, year As Integer, month As Integer)
            Me.CustomerSettingId = customerSettingId
            Me.CustomerItemNo = customerItemNo
            Me.OrderType = orderType
            Me.ProratedType = proratedType
            Me.ShipScheduledDate = DateSerial(year, month, 1)
        End Sub

    End Class
    ''' <summary>
    ''' ORDERS 受け渡し用の行DTO（Repository向け）
    ''' </summary>
    Public Class OrdersRow

        ' ====== 主キー／識別系 ======
        ' クラスの共通化のため 生産計画で使用する際は ProdPlanId を OrderId として使用する
        ' GetOrder で 取得した DataTable(DataRow)を ToClass で変換する際に Field名を指定して
        ' OrderId に取り込んでいる
        Public Property OrderId As Long?                        ' ORDER_ID NUMBER(10,0)
        'Public Property ProdPlanId As Long?                     ' ORDER_ID NUMBER(10,0)

        ' ====== 取引先・品目等 ======
        Public Property CustomerSettingId As Long               ' CUSTOMER_SETTING_ID NUMBER(10,0)
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
        Public Property UsageLocation As String                 ' USAGE_LOCATION VARCHAR2(45)
        Public Property ProductionCategory As String            ' PRODUCTION_CATEGORY VARCHAR2(45)
        Public Property Char2 As String                         ' CHAR_2 VARCHAR2(45)
        Public Property ContainerNo As String                   ' CONTAINER_NO VARCHAR2(45)
        Public Property Char3 As String                         ' CHAR_3 VARCHAR2(45)
        Public Property Char4 As String                         ' CHAR_4 VARCHAR2(45)
        Public Property Char4_2 As String                       ' CHAR_4_2 VARCHAR2(45)
        Public Property Char5 As String                         ' CHAR_5 VARCHAR2(45)
        Public Property Char5_2 As String                       ' CHAR_5_2 VARCHAR2(45)
        Public Property Char6 As String                         ' CHAR_6 VARCHAR2(45)
        Public Property OrderReason As String                   ' ORDER_REASON VARCHAR2(45)
        Public Property CustomerLotNo As String                 ' CUSTOMER_LOT_NO VARCHAR2(45)
        Public Property InitialFlag As String                   ' INITIAL_FLAG VARCHAR2(45)
        Public Property Char50 As String                        ' CHAR_50 VARCHAR2(60)
        Public Property TransportMethod As String               ' TRANSPORT_METHOD VARCHAR2(3)
        Public Property CustomerOrderLineNo As String           ' CUSTOMER_ORDER_LINE_NO VARCHAR2(2)
        Public Property CustomerInfoType As String              ' CUSTOMER_INFO_TYPE VARCHAR2(50)
        Public Property InfoType As String                      ' INFO_TYPE CHAR(1)
        Public Property SelfFcstFlag As String                  ' SELF_FCST_FLAG CHAR(1)
        Public Property SelfFcstDeleteFlag As String            ' SELF_FCST_DELETE_FLAG CHAR(1)
        Public Property ImpRunId As Long?                      ' IMP_RUN_ID NUMBER(10,0)
        Public Property Status As String                        ' STATUS VARCHAR2(20)
        Public Property ActiveFlag As String                    ' ACTIVE_FLAG CHAR(1)

        ' ====== 数値系 ======
        Public Property DemandQty As Long?                      ' DEMAND_QTY NUMBER(10,0)
        Public Property TotalShipQty As Decimal?                ' TOTAL_SHIP_QTY NUMBER(18,6)
        Public Property ContainerCapacity As Decimal?           ' CONTAINER_CAPACITY NUMBER(18,6)
        Public Property OrderTime As Decimal?                   ' ORDER_TIME NUMBER(18,6)
        Public Property SalesUnitPrice As Decimal?              ' SALES_UNIT_PRICE NUMBER(18,6)
        Public Property DeliveryTime As Decimal?                ' DELIVERY_TIME NUMBER(18,6)
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
        Public Property PreDailyDeliveryDate As Date?           ' PRE_DAILY_DELIVERY_DATE DATE

        ' 監査系
        Public Property CreatedAt As Date?                      ' CREATED_AT
        Public Property CreatedUserId As String                 ' CREATED_USER_ID
        Public Property CreatedPgId As String                   ' CREATED_PG_ID
        Public Property UpdatedAt As Date?                      ' UPDATED_AT
        Public Property UpdatedUserId As String                 ' UPDATED_USER_ID
        Public Property UpdatedPgId As String                   ' UPDATED_PG_ID

        Public Function IsCorrect() As Boolean

            'CUSTOMER_SETTING_ID（取引先設定ID）
            'CUSTOMER_CODE（取引先コード）
            'CUSTOMER_ORDER_NO（客先発注No）
            '※ORDER_TYPE（受注区分）が[2 Or 3]のレコードのみ
            'SHIP_TO（出荷先）
            'DUE_DATE（希望納期）
            'SHIP_SCHEDULED_DATE（出荷予定日）
            'CUSTOMER_ITEM_NO（客先品目No）
            'DEMAND_QTY（需要数）
            'CUSTOMER_ORDER_LINE_NO(客先発注No行番号)
            '※ORDER_TYPE（受注区分）が[2 Or 3]のレコードのみ
            'PRE_DAILY_ORDER_QTY（日割前受注数）
            'PRE_DAILY_DELIVERY_DATE（日割前納期）
            'ORDER_TYPE（受注区分）
            'PRORATED_TYPE(分割区分)

            If (CustomerSettingId = -1) Then
                Return False
            End If

            If (CustomerCode Is Nothing) Then
                Return False
            End If

            If (OrderType = -1) Then
                Return False
            End If
            ' 確定[ORDER_TYPE=2]、納入指示[ORDER_TYPE=3] の場合 必須
            If (OrderType = 2 And OrderType = 3) Then
                If (CustomerOrderNo Is Nothing) Then
                    Return False
                End If

                If (CustomerOrderLineNo Is Nothing) Then
                    Return False
                End If
            End If

            If (ShipTo Is Nothing) Then
                Return False
            End If

            If (DueDate = DateTime.MinValue) Then
                Return False
            End If

            If (ShipScheduledDate = DateTime.MinValue) Then
                Return False
            End If

            If (CustomerItemNo Is Nothing) Then
                Return False
            End If

            If (DemandQty = -1) Then
                Return False
            End If

            If (PreDailyOrderQty = -1) Then
                Return False
            End If

            If (PreDailyDeliveryDate = DateTime.MinValue) Then
                Return False
            End If

            If (ProratedType = -1) Then
                Return False
            End If

            Return True

        End Function


        ''' <summary>
        ''' OrdersRow リストを OrdersStageRow リストに変換をする
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        Public Shared Function ToOrdersStageRow(src As List(Of OrdersRow)) As List(Of OrdersStageRow)

            Dim dst = New List(Of OrdersStageRow)
            For Each st In src

                dst.Add(ToOrdersStageRow(st))
            Next
            Return dst

        End Function

        ''' <summary>
        ''' OrdersRow を OrdersStageRow に変換をする
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        Public Shared Function ToOrdersStageRow(src As OrdersRow) As OrdersStageRow
            Dim dst = New OrdersStageRow
            dst.OrderId = src.OrderId
            dst.CustomerSettingId = src.CustomerSettingId
            dst.CustomerCode = src.CustomerCode
            dst.BillingTo = src.BillingTo
            dst.CustomerOrderNo = src.CustomerOrderNo
            dst.DemandStatus = src.DemandStatus
            dst.ShipTo = src.ShipTo
            dst.CustomerItemNo = src.CustomerItemNo
            dst.ItemNo = src.ItemNo
            dst.DemandUnit = src.DemandUnit
            dst.CurrencyCode = src.CurrencyCode
            dst.ShipStockLocation = src.ShipStockLocation
            dst.CompanyId = src.CompanyId
            dst.ProductCode = src.ProductCode
            dst.BillingStandard = src.BillingStandard
            dst.ShipProcessType = src.ShipProcessType
            dst.DeliveryInstrFlag = src.DeliveryInstrFlag
            dst.OrderNo = src.OrderNo
            dst.Remarks = src.Remarks
            dst.DeliveryCode = src.DeliveryCode
            'dst.UsageLocation = src.UsageLocation
            'dst.ProductionCategory = src.ProductionCategory
            'dst.Char2 = src.Char2
            'dst.ContainerNo = src.ContainerNo
            'dst.Char3 = src.Char3
            'dst.Char4 = src.Char4
            'dst.Char4_2 = src.Char4_2
            'dst.Char5 = src.Char5
            'dst.Char5_2 = src.Char5_2
            'dst.Char6 = src.Char6
            'dst.OrderReason = src.OrderReason
            'dst.CustomerLotNo = src.CustomerLotNo
            'dst.InitialFlag = src.InitialFlag
            'dst.Char50 = src.Char50
            dst.TransportMethod = src.TransportMethod
            dst.CustomerOrderLineNo = src.CustomerOrderLineNo
            dst.CustomerInfoType = src.CustomerInfoType
            dst.InfoType = src.InfoType
            dst.SelfFcstFlag = src.SelfFcstFlag
            dst.SelfFcstDeleteFlag = src.SelfFcstDeleteFlag
            dst.ImpRunId = src.ImpRunId
            dst.Status = src.Status
            dst.ActiveFlag = src.ActiveFlag
            dst.DemandQty = src.DemandQty
            dst.TotalShipQty = src.TotalShipQty
            'dst.ContainerCapacity = src.ContainerCapacity
            'dst.OrderTime = src.OrderTime
            'dst.SalesUnitPrice = src.SalesUnitPrice
            'dst.DeliveryTime = src.DeliveryTime
            dst.PreDailyOrderQty = src.PreDailyOrderQty
            dst.ImpFileId = src.ImpFileId
            dst.OrderType = src.OrderType
            dst.ProratedType = src.ProratedType
            dst.ReconcileType = src.ReconcileType
            'Pharse2
            dst.StraOrderQty = src.StraOrderQty
            dst.StraShipQty = src.StraShipQty
            dst.StraOrderBacklog = src.StraOrderBacklog
            'Pharse2
            dst.OrderDate = src.OrderDate
            dst.DueDate = src.DueDate
            dst.ShipScheduledDate = src.ShipScheduledDate
            dst.ShipDate = src.ShipDate
            dst.ShipPlanDate = src.ShipPlanDate
            dst.PreDailyDeliveryDate = src.PreDailyDeliveryDate
            dst.CreatedAt = src.CreatedAt
            dst.CreatedUserId = src.CreatedUserId
            dst.CreatedPgId = src.CreatedPgId
            dst.UpdatedAt = src.UpdatedAt
            dst.UpdatedUserId = src.UpdatedUserId
            dst.UpdatedPgId = src.UpdatedPgId
            Return dst

        End Function
    End Class

End Namespace
