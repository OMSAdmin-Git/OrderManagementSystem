
Imports System.ComponentModel
Imports System.Data
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Microsoft.VisualBasic.ApplicationServices
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports DataTable = System.Data.DataTable

Namespace OMS.Data

    Public Class OrderStageRepository
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
        '''' <summary>
        '''' OrderRow class をDBに追加する
        '''' </summary>
        '''' <param name="row"></param>
        'Public Sub Insert(row As OrdersRow)

        '    Insert(OrdersTable.Orders, row)

        'End Sub
        '''' <summary>
        '''' OrderRow class をDBに追加する
        '''' </summary>
        '''' <param name="row"></param>
        'Public Sub Insert(type As OrdersTable, row As OrdersRow)

        '    Dim records As IEnumerable(Of OrdersRow) = {row}
        '    InsertRange(type, records)

        'End Sub
        '''' <summary>
        '''' OrderRow class をDBに追加する
        '''' R.sagisaka Add
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="row"></param>
        '''' <returns>string: Number:xxxx Errormessage</returns>
        'Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As OrdersRow) As String

        '    Return Insert(conn, tran, OrdersTable.Orders, row)

        'End Function
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
            Return InsertRange(conn, tran, type, OrdersRow.ToOrdersStageRow(records))

        End Function
        ''' <summary>
        ''' OrderRow class をDBに追加する
        ''' R.sagisaka Add
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, row As OrdersStageRow) As String

            Dim records = New List(Of OrdersStageRow)()
            records.Add(row)
            Return InsertRange(conn, tran, type, records)

        End Function
        '''' <summary>
        '''' OrderRow class リストをDBに追加する
        '''' </summary>
        '''' <param name="records"></param>
        'Public Sub InsertRange(records As IEnumerable(Of OrdersRow))
        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            InsertRange(conn, tran, OrdersTable.Orders, OrdersRow.ToOrdersStageRow(records.ToList()))
        '            tran.Commit()
        '        End Using
        '    End Using
        'End Sub
        '''' <summary>
        '''' OrderRow class リストをDBに追加する
        '''' </summary>
        '''' <param name="records"></param>
        'Public Sub InsertRange(type As OrdersTable, records As IEnumerable(Of OrdersRow))
        '    Using conn As New OracleConnection(_connectionString)
        '        conn.Open()
        '        Using tran As OracleTransaction = conn.BeginTransaction()
        '            InsertRange(conn, tran, type, OrdersRow.ToOrdersStageRow(records.ToList()))
        '            tran.Commit()
        '        End Using
        '    End Using
        'End Sub
        ''' <summary>
        ''' Order record 追加 (Excel から追加するとき orderRow レコードを stageに変換して)
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As IEnumerable(Of OrdersRow)) As String

            Return InsertRange(conn, tran, OrdersTable.Orders, records.ToList())

        End Function
        '''' <summary>
        '''' Order record 追加 (Excel から追加するとき orderRow レコード)
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="records"></param>
        '''' <returns>string: Number:xxxx Errormessage</returns>
        'Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, records As IEnumerable(Of OrdersRow)) As String

        '    Return InsertRange(conn, tran, type, records.ToList())

        'End Function
        ''' <summary>
        ''' Order record 追加 (Excel から追加するとき orderRow レコードを stageに変換して)
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of OrdersRow)) As String

            Return InsertRange(conn, tran, OrdersTable.Orders, records)

        End Function
        ''' <summary>
        ''' Order record 追加 (Excel から追加するとき orderRow レコードを stageに変換して)
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, records As List(Of OrdersRow)) As String

            Dim ordersRows = OrdersRow.ToOrdersStageRow(records)
            Return InsertRange(conn, tran, type, ordersRows)

        End Function

        ''' <summary>
        ''' Order record 追加
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of OrdersStageRow)) As String

            Return InsertRange(conn, tran, OrdersTable.Orders, records)

        End Function
        ''' <summary>
        ''' Order record 追加
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, records As List(Of OrdersStageRow)) As String

            Dim errorMessage As String = ""
            If records Is Nothing Then Return errorMessage
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO {GetTableName(type)} (")
                sb.AppendLine($"  {GetIdName(type)}, ")
                sb.AppendLine("  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, ")
                sb.AppendLine("  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, ")
                sb.AppendLine("  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, ")
                sb.AppendLine("  product_code, billing_standard, ship_process_type, delivery_instr_flag, ")
                sb.AppendLine("  order_no, remarks, delivery_code, ")
                sb.AppendLine("  ship_date, ")
                sb.AppendLine("  transport_method, ship_plan_date, customer_order_line_no, ")
                sb.AppendLine("  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, ")
                sb.AppendLine("  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, ")
                sb.AppendLine("  reconcile_type, imp_run_id, status, active_flag, ")
                sb.AppendLine("  created_at, created_user_id, created_pg_id, ")
                sb.AppendLine("  updated_at, updated_user_id, updated_pg_id ")
                If (type = OrdersTable.Orders) Then
                    sb.AppendLine(",  stra_order_qty, stra_ship_qty, stra_order_backlog ")
                End If
                sb.AppendLine(") VALUES (")
                sb.AppendLine("  :p_order_id, ")
                sb.AppendLine("  :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, ")
                sb.AppendLine("  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, ")
                sb.AppendLine("  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, ")
                sb.AppendLine("  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, ")
                sb.AppendLine("  :p_order_no, :p_remarks, :p_delivery_code, ")
                sb.AppendLine("  :p_ship_date, ")
                sb.AppendLine("  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, ")
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
                $"  {GetIdName(type)}, " &
                "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                "  order_no, remarks, " &
                "  ship_date, " &
                "  transport_method, ship_plan_date, customer_order_line_no, " &
                "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                "  reconcile_type, imp_run_id, status, active_flag, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id, " &
                "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                ") VALUES (" &
                "  :p_order_id, " &
                "  :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, " &
                "  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, " &
                "  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, " &
                "  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, " &
                "  :p_order_no, :p_remarks, " &
                "  :p_ship_date, " &
                "  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
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
                        cmd.Parameters.Add(":p_order_id", OracleDbType.Int64).Value = r.OrderId ' NUMBER(10,0))

                        ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                        cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = r.CustomerSettingId ' NUMBER(10,0))
                        'cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
                        cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerCode, 25)
                        cmd.Parameters.Add(":p_billing_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.BillingTo, 25)
                        cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 40).Value = SafeVarchar(r.CustomerOrderNo, 40)
                        'cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = NormalizeYN(r.DemandStatus) ' 1桁記号想定
                        cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = SafeVarchar(r.DemandStatus, 1)
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
                        'cmd.Parameters.Add(":p_ship_process_type", OracleDbType.Char, 1).Value = NormalizeYN(r.ShipProcessType)
                        cmd.Parameters.Add(":p_ship_process_type", OracleDbType.Char, 1).Value = SafeVarchar(r.ShipProcessType, 1)
                        cmd.Parameters.Add(":p_delivery_instr_flag", OracleDbType.Char, 1).Value = NormalizeYN(r.DeliveryInstrFlag)

                        cmd.Parameters.Add(":p_order_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderNo, 45)
                        cmd.Parameters.Add(":p_remarks", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Remarks, 45)
                        cmd.Parameters.Add(":p_delivery_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.DeliveryCode, 25)

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

#If False Then
            Dim errorMessage As String = ""
            If records Is Nothing Then Return errorMessage
            Try
                Const sql As String =
                "INSERT INTO orders_stage (" &
                "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                "  order_no, remarks, delivery_code, total_ship_qty, ship_date, " &
                "  transport_method, ship_plan_date, customer_order_line_no, " &
                "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                "  reconcile_type, imp_run_id, status, active_flag, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id " &
                ") VALUES (" &
                "  ;p_order_id, :p_customer_setting_id, :p_customer_code, :p_billing_to, :p_customer_order_no, :p_demand_status, :p_ship_to, " &
                "  :p_order_date, :p_due_date, :p_ship_scheduled_date, :p_customer_item_no, :p_item_no, " &
                "  :p_demand_qty, :p_demand_unit, :p_currency_code, :p_ship_stock_location, :p_company_id, " &
                "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, " &
                "  :p_product_code, :p_billing_standard, :p_ship_process_type, :p_delivery_instr_flag, " &
                "  :p_order_no, :p_remarks, :p_delivery_code, :p_total_ship_qty, :p_ship_date, " &
                "  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
                "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_id, " &
                "  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag, " &
                "  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, " &
                "  :p_created_at, :p_created_user_id, :p_created_pg_id, " &
                "  :p_updated_at, :p_updated_user_id, :p_updated_pg_id " &
                ")"

                Using cmd As New OracleCommand(sql, conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    For Each r In records
                        cmd.Parameters.Clear()

                        ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                        cmd.Parameters.Add(":p_order_id", OracleDbType.Int64).Value = r.OrderId ' NUMBER(10,0))
                        cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = r.CustomerSettingId ' NUMBER(10,0))
                        'cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
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
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            Finally
            End Try
            Return errorMessage
#End If
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
            Optional ByVal shipScheduledDate As Date? = Nothing,
            Optional ByVal additionalConditions As String = Nothing
        ) As DataTable
            Return GetOrders(conn, tran, OrdersTable.Orders, status, activeFlag, customerSettingId, orderId, shipScheduledDate, additionalConditions)
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
            Optional ByVal shipScheduledDate As Date? = Nothing,
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
                'sb.AppendLine("AND UPPER(status) LIKE UPPER(:p_status) ESCAPE '\' ")
                sb.AppendLine(" AND status = :p_status ")
                prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
            End If

            If Not String.IsNullOrEmpty(activeFlag) Then
                sb.AppendLine("AND UPPER(active_flag) = :p_active ")
                'sb.AppendLine("active_flag = :p_active ")
                prm.Add(New OracleParameter(":p_active", OracleDbType.Char) With {.Value = activeFlag})
            End If

            If shipScheduledDate IsNot Nothing Then
                sb.AppendLine("AND UPPER(ship_scheduled_date) = :p_shipScheduledDate ")
                prm.Add(New OracleParameter(":p_shipScheduledDate", OracleDbType.Date) With {.Value = shipScheduledDate})
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
        End Function
        ''' <summary>
        ''' 指定の customer setting id list のレコードを抽出
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <param name="customerSettingIdList"></param>
        ''' <returns></returns>
        Public Function GetRegisteredOrders(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, customerSettingIdList As List(Of Long)) As DataTable

            Dim indexList = ""
            customerSettingIdList.ForEach(Sub(x) indexList &= $"{x},")
            Dim additionalParam As String = $"AND customer_setting_id IN ({indexList.TrimEnd(","c)}) ORDER BY ship_scheduled_date "

            Return GetOrders(conn, tran, type, additionalConditions:=additionalParam)

        End Function
        'Public Function GetOrders(conn As OracleConnection, tran As OracleTransaction,
        '    Optional ByVal status As String = Nothing,
        '    Optional ByVal activeFlag As String = Nothing,
        '    Optional ByVal customerSettingId As Long = Nothing
        ') As DataTable

        '    Dim dt As New DataTable()
        '    Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
        '        cmd.CommandText = "
        '                SELECT *
        '                FROM orders_stage
        '                WHERE status = :p_status 
        '                AND active_flag    = :p_activeFlag 
        '                AND customer_Setting_Id    = :p_customerSettingId "
        '        cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(status, 20)
        '        cmd.Parameters.Add(":p_activeFlag", OracleDbType.Char).Value = activeFlag
        '        cmd.Parameters.Add(":p_customerSettingId", OracleDbType.Int64).Value = customerSettingId

        '        Using reader As OracleDataReader = cmd.ExecuteReader()
        '            dt.Load(reader)
        '        End Using
        '    End Using
        '    Return dt

        'End Function

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
        ''' customerSettingId と status で OrderStage レコードを削除する
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="status"></param>
        ''' <returns></returns>
        Public Function Delete(conn As OracleConnection, tran As OracleTransaction,
                                Optional ByVal customerSettingId As Integer? = Nothing,
                                Optional ByVal demandQty As Long? = Nothing,
                                Optional ByVal status As String = Nothing) As String
            Return Delete(conn, tran, OrdersTable.Orders, customerSettingId, demandQty, status)
        End Function

        ''' <summary>
        ''' Table 全レコード削除
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Public Function Truncate(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable) As String

            Dim errorMessage As String = ""
            Try
                Dim dt As New DataTable()
                Dim sb As New StringBuilder()
                Dim prm As New List(Of OracleParameter)()
                sb.AppendLine("TRUNCATE ")
                sb.AppendLine($"TABLE {GetTableName(type)} ")
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    cmd.ExecuteNonQuery()
                End Using
            Catch e As OracleException
                errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
            End Try
            Return errorMessage

        End Function

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
                                Optional ByVal status As String = Nothing) As String

            Dim errorMessage As String = ""
            Try
                'Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                'cmd.CommandText = "
                '    DELETE FROM orders_stage
                '    WHERE Customer_Setting_Id = :p_customerSettingId
                '    AND Status = :p_status
                '"
                'AddDecimal(cmd, ":customerSettingId", customerSettingId)
                'AddVarchar(cmd, ":status", status)
                'cmd.ExecuteNonQuery()
                'End Using

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
        '''' orderId から Order Stage Record 取得
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="orderId"></param>
        '''' <returns></returns>
        'Public Function GetOrderStage(conn As OracleConnection, tran As OracleTransaction, orderId As Long) As DataRow
        '    Dim errorMessage = ""
        '    Dim dt As New DataTable()
        '    Try
        '        Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
        '            cmd.CommandText = "
        '                SELECT *
        '                   fusrdec1 As AssortLeadTime
        '                FROM orders_stage 
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
        '''' <summary>
        '''' 指定の Orders レコードの　DemandQty に 0をセットする
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="orders"></param>
        '''' <returns></returns>
        'Public Function UpdateDemandQtyBlank(conn As OracleConnection, tran As OracleTransaction, orders As List(Of OrdersRow)) As String

        '    Dim errorMessage As String = ""
        '    For Each orderRow In orders
        '        errorMessage = Update(conn, tran, kOrderId:=orderRow.OrderId, demandQty:=0)
        '        If (errorMessage <> "") Then
        '            Exit For
        '        End If
        '    Next
        '    Return errorMessage

        'End Function

        '''' <summary>
        '''' OrderStageRows 配列のUpdate OrderId で
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="orderStageRows"></param>
        '''' <param name="status"></param>
        '''' <param name="impFilesStageId"></param>
        '''' <param name="impFileId"></param>
        '''' <param name="impRunId"></param>
        '''' <param name="activeFlag"></param>
        '''' <param name="createdAt"></param>
        '''' <param name="createdUserId"></param>
        '''' <param name="createdPgId"></param>
        '''' <param name="updatedAt"></param>
        '''' <param name="updatedUserId"></param>
        '''' <param name="updatedPgId"></param>
        '''' <returns></returns>
        'Public Function UpdateByOrderId(conn As OracleConnection, tran As OracleTransaction, orderStageRows As List(Of OrdersStageRow),
        '                        Optional ByVal status As String = Nothing,
        '                        Optional ByVal impFilesStageId As Long? = Nothing,
        '                        Optional ByVal impFileId As Long? = Nothing,
        '                        Optional ByVal impRunId As String = Nothing,
        '                        Optional ByVal activeFlag As String = Nothing,
        '                        Optional ByVal createdAt As Date? = Nothing,
        '                        Optional ByVal createdUserId As String = Nothing,
        '                        Optional ByVal createdPgId As String = Nothing,
        '                        Optional ByVal updatedAt As Date? = Nothing,
        '                        Optional ByVal updatedUserId As String = Nothing,
        '                        Optional ByVal updatedPgId As String = Nothing
        '                    ) As String
        '    Return UpdateByOrderId(conn, tran, orderStageRows, status, impFilesStageId, impFileId, impRunId, activeFlag, createdAt, createdUserId, createdPgId, updatedAt, updatedUserId, updatedPgId)
        'End Function

        ''' <summary>
        ''' OrderStageRows 配列のUpdate OrderId で
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="orderStageRows"></param>
        ''' <param name="status"></param>
        ''' <param name="impFileStageId"></param>
        ''' <param name="impFileId"></param>
        ''' <param name="impRunId"></param>
        ''' <param name="activeFlag"></param>
        ''' <param name="createdAt"></param>
        ''' <param name="createdUserId"></param>
        ''' <param name="createdPgId"></param>
        ''' <param name="updatedAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateByOrderId(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable, orderStageRows As List(Of OrdersStageRow),
                                Optional ByVal status As String = Nothing,
                                Optional ByVal impFileStageId As Long? = Nothing,
                                Optional ByVal impFileId As Long? = Nothing,
                                Optional ByVal impRunId As Long? = Nothing,
                                Optional ByVal activeFlag As String = Nothing,
                                Optional ByVal createdAt As Date? = Nothing,
                                Optional ByVal createdUserId As String = Nothing,
                                Optional ByVal createdPgId As String = Nothing,
                                Optional ByVal updatedAt As Date? = Nothing,
                                Optional ByVal updatedUserId As String = Nothing,
                                Optional ByVal updatedPgId As String = Nothing
                            ) As String
            Dim errors As String = ""
            For Each row In orderStageRows

                errors &= Update(conn, tran, type, kOrderId:=row.OrderId,
                                status:=status,
                                impFileStageId:=impFileStageId,
                                impFileId:=impFileId,
                                impRunId:=impRunId,
                                activeFlag:=activeFlag,
                                createdAt:=createdAt,
                                createdUserId:=createdUserId,
                                createdPgId:=createdPgId,
                                updatedAt:=updatedAt,
                                updatedUserId:=updatedUserId,
                                updatedPgId:=updatedPgId)
            Next
            Return errors

        End Function
        '''' <summary>
        '''' 一括 Update できる場合
        '''' </summary>
        '''' <param name="conn"></param>
        '''' <param name="tran"></param>
        '''' <param name="kOrderId"></param>
        '''' <param name="kCustomerSettingId"></param>
        '''' <param name="kDemandQty"></param>
        '''' <param name="kStatus"></param>
        '''' <param name="orderId"></param>
        '''' <param name="customerSettingId"></param>
        '''' <param name="demandQty"></param>
        '''' <param name="status"></param>
        '''' <returns></returns>
        'Public Function Update(conn As OracleConnection, tran As OracleTransaction,
        '                        Optional ByVal kOrderStageId As Long? = Nothing,
        '                        Optional ByVal kOrderId As Long? = Nothing,
        '                        Optional ByVal kCustomerSettingId As Integer? = Nothing,
        '                        Optional ByVal kDemandQty As Long? = Nothing,
        '                        Optional ByVal kStatus As String = Nothing,
        '                        Optional ByVal orderId As Long? = Nothing,
        '                        Optional ByVal customerSettingId As Integer? = Nothing,
        '                        Optional ByVal demandQty As Long? = Nothing,
        '                        Optional ByVal status As String = Nothing,
        '                        Optional ByVal impFilesStageId As Long? = Nothing,
        '                        Optional ByVal impFileId As Long? = Nothing,
        '                        Optional ByVal impRunId As String = Nothing,
        '                        Optional ByVal activeFlag As String = Nothing,
        '                        Optional ByVal createdAt As Date? = Nothing,
        '                        Optional ByVal createdUserId As String = Nothing,
        '                        Optional ByVal createdPgId As String = Nothing,
        '                        Optional ByVal updatedAt As Date? = Nothing,
        '                        Optional ByVal updatedUserId As String = Nothing,
        '                        Optional ByVal updatedPgId As String = Nothing
        '                        ) As String

        '    Return Update(conn, tran, OrdersTable.Orders,
        '                    kOrderStageId,
        '                    kOrderId,
        '                    kCustomerSettingId,
        '                    kDemandQty,
        '                    kStatus,
        '                    orderId,
        '                    customerSettingId,
        '                    demandQty,
        '                    status,
        '                    impFilesStageId,
        '                    impFileId,
        '                    impRunId,
        '                    activeFlag,
        '                    createdAt,
        '                    createdUserId,
        '                    createdPgId,
        '                    updatedAt,
        '                    updatedUserId,
        '                    updatedPgId)
        'End Function

        ''' <summary>
        ''' 一括 Update できる場合
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="type"></param>
        ''' <param name="kOrderStageId"></param>
        ''' <param name="kOrderId"></param>
        ''' <param name="kCustomerSettingId"></param>
        ''' <param name="kDemandQty"></param>
        ''' <param name="kStatus"></param>
        ''' <param name="kActiveFlag"></param>
        ''' <param name="kShipScheduledDate"></param>
        ''' <param name="orderId"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="demandQty"></param>
        ''' <param name="status"></param>
        ''' <param name="impFileStageId"></param>
        ''' <param name="impFileId"></param>
        ''' <param name="impRunId"></param>
        ''' <param name="activeFlag"></param>
        ''' <param name="shipScheduledDate"></param>
        ''' <param name="orderNo"></param>
        ''' <param name="ShipPlanDate"></param>
        ''' <param name="createdAt"></param>
        ''' <param name="createdUserId"></param>
        ''' <param name="createdPgId"></param>
        ''' <param name="updatedAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
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
                                Optional ByVal impFileStageId As Long? = Nothing,
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
                sb.AppendLine($"SET ")

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

                If impFileStageId IsNot Nothing Then
                    sb.AppendLine(" imp_file_stage_id = :p_impFileStageId, ")
                    prm.Add(New OracleParameter(":p_impFileStageId", OracleDbType.Int64) With {.Value = impFileStageId})
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
        '''' <summary>
        '''' 納期設定
        '''' OrderStage の id より
        '''' R.sagisaka create
        '''' </summary>
        '''' <param name="stageId"></param>
        '''' <param name="due"></param>
        '''' <returns></returns>
        'Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction,
        '                                stageId As Integer,
        '                                shipScheduledDate As Date,
        '                                shipDate As Date,
        '                                status As Integer,
        '                                updateAt As Date,
        '                                updatedUserId As String,
        '                                updatedPgId As String,
        '                                due As Boolean) As String

        '    ' SHIP_SCHEDULED_DATE(出荷予定日)
        '    ' SHIP_DATE(出荷日)
        '    ' STATUS(ステータス)
        '    ' UPDATED_AT(更新日時)
        '    ' UPDATED_USER_ID(更新ユーザーID)
        '    ' UPDATED_PG_ID(更新プログラムID)
        '    ' CUSTOMER_SETTING_ID(取引先設定ID)
        '    Dim errorMessage = ""
        '    Try
        '        Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
        '            cmd.CommandText = "
        '                UPDATE orders_stage 
        '                 SET ship_scheduled_date = :p_ship_scheduled_date,
        '                      ship_date          = :p_ship_date,
        '                      status             = :p_status,
        '                      updated_at         = :p_update_at,
        '                      updated_user_id    = :p_updated_user_id,
        '                      updated_pg_id      = :p_updated_pg_id,
        '                  WHERE stage_id         = :p_stageid
        '            "
        '            AddDate(cmd, ":p_ship_scheduled_date", shipScheduledDate)
        '            AddDate(cmd, ":p_ship_date", shipDate)
        '            AddVarchar(cmd, ":p_status", status)
        '            AddDate(cmd, ":p_updated_at", updateAt)
        '            AddVarchar(cmd, ":p_updated_user_id", updatedUserId)
        '            AddVarchar(cmd, ":p_updated_pg_id", updatedPgId)
        '            AddIntOrNull(cmd, ":p_stageid", stageId)
        '            cmd.ExecuteNonQuery()
        '        End Using
        '    Catch e As OracleException
        '        errorMessage = "Number: " & e.Number & vbCrLf & "Message: " & e.Message
        '    Finally
        '    End Try
        '    Return errorMessage
        'End Function
        ''' <summary>
        ''' 納期設定
        ''' Order の id より
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="orderId"></param>
        ''' <param name="shipScheduledDate"></param>
        ''' <param name="shipDate"></param>
        ''' <param name="shipPlanDate"></param>
        ''' <param name="dueDate"></param>
        ''' <param name="status"></param>
        ''' <param name="updatedAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction,
                                        Optional orderId As Integer? = Nothing,
                                        Optional shipScheduledDate As Date? = Nothing,
                                        Optional shipDate As Date? = Nothing,
                                        Optional shipPlanDate As Date? = Nothing,
                                        Optional dueDate As Date? = Nothing,
                                        Optional status As String = Nothing,
                                        Optional updatedAt As Date? = Nothing,
                                        Optional updatedUserId As String = Nothing,
                                        Optional updatedPgId As String = Nothing
                                        ) As String

            Return UpdateDeadline(conn, tran, OrdersTable.Orders, orderId, shipScheduledDate, shipDate, shipPlanDate, dueDate, status, updatedAt, updatedUserId, updatedPgId)

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
        ''' <param name="shipPlanDate"></param>
        ''' <param name="dueDate"></param>
        ''' <param name="status"></param>
        ''' <param name="updatedAt"></param>
        ''' <param name="updatedUserId"></param>
        ''' <param name="updatedPgId"></param>
        ''' <returns></returns>
        Public Function UpdateDeadline(conn As OracleConnection, tran As OracleTransaction, type As OrdersTable,
                                        Optional orderId As Integer? = Nothing,
                                        Optional shipScheduledDate As Date? = Nothing,
                                        Optional shipDate As Date? = Nothing,
                                        Optional shipPlanDate As Date? = Nothing,
                                        Optional dueDate As Date? = Nothing,
                                        Optional status As String = Nothing,
                                        Optional updatedAt As Date? = Nothing,
                                        Optional updatedUserId As String = Nothing,
                                        Optional updatedPgId As String = Nothing
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
                'Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}

                'cmd.CommandText = $"
                '    UPDATE {GetTableName(type)} 
                '     SET 
                '          ship_scheduled_date = :p_ship_scheduled_date, 
                '          ship_date           = :p_ship_date, 
                '          status              = :p_status, 
                '          updated_at          = :p_updated_at, 
                '          due_date            = :p_due_date,
                '          updated_user_id     = :p_updated_user_id, 
                '          updated_pg_id       = :p_updated_pg_id 
                '      WHERE {GetIdName(type)} = :p_orderid 
                '"
                'If (shipScheduledDate IsNot Nothing) Then
                '    AddDate(cmd, ":p_ship_scheduled_date", shipScheduledDate)
                'End If
                'AddDate(cmd, ":p_ship_date", shipDate)
                'AddVarchar(cmd, ":p_status", status)
                'AddDate(cmd, ":p_updated_at", updatedAt)
                'AddVarchar(cmd, ":p_updated_user_id", updatedUserId)
                'AddVarchar(cmd, ":p_updated_pg_id", updatedPgId)
                'AddIntOrNull(cmd, ":p_orderid", orderId)
                'cmd.ExecuteNonQuery()
                'End Using

                Dim sb As New StringBuilder()
                Dim prm As New List(Of OracleParameter)()

                sb.AppendLine($"UPDATE {GetTableName(type)} ")
                sb.AppendLine("SET ")

                If shipScheduledDate IsNot Nothing Then
                    sb.AppendLine("ship_scheduled_date = :p_ship_scheduled_date, ")
                    prm.Add(New OracleParameter(":p_ship_scheduled_date", OracleDbType.Date) With {.Value = shipScheduledDate})
                End If
                If shipDate IsNot Nothing Then
                    sb.AppendLine("ship_date = :p_ship_date, ")
                    prm.Add(New OracleParameter(":p_ship_date", OracleDbType.Date) With {.Value = shipDate})
                End If
                If shipPlanDate IsNot Nothing Then
                    sb.AppendLine("ship_plan_date = :p_ship_plan_date, ")
                    prm.Add(New OracleParameter(":p_ship_plan_date", OracleDbType.Date) With {.Value = shipPlanDate})
                End If
                If status IsNot Nothing Then
                    sb.AppendLine("status = :p_status, ")
                    prm.Add(New OracleParameter(":p_status", OracleDbType.Varchar2) With {.Value = status})
                End If
                If updatedAt IsNot Nothing Then
                    sb.AppendLine("updated_at = :p_updated_at, ")
                    prm.Add(New OracleParameter(":p_updated_at", OracleDbType.Date) With {.Value = updatedAt})
                End If
                If dueDate IsNot Nothing Then
                    sb.AppendLine("due_date = :p_due_date, ")
                    prm.Add(New OracleParameter(":p_due_date", OracleDbType.Date) With {.Value = dueDate})
                End If
                If updatedUserId IsNot Nothing Then
                    sb.AppendLine("updated_user_id = :p_updated_user_id, ")
                    prm.Add(New OracleParameter(":p_updated_user_id", OracleDbType.Varchar2) With {.Value = updatedUserId})
                End If
                If updatedPgId IsNot Nothing Then
                    sb.AppendLine("updated_pg_id = :p_updated_pg_id ")
                    prm.Add(New OracleParameter(":p_updated_pg_id", OracleDbType.Varchar2) With {.Value = updatedPgId})
                End If

                If orderId IsNot Nothing Then
                    sb.AppendLine($"WHERE {GetIdName(type)} = :p_orderid ")
                    prm.Add(New OracleParameter(":p_orderid", OracleDbType.Int64) With {.Value = orderId})
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
        ''' DataRow to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As OrdersStageRow

            If dt Is Nothing Then
                Return Nothing
            End If
            Dim osr = New OrdersStageRow
            If (dt.Table.Columns.Contains("order_id")) Then
                osr.OrderId = dt.Field(Of Long)("order_id")
            ElseIf (dt.Table.Columns.Contains("prod_plan_id")) Then
                osr.OrderId = dt.Field(Of Long)("prod_plan_id")
            End If
            'osr.OrderId = dt.Field(Of Long?)("order_id")
            osr.CustomerSettingId = dt.Field(Of Long?)("customer_setting_id")
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
            osr.TotalShipQty = dt.Field(Of Decimal?)("total_ship_qty")
            osr.PreDailyOrderQty = dt.Field(Of Decimal?)("pre_daily_order_qty")
            osr.ImpFileId = dt.Field(Of Long?)("imp_file_id")
            osr.OrderType = dt.Field(Of Int16?)("order_type")
            osr.ProratedType = dt.Field(Of Int16?)("prorated_type")
            osr.ReconcileType = dt.Field(Of Int16?)("reconcile_type")
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
            osr.CreatedAt = dt.Field(Of DateTime?)("created_at")
            osr.CreatedUserId = dt.Field(Of String)("created_user_id")
            osr.CreatedPgId = dt.Field(Of String)("created_pg_id")
            osr.UpdatedAt = dt.Field(Of DateTime?)("updated_at")
            osr.UpdatedUserId = dt.Field(Of String)("updated_user_id")
            osr.UpdatedPgId = dt.Field(Of String)("updated_pg_id")
            ' Stage 固有
            osr.StageId = dt.Field(Of Long?)("stage_id")
            osr.ImpFilesStageId = dt.Field(Of Long?)("imp_file_stage_id")

            Return osr

        End Function
        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As List(Of OrdersStageRow)

            Dim osrs = New List(Of OrdersStageRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As IEnumerable(Of DataRow)) As List(Of OrdersStageRow)

            Dim osrs = New List(Of OrdersStageRow)()

            For Each dtRow In dt
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

        '2026/03/26 酒井 st


        ''' <summary>
        ''' 受注ワークを削除する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="CustomerSettingId">処理中の取引先設定ID</param>
        ''' <param name="FolderType">処理中のフォルダ区分 1:内示2:確定3:納入指示4:混在</param>
        Public Sub DeleteRange(ByVal tran As OracleTransaction, ByVal CustomerSettingId As Long, ByVal FolderType As Long)

            'Dim pCustomerSettingId As String = If(String.IsNullOrWhiteSpace(CustomerSettingId), Nothing, CustomerSettingId.Trim())
            'If CustomerSettingId <= 0 Then
            '    Return
            'End If
            Dim Where As String = ""

            Select Case FolderType
                Case 1
                    Where = "AND order_type = 1"
                Case 2
                    Where = "AND order_type = 2"
                Case 3
                    Where = "AND order_type = 3"
                Case 4
                    Where = ""
            End Select

            Dim sql As String = $"
                         DELETE FROM orders_stage 
                         WHERE 1=1 
                         {Where}
                         AND customer_setting_id = :p_customer_setting_id"



            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = CustomerSettingId
                'cmd.Parameters.Add(":p_folder_type", OracleDbType.Int32).Value = FolderType
                cmd.ExecuteNonQuery()

            End Using

        End Sub

        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        Public Function GetProratedType(ByVal CustomerSettingId As Long, ByVal FolderType As Long) As Integer

            Const sql As String =
                            " SELECT prorated_type FROM imp_rule_mst " &
                            " WHERE 1=1 " &
                            " AND customer_setting_id = :p_customer_setting_id " &
                            " AND folder_type = :p_folder_type "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = CustomerSettingId
                    cmd.Parameters.Add(":p_folder_type", OracleDbType.Int64).Value = FolderType

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        Public Function GetCurrencyCode(ByVal CustomerCode As String) As String

            Dim pCustomerCode As String = If(String.IsNullOrWhiteSpace(CustomerCode), Nothing, CustomerCode.Trim())

            Const sql As String =
                        " SELECT fcurr FROM sectm " &
                        " WHERE fsectcd = :p_customer_code "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode, 25)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        'Public Function GetProductCode(ByVal CustomerItemNo As String) As String
        Public Function GetProductCode(ByVal CustomerItemNo As String, ByVal CustomerCode As String) As String

            Dim pCustomerItemNo As String = If(String.IsNullOrWhiteSpace(CustomerItemNo), Nothing, CustomerItemNo.Trim())
            Dim pCustomerCode As String = If(String.IsNullOrWhiteSpace(CustomerCode), Nothing, CustomerCode.Trim())

            'Const sql As String =
            '            " SELECT fprdcd FROM prdslsodrm " &
            '            " WHERE fcustitemno = :p_customer_item_no "
            Const sql As String =
                        " SELECT fprdcd FROM prdslsodrm " &
                        " WHERE fcustitemno = :p_customer_item_no " &
                        " AND fcustcd = :p_customer_code "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(pCustomerItemNo, 45)
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode, 25)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        Public Function GetDemandUnit(ByVal ProductCode As String) As String

            Dim pProductCode As String = If(String.IsNullOrWhiteSpace(ProductCode), Nothing, ProductCode.Trim())

            Const sql As String =
                        " SELECT funit FROM itemm " &
                        " WHERE fitemno = :p_product_code "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_product_code", OracleDbType.Varchar2, 45).Value = SafeVarchar(pProductCode, 45)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        Public Function GetShipStockLocation(ByVal CustomerCode As String, ByVal DeliveryCode As String) As String

            Dim pCustomerCode As String = If(String.IsNullOrWhiteSpace(CustomerCode), Nothing, CustomerCode.Trim())
            Dim pDeliveryCode As String = If(String.IsNullOrWhiteSpace(DeliveryCode), Nothing, DeliveryCode.Trim())
            Dim currentfuppsect As String = Nothing   '上位部署

            Const sql As String =
                        " SELECT fuppsect FROM sectm " &
                        " WHERE fsectcd = :p_customer_code "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode, 25)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    currentfuppsect = Convert.ToString(obj)

                End Using

                If currentfuppsect IsNot Nothing Then

                    Const sql2 As String =
                        " SELECT fshpwhcd FROM shproutm " &
                        " WHERE fshptocd = :p_customer_code " &
                        " AND fpriority = :p_fuppsect "

                    Using cmd As New OracleCommand(sql2, conn)

                        cmd.BindByName = True
                        cmd.CommandType = CommandType.Text

                        cmd.Parameters.Clear()
                        cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode & pDeliveryCode, 25)
                        If currentfuppsect <> "F" Then
                            cmd.Parameters.Add(":p_fuppsect", OracleDbType.Int64).Value = 1
                        Else
                            cmd.Parameters.Add(":p_fuppsect", OracleDbType.Int64).Value = 2
                        End If

                        Dim obj = cmd.ExecuteScalar()
                        If obj Is Nothing OrElse obj Is DBNull.Value Then
                            Return Nothing
                        End If
                        Return Convert.ToString(obj)


                    End Using

                End If

                Return Nothing

            End Using

        End Function

        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        Public Function GetInfoType(ByVal CustomerSettingId As Long, ByVal CustomerInfoType As String) As String

            Dim pCustomerInfoType As String = If(String.IsNullOrWhiteSpace(CustomerInfoType), Nothing, CustomerInfoType.Trim())

            Const sql As String =
                            " SELECT info_type FROM info_type_mst " &
                            " WHERE 1=1 " &
                            " AND customer_setting_id = :p_customer_setting_id " &
                            " AND customer_info_type = :p_customer_info_type "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = CustomerSettingId
                    cmd.Parameters.Add(":p_customer_info_type", OracleDbType.Varchar2, 50).Value = SafeVarchar(pCustomerInfoType, 50)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        Public Function GetReconcileType(ByVal CustomerSettingId As Long, ByVal FolderType As Long) As Integer

            Const sql As String =
                            " SELECT reconcile_type FROM imp_rule_mst " &
                            " WHERE 1=1 " &
                            " AND customer_setting_id = :p_customer_setting_id " &
                            " AND folder_type = :p_folder_type "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = CustomerSettingId
                    cmd.Parameters.Add(":p_folder_type", OracleDbType.Int64).Value = FolderType

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        ''' <summary>
        ''' 出荷先を取得する
        ''' </summary>
        ''' <param name="CustomerCode">処理中の取引先コード</param>
        ''' <param name="DeliveryCode">処理中の納入先コード</param>
        Public Function GetShipTo(ByVal CustomerCode As String, ByVal DeliveryCode As String) As String

            Dim pCustomerCode As String = If(String.IsNullOrWhiteSpace(CustomerCode), Nothing, CustomerCode.Trim())
            Dim pDeliveryCode As String = If(String.IsNullOrWhiteSpace(DeliveryCode), Nothing, DeliveryCode.Trim())

            '2026/06/01 酒井 st
            'Const sql As String =
            '            " SELECT fsectcd FROM sectd " &
            '            " WHERE fsecttyp = 'ST' " &
            '            " AND fsectcd = :p_customer_code "
            Const sql As String =
                        " SELECT sectm.fsectcd FROM sectm " &
                        " INNER JOIN sectd ON sectm.fsectcd = sectd.fsectcd " &
                        " AND sectd.fsecttyp = 'ST' " &
                        " WHERE 1=1 " &
                        " AND sectm.fsectcd = :p_customer_code "
            '2026/06/01 酒井 ed

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode & pDeliveryCode, 25)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        '2026/06/01 酒井 st
        ''' <summary>
        ''' 請求先を取得する
        ''' </summary>
        ''' <param name="CustomerCode">処理中の取引先コード</param>
        Public Function GetBillingTo(ByVal CustomerCode As String) As String

            Dim pCustomerCode As String = If(String.IsNullOrWhiteSpace(CustomerCode), Nothing, CustomerCode.Trim())

            Const sql As String =
                        " SELECT sectm.fbilltocd FROM sectm " &
                        " INNER JOIN sectd ON sectm.fsectcd = sectd.fsectcd " &
                        " AND sectd.fsecttyp = 'CU' " &
                        " WHERE 1=1 " &
                        " AND sectm.fsectcd = :p_customer_code "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(pCustomerCode, 25)

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function
        '2026/06/01 酒井 ed

        Public Function GetBillingTo() As String

            Const sql As String =
                        " SELECT sectm.fbilltocd FROM sectm " &
                        " INNER JOIN sectd ON sectm.fsectcd = sectd.fsectcd " &
                        " WHERE sectd.fsecttyp = 'CU' "

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql, conn)

                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Clear()

                    Dim obj = cmd.ExecuteScalar()
                    If obj Is Nothing OrElse obj Is DBNull.Value Then
                        Return Nothing
                    End If
                    Return Convert.ToString(obj)

                End Using
            End Using

        End Function

        Public Function InsertRange(ByVal tran As OracleTransaction, records As IEnumerable(Of OrdersStageRow)) As Integer
            'If records Is Nothing Then Return

            ' 実行
            Dim insertCount As Integer = 0

            Const sql As String =
                        "INSERT INTO orders_stage (" &
                        "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  order_no, remarks, delivery_code, " &
                        "  total_ship_qty, " &
                        "  ship_date, " &
                        "  transport_method, ship_plan_date, customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_stage_id, imp_file_id, " &
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
                        "  :p_order_no, :p_remarks, :p_delivery_code, " &
                        "  :p_total_ship_qty, " &
                        "  :p_ship_date, " &
                        "  :p_transport_method, :p_ship_plan_date, :p_customer_order_line_no, " &
                        "  :p_pre_daily_order_qty, :p_pre_daily_delivery_date, :p_imp_file_stage_id, :p_imp_file_id, " &
                        "  :p_order_type, :p_prorated_type, :p_customer_info_type, :p_info_type, :p_self_fcst_flag, :p_self_fcst_delete_flag, " &
                        "  :p_reconcile_type, :p_imp_run_id, :p_status, :p_active_flag, " &
                        "  :p_created_at, :p_created_user_id, :p_created_pg_id, " &
                        "  :p_updated_at, :p_updated_user_id, :p_updated_pg_id, " &
                        "  :p_stra_order_qty, :p_stra_ship_qty, :p_stra_order_backlog " &
                        ")"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                For Each r In records
                    cmd.Parameters.Clear()

                    ' 例：文字列は SafeVarchar で桁超を丸め（定義長に合わせる）
                    'cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerSettingId, 25)
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = r.CustomerSettingId ' NUMBER(10,0)
                    cmd.Parameters.Add(":p_customer_code", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.CustomerCode, 25)
                    cmd.Parameters.Add(":p_billing_to", OracleDbType.Varchar2, 25).Value = SafeVarchar(r.BillingTo, 25)
                    cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 40).Value = SafeVarchar(r.CustomerOrderNo, 40)
                    'cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = NormalizeYN(r.DemandStatus) ' 1桁記号想定
                    cmd.Parameters.Add(":p_demand_status", OracleDbType.Char, 1).Value = SafeVarchar(r.DemandStatus, 1)
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

                    'cmd.Parameters.Add(":p_order_time", OracleDbType.Decimal).Value = r.OrderTime       ' NUMBER(18,6)
                    'cmd.Parameters.Add(":p_sales_unit_price", OracleDbType.Decimal).Value = r.SalesUnitPrice  ' NUMBER(18,6)
                    'cmd.Parameters.Add(":p_delivery_time", OracleDbType.Decimal).Value = r.DeliveryTime    ' NUMBER(18,6)
                    'cmd.Parameters.Add(":p_usage_location", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.UsageLocation, 45)

                    cmd.Parameters.Add(":p_total_ship_qty", OracleDbType.Decimal).Value = r.TotalShipQty
                    'cmd.Parameters.Add(":p_production_category", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ProductionCategory, 45)
                    'cmd.Parameters.Add(":p_char_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char2, 45)
                    'cmd.Parameters.Add(":p_container_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ContainerNo, 45)

                    'cmd.Parameters.Add(":p_char_3", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char3, 45)
                    'cmd.Parameters.Add(":p_char_4", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4, 45)
                    'cmd.Parameters.Add(":p_char_4_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char4_2, 45)
                    'cmd.Parameters.Add(":p_char_5", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5, 45)
                    'cmd.Parameters.Add(":p_char_5_2", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char5_2, 45)
                    'cmd.Parameters.Add(":p_char_6", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.Char6, 45)

                    'cmd.Parameters.Add(":p_order_reason", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.OrderReason, 45)
                    'cmd.Parameters.Add(":p_container_capacity", OracleDbType.Decimal).Value = r.ContainerCapacity
                    'cmd.Parameters.Add(":p_customer_lot_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.CustomerLotNo, 45)
                    'cmd.Parameters.Add(":p_initial_flag", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.InitialFlag, 45)

                    cmd.Parameters.Add(":p_ship_date", OracleDbType.Date).Value = r.ShipDate
                    'cmd.Parameters.Add(":p_char_50", OracleDbType.Varchar2, 60).Value = SafeVarchar(r.Char50, 60)
                    cmd.Parameters.Add(":p_transport_method", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.TransportMethod, 3)
                    cmd.Parameters.Add(":p_ship_plan_date", OracleDbType.Date).Value = r.ShipPlanDate
                    cmd.Parameters.Add(":p_customer_order_line_no", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CustomerOrderLineNo, 2)

                    cmd.Parameters.Add(":p_pre_daily_order_qty", OracleDbType.Decimal).Value = r.PreDailyOrderQty
                    cmd.Parameters.Add(":p_pre_daily_delivery_date", OracleDbType.Date).Value = r.PreDailyDeliveryDate

                    cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = r.ImpFileStageId ' NUMBER(10,0)
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
                    cmd.Parameters.Add(":p_stra_order_qty", OracleDbType.Decimal).Value = r.StraOrderQty
                    cmd.Parameters.Add(":p_stra_ship_qty", OracleDbType.Decimal).Value = r.StraShipQty
                    cmd.Parameters.Add(":p_stra_order_backlog", OracleDbType.Decimal).Value = r.StraOrderBacklog

                    cmd.ExecuteNonQuery()


                Next

                insertCount = records.Count

            End Using

            Return insertCount

        End Function

        ''' <summary>
        ''' 受注ワーク(orders_stage)に正規(orders)のレコードを追加する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="CustomerSettingId">処理中の取引先設定ID</param>
        Public Sub InsertStageFromOrders(ByVal tran As OracleTransaction, ByVal CustomerSettingId As Long)

            If CustomerSettingId <= 0 Then
                Return
            End If

            'Const sql As String =
            '            "INSERT INTO orders_stage (" &
            '            "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
            '            "  ship_plan_date, customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id" &
            '            ") SELECT " &
            '            "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
            '            "  ship_plan_date, customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id " &
            '            "  FROM orders " &
            '            "  WHERE customer_setting_id = :p_customer_setting_id " &
            '            "  AND order_type IN (1,2) " &
            '            "  AND active_flag = 'Y' "
            Const sql As String =
                "INSERT INTO orders_stage (" &
                "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
                "  ship_plan_date, customer_order_line_no, " &
                "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                "  reconcile_type, imp_run_id, status, active_flag, " &
                "  created_at, created_user_id, created_pg_id, " &
                "  updated_at, updated_user_id, updated_pg_id, " &
                "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                ") SELECT " &
                "  o.order_id, o.customer_setting_id, o.customer_code, o.billing_to, o.customer_order_no, o.demand_status, o.ship_to, " &
                "  o.order_date, o.due_date, o.ship_scheduled_date, o.customer_item_no, o.item_no, " &
                "  o.demand_qty, o.demand_unit, o.currency_code, o.ship_stock_location, o.company_id, " &
                "  o.product_code, o.billing_standard, o.ship_process_type, o.delivery_instr_flag, " &
                "  o.order_no, o.remarks, o.delivery_code, o.total_ship_qty, o.ship_date, o.transport_method, " &
                "  o.ship_plan_date, o.customer_order_line_no, " &
                "  o.pre_daily_order_qty, o.pre_daily_delivery_date, o.imp_file_id, " &
                "  o.order_type, o.prorated_type, o.customer_info_type, o.info_type, o.self_fcst_flag, o.self_fcst_delete_flag, " &
                "  o.reconcile_type, o.imp_run_id, o.status, o.active_flag, " &
                "  o.created_at, o.created_user_id, o.created_pg_id, " &
                "  o.updated_at, o.updated_user_id, o.updated_pg_id, " &
                "  o.stra_order_qty, o.stra_ship_qty, o.stra_order_backlog " &
                "  FROM orders o " &
                "  WHERE o.customer_setting_id = :p_customer_setting_id " &
                "  AND o.order_type IN (1,2) " &
                "  AND o.active_flag = 'Y' " &
                "  AND NOT EXISTS ( " &
                "    SELECT 1 FROM orders_stage os " &
                "    WHERE os.order_id = o.order_id " &
                "  ) "

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = CustomerSettingId

                'cmd.ExecuteNonQuery()
                ' 実行
                Dim updatedCount As Integer = cmd.ExecuteNonQuery()

            End Using

        End Sub

        ''' <summary>
        ''' 今回取込した内示データ取得
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="ImpFileStageId">処理中の取込ワークファイルID</param>
        Public Function GetNaijiData(ByVal tran As OracleTransaction, ByVal ImpFileStageId As Long) As DataTable

            Dim dt As New DataTable()

            Const sql As String =
                        "  SELECT * FROM orders_stage " &
                        "  WHERE imp_file_stage_id = :p_imp_file_stage_id " &
                        "  AND order_type = 1 " &
                        "  AND active_flag = 'Y' " &
                        "  ORDER BY due_date ASC "

            Using cmd As New OracleCommand(sql, tran.Connection)

                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = ImpFileStageId

                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' 今回取込（impFileStageId）を起点に、同一取引先の過去の内示データをすべて洗い替え（無効化）する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="impFileStageId">処理対象の一時取込ファイルID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub ReplaceNaijiRelation(ByVal tran As OracleTransaction,
                                        ByVal impFileStageId As Long,
                                        ByVal customerSettingId As Long,
                                        ByVal updatedAt As DateTime,
                                        ByVal updateUserId As String,
                                        ByVal updatepgId As String)

            Dim NCount As Integer = 0
            Dim YCount As Integer = 0

            Dim sqlCount As String = "
                    SELECT COUNT(*) 
                    FROM orders_stage 
                    WHERE imp_file_stage_id = :p_imp_file_stage_id 
                    AND order_type = 1 
                    AND self_fcst_flag = 'N' 
                    AND active_flag = 'Y'"

            Using cmd As New OracleCommand(sqlCount, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータ設定
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId

                '実行 (取引先内示件数を取得)
                NCount = Convert.ToInt32(cmd.ExecuteScalar())
            End Using

            sqlCount = "
                    SELECT COUNT(*) 
                    FROM orders_stage 
                    WHERE imp_file_stage_id = :p_imp_file_stage_id 
                    AND order_type = 1 
                    AND self_fcst_flag = 'Y' 
                    AND active_flag = 'Y'"

            Using cmd As New OracleCommand(sqlCount, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータ設定
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId

                '実行 (ASTI追加内示件数を取得)
                YCount = Convert.ToInt32(cmd.ExecuteScalar())
            End Using


            Dim selfFcstFlag As String = Nothing
            Dim selfFcstDeleteFlag As String = Nothing

            If NCount >= 1 AndAlso YCount = 0 Then

                '取引先内示の削除処理
                selfFcstFlag = "N"
                UpdateNaijiStatusReplaced(tran, impFileStageId, customerSettingId, updatedAt, updateUserId, updatepgId, selfFcstFlag, selfFcstDeleteFlag)

                'ASTI追加内示の自動削除処理
                selfFcstFlag = "Y"
                selfFcstDeleteFlag = "Y"
                UpdateNaijiStatusReplaced(tran, impFileStageId, customerSettingId, updatedAt, updateUserId, updatepgId, selfFcstFlag, selfFcstDeleteFlag)

            ElseIf NCount = 0 AndAlso YCount >= 1 Then

                'ASTI追加内示の削除処理
                selfFcstFlag = "Y"
                UpdateNaijiStatusReplaced(tran, impFileStageId, customerSettingId, updatedAt, updateUserId, updatepgId, selfFcstFlag, selfFcstDeleteFlag)

            ElseIf NCount >= 1 AndAlso YCount >= 1 Then

                '全件(取引先内示とASTI追加内示)の削除処理
                UpdateNaijiStatusReplaced(tran, impFileStageId, customerSettingId, updatedAt, updateUserId, updatepgId, selfFcstFlag, selfFcstDeleteFlag)

            End If

        End Sub

        ''' <summary>
        ''' 一致するデータの STATUS を 'REPLACED' に更新する
        ''' </summary>
        '''  ''' <param name="tran">トランザクション</param>
        ''' <param name="impFileStageId">処理対象の一時取込ファイルID</param>
        Public Sub UpdateNaijiStatusReplaced(ByVal tran As OracleTransaction,
                                        ByVal impFileStageId As Long,
                                        ByVal customerSettingId As Long,
                                        ByVal updatedAt As DateTime,
                                        ByVal updateUserId As String,
                                        ByVal updatepgId As String,
                                        Optional ByVal selfFcstFlag As String = Nothing,
                                        Optional ByVal selfFcstDeleteFlag As String = Nothing)



            '2026/05/22 st 酒井
            ' 一致するデータの STATUS を 'REPLACED' に更新
            'Dim sql As String =
            '            " UPDATE orders_stage SET" &
            '            " status = 'REPLACED', " &
            '            " active_flag = 'N', " &
            '            " updated_at = :p_updated_at, " &
            '            " updated_user_id = :p_user_id, " &
            '            " updated_pg_id = :p_updated_pg_id " &
            '            " WHERE 1=1 " &
            '            " AND order_type = 1 " &
            '            " AND active_flag = 'Y' " &
            '            " AND customer_setting_id = :p_customer_setting_id " &
            '            " AND imp_file_stage_id <> :p_imp_file_stage_id "
            Dim sql As String =
                        " UPDATE orders_stage SET" &
                        " status = 'REPLACED', " &
                        " active_flag = 'N', " &
                        " updated_at = :p_updated_at, " &
                        " updated_user_id = :p_user_id, " &
                        " updated_pg_id = :p_updated_pg_id " &
                        " WHERE 1=1 " &
                        " AND order_type = 1 " &
                        " AND active_flag = 'Y' " &
                        " AND customer_setting_id = :p_customer_setting_id " &
                        " AND imp_file_stage_id IS NULL "
            '2026/05/22 ed 酒井

            If selfFcstFlag IsNot Nothing Then
                sql &= " AND self_fcst_flag = :p_self_fcst_flag "
            End If

            If selfFcstDeleteFlag IsNot Nothing Then
                sql &= " AND self_fcst_delete_flag = :p_self_fcst_delete_flag "
            End If

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                If selfFcstFlag IsNot Nothing Then
                    cmd.Parameters.Add(New OracleParameter(":p_self_fcst_flag", OracleDbType.Varchar2) With {.Value = selfFcstFlag})
                End If

                If selfFcstDeleteFlag IsNot Nothing Then
                    cmd.Parameters.Add(New OracleParameter(":p_self_fcst_delete_flag", OracleDbType.Varchar2) With {.Value = selfFcstDeleteFlag})
                End If

                cmd.ExecuteNonQuery()
            End Using

        End Sub

        ''' <summary>
        ''' 今回取り込んだ内示データのステータスを 'PROCESSED' に更新する
        ''' </summary>
        '''  ''' <param name="tran">トランザクション</param>
        ''' <param name="impFileStageId">処理対象の一時取込ファイルID</param>
        Public Sub UpdateNaijiStatusProcessed(ByVal tran As OracleTransaction, ByVal impFileStageId As Long)

            ' 今回のimp_file_stage_idに一致するデータの STATUS を 'PROCESSED' に更新
            Const sql As String =
                        " UPDATE orders_stage " &
                        " SET status = 'PROCESSED' " &
                        " WHERE 1=1 " &
                        " AND imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND order_type = 1"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId

                cmd.ExecuteNonQuery()
            End Using

        End Sub

        ''' <summary>
        ''' 打切処理
        ''' </summary>
        ''' <param name="tran">実行中のトランザクション</param>
        ''' <param name="impFileStageId">処理中の一時取込ファイルID</param>
        ''' <param name="OrderType">処理中の受注区分</param>
        Public Sub UpdateClese(ByVal tran As OracleTransaction, ByVal ImpFileStageId As Long, ByVal OrderType As Integer)

            Const sql As String =
                        "UPDATE orders_stage SET " &
                        "  status = 'PROCESSED', " &
                        "  active_flag = 'N' " &
                        "  WHERE 1=1 " &
                        "  AND imp_file_stage_id = :p_imp_file_stage_id " &
                        "  AND order_type = :p_order_type " &
                        "  AND info_type = 'N' "

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = ImpFileStageId
                cmd.Parameters.Add(":p_order_type", OracleDbType.Int16).Value = OrderType

                cmd.ExecuteNonQuery()

            End Using

        End Sub

        ''' <summary>
        ''' 取消処理
        ''' </summary>
        ''' <param name="tran">実行中のトランザクション</param>
        ''' <param name="impFileStageId">処理中の一時取込ファイルID</param>
        ''' <param name="OrderType">処理中の受注区分</param>
        Public Sub UpdateCancel(ByVal tran As OracleTransaction, ByVal ImpFileStageId As Long, ByVal OrderType As Integer)

            Const sql As String =
                        "UPDATE orders_stage SET " &
                        "  status = 'PROCESSED', " &
                        "  demand_qty = 0 " &
                        "  WHERE 1=1 " &
                        "  AND imp_file_stage_id = :p_imp_file_stage_id " &
                        "  AND order_type = :p_order_type " &
                        "  AND info_type = 'D' "

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = ImpFileStageId
                cmd.Parameters.Add(":p_order_type", OracleDbType.Int16).Value = OrderType

                cmd.ExecuteNonQuery()

            End Using

        End Sub

        ''' <summary>
        ''' 確定データを無効化して、今回の取込データを更新する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        ''' <param name="impFileStageId">処理中の取込ワークファイルID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub ReplaceKakuteiRelation(ByVal tran As OracleTransaction,
                                        ByVal customerSettingId As Long,
                                        ByVal impFileStageId As Long,
                                        ByVal updatedAt As DateTime,
                                        ByVal updateUserId As String,
                                        ByVal updatepgId As String)

            ' 過去データ（今回取込データ以外）で、キーが一致するものを無効化
            Dim sql As String =
                        " UPDATE orders_stage tgt" &
                        " SET " &
                        " tgt.status = 'REPLACED', " &
                        " tgt.active_flag = 'N', " &
                        " tgt.updated_at = :p_updated_at, " &
                        " tgt.updated_user_id = :p_user_id, " &
                        " tgt.updated_pg_id = :p_updated_pg_id " &
                        " WHERE tgt.order_type = 2 " &
                        " AND tgt.active_flag = 'Y' " &
                        " AND tgt.imp_file_id <> :p_imp_file_stage_id " &
                        " AND EXISTS ( " &
                        " SELECT 1 " &
                        " FROM orders_stage cur " &
                        " WHERE cur.imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND cur.customer_setting_id = :p_customer_setting_id " &
                        " AND cur.order_type = 2 " &
                        " AND cur.active_flag = 'Y' " &
                        " AND NVL(cur.info_type, 'U') = 'U' " &
                        " AND cur.customer_setting_id = tgt.customer_setting_id " &
                        " AND cur.customer_order_no = tgt.customer_order_no) "

            Using cmd As New OracleCommand(sql, tran.Connection)

                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                cmd.ExecuteNonQuery()

            End Using

            ' 今回の取込データを更新
            sql =
                        " UPDATE orders_stage tgt" &
                        " SET " &
                        " tgt.info_type = 'U', " &
                        " tgt.status = 'PROCESSED', " &
                        " tgt.updated_at = :p_updated_at, " &
                        " tgt.updated_user_id = :p_user_id, " &
                        " tgt.updated_pg_id = :p_updated_pg_id " &
                        " WHERE tgt.order_type = 2 " &
                        " AND tgt.active_flag = 'Y' " &
                        " AND NVL(tgt.info_type,'U') = 'U' " &
                        " AND tgt.imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND EXISTS ( " &
                        " SELECT 1 " &
                        " FROM orders_stage cur " &
                        " WHERE cur.imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND cur.customer_setting_id = :p_customer_setting_id " &
                        " AND cur.order_type = 2 " &
                        " AND cur.active_flag = 'Y' " &
                        " AND NVL(cur.info_type, 'U') = 'U' " &
                        " AND cur.customer_setting_id = tgt.customer_setting_id " &
                        " AND cur.customer_order_no = tgt.customer_order_no) "

            Using cmd As New OracleCommand(sql, tran.Connection)

                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                cmd.ExecuteNonQuery()

            End Using

        End Sub

        ''' <summary>
        ''' 今回取込した確定データ取得
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="ImpFileStageId">処理中の取込ワークファイルID</param>
        Public Function GetKakuteiData(ByVal tran As OracleTransaction, ByVal ImpFileStageId As Long) As DataTable

            Dim dt As New DataTable()

            Const sql As String =
                        "  SELECT * FROM orders_stage " &
                        "  WHERE imp_file_stage_id = :p_imp_file_stage_id " &
                        "  AND order_type = 2 " &
                        "  AND active_flag = 'Y' " &
                        "  AND (info_type IS NULL OR info_type = 'I') " &
                        "  ORDER BY imp_file_stage_id ASC, item_no ASC, due_date ASC "

            Using cmd As New OracleCommand(sql, tran.Connection)

                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = ImpFileStageId

                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' 内示消込処理を実行する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        ''' <param name="impFileStageId">処理中の取込ワークファイルID</param>
        ''' <param name="orderType">受注区分（2:確定 or 3:納入指示）</param>
        ''' <param name="reconcileType">消込条件（1:順次, 2:同月まで, 3:同月のみ）</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub ReconcileForecast(ByVal tran As OracleTransaction,
                                        ByVal customerSettingId As Long,
                                        ByVal impFileStageId As Long,
                                        ByVal orderType As Integer,
                                        ByVal reconcileType As Integer,
                                        ByVal updatedAt As DateTime,
                                        ByVal updateUserId As String,
                                        ByVal updatepgId As String)


            ' --- 1. 今回取込データの集計（品目ごとの需要数の合計を取得、納期は最も古いものを採用） ---
            Dim curList As New Dictionary(Of String, OrderSummaryRow)

            Dim curSelect As String = ""
            Dim curWhere As String = ""
            Dim curGroupBy As String = ""
            Select Case reconcileType
                Case 1
                    curSelect = " MIN(due_date) AS earliest_due_date,"
                    'curWhere = ""
                    If orderType = 2 Then
                        curWhere = " AND (info_type IS NULL OR info_type = 'I')"
                    End If
                    curGroupBy = ""
                Case 2
                    curSelect = " MAX(due_date) AS earliest_due_date,"
                    'curWhere = " AND (info_type IS NULL OR info_type = 'I')"
                    If orderType = 2 Then
                        curWhere = " AND (info_type IS NULL OR info_type = 'I')"
                    End If
                    curGroupBy = ""
                Case 3
                    curSelect = " TRUNC(due_date, 'MM') AS earliest_due_date,"
                    'curWhere = ""
                    If orderType = 2 Then
                        curWhere = " AND (info_type IS NULL OR info_type = 'I')"
                    End If
                    curGroupBy = " ,TRUNC(due_date, 'MM')"
                Case Else
                    curSelect = " MIN(due_date) AS earliest_due_date,"
                    'curWhere = ""
                    If orderType = 2 Then
                        curWhere = " AND (info_type IS NULL OR info_type = 'I')"
                    End If
                    curGroupBy = ""
            End Select


            Dim cursql As String = $"
                SELECT
                    item_no,
                    {curSelect}
                    SUM(demand_qty) AS total_demand_qty
                FROM orders_stage
                WHERE imp_file_stage_id = :p_imp_file_stage_id
                    AND order_type = :p_order_type
                    AND status = 'IMPORTED'
                    AND active_flag = 'Y'
                    {curWhere}
                GROUP BY order_type,item_no
                    {curGroupBy}
                ORDER BY item_no ASC"

            Using cmdCur As New OracleCommand(cursql, tran.Connection)
                cmdCur.Transaction = tran
                cmdCur.BindByName = True
                cmdCur.CommandType = CommandType.Text
                cmdCur.Parameters.Clear()

                cmdCur.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId
                cmdCur.Parameters.Add(":p_order_type", OracleDbType.Int32).Value = orderType

                Using dr As OracleDataReader = cmdCur.ExecuteReader()
                    While dr.Read()

                        Dim itemNo As String = dr("item_no").ToString()
                        Dim dueDate As DateTime = Convert.ToDateTime(dr("earliest_due_date"))

                        'Dim row As New OrderSummaryRow With {
                        '    .ItemNo = dr("item_no").ToString(),
                        '    .EarliestDueDate = Convert.ToDateTime(dr("earliest_due_date")),
                        '    .TotalDemandQty = Convert.ToDecimal(dr("total_demand_qty"))
                        '}
                        'curList.Add(row.ItemNo, row)

                        Dim row As New OrderSummaryRow With {
                            .ItemNo = itemNo,
                            .EarliestDueDate = dueDate,
                            .TotalDemandQty = Convert.ToDecimal(dr("total_demand_qty"))
                        }
                        Dim dictKey As String = itemNo
                        If reconcileType = 3 Then
                            ' 品目No_202310 のような形式で月ごとに枠を管理する
                            dictKey = $"{itemNo}_{dueDate:yyyyMM}"
                        End If

                        curList.Add(dictKey, row)

                    End While
                End Using
            End Using

            ' --- 2. 今回取込データの内示データを古い順に取得
            Dim pastSql As String = $"
                    SELECT p.rowid, p.item_no, p.demand_qty, p.due_date
                    FROM orders_stage p
                    WHERE p.order_type = 1 
                      AND p.active_flag = 'Y' 
                      AND p.self_fcst_flag = 'N' 
                    ORDER BY p.item_no, p.due_date, p.rowid"

            Using cmdPast As New OracleCommand(pastSql, tran.Connection)
                cmdPast.Transaction = tran
                cmdPast.BindByName = True
                cmdPast.CommandType = CommandType.Text
                cmdPast.Parameters.Clear()

                Using drPast As OracleDataReader = cmdPast.ExecuteReader()
                    While drPast.Read()
                        Dim itemNo As String = drPast("item_no").ToString()
                        Dim forecastDate As DateTime = Convert.ToDateTime(drPast("due_date"))

                        Dim dictKey As String = itemNo
                        If reconcileType = 3 Then
                            dictKey = $"{itemNo}_{forecastDate:yyyyMM}"
                        End If

                        ' Dictionaryに該当品目（かつ該当月）の消込枠があるか確認
                        If curList.ContainsKey(dictKey) Then
                            Dim summary = curList(dictKey)

                            ' --- 消込対象かどうかの判定 (reconcileTypeによる比較) ---
                            Dim isTarget As Boolean = False
                            ' 年月のみで比較するための変数作成
                            Dim fcstMonth As DateTime = New DateTime(forecastDate.Year, forecastDate.Month, 1)
                            Dim curMonth As DateTime = New DateTime(summary.EarliestDueDate.Year, summary.EarliestDueDate.Month, 1)

                            Select Case reconcileType
                                Case 1 : isTarget = True ' 1:順次（全対象）
                                Case 2 : isTarget = (fcstMonth <= curMonth) ' 2:同月まで
                                Case 3 : isTarget = (fcstMonth = curMonth)  ' 3:同月のみ
                            End Select

                            ' 対象であり、かつまだ消込枠(注文合計)が残っている場合のみ処理
                            If isTarget AndAlso summary.TotalDemandQty > 0 Then
                                Dim rid As String = drPast("rowid").ToString()
                                Dim pastQty As Decimal = Convert.ToDecimal(drPast("demand_qty"))
                                Dim remainLimit As Decimal = summary.TotalDemandQty ' 現在の消込可能残数

                                Dim newDemandQty As Decimal = 0

                                ' --- 消込計算ロジック ---
                                If pastQty > remainLimit Then
                                    ' 内示残数の方が大きい場合：内示を一部減らし、消込枠を使い切る
                                    newDemandQty = pastQty - remainLimit
                                    summary.TotalDemandQty = 0
                                Else
                                    ' 消込枠の方が多い場合：この内示レコードを0にし、残った枠を次の納期分へ
                                    newDemandQty = 0
                                    summary.TotalDemandQty -= pastQty
                                End If

                                ' --- データベース更新 (rowid指定でピンポイント更新) ---
                                UpdateOrderRow(tran, rid, newDemandQty, updatedAt, updateUserId, updatepgId)
                            End If
                        End If
                    End While
                End Using

            End Using

        End Sub

        ''' <summary>
        ''' レコード更新用(rowid指定でピンポイント更新)
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="rowid">更新対象の行ID</param>
        ''' <param name="qty">需要数</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Private Sub UpdateOrderRow(ByVal tran As OracleTransaction,
                                   ByVal rowid As String,
                                   ByVal qty As Decimal,
                                   ByVal updatedAt As DateTime,
                                   ByVal updateUserId As String,
                                   ByVal updatepgId As String)

            Dim sql As String = "
                    UPDATE orders_stage SET 
                        demand_qty = :p_qty,
                        status = CASE WHEN :p_qty <= 0 THEN 'RECONCILED' ELSE status END,
                        active_flag = CASE WHEN :p_qty <= 0 THEN 'N' ELSE active_flag END,
                        updated_at = :p_updated_at,
                        updated_user_id = :p_user_id,
                        updated_pg_id = :p_updated_pg_id
                    WHERE rowid = :p_rid"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_qty", OracleDbType.Decimal).Value = qty
                cmd.Parameters.Add(":p_rid", OracleDbType.Varchar2).Value = rowid
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 今回取込したレコードのステータス更新(確定 新規)を実行する
        ''' </summary>
        ''' <param name="tran">実行中のトランザクション</param>
        ''' <param name="impFileStageId">処理中の一時取込ファイルID</param>
        Public Sub UpdateKakuteiNewOrders(ByVal tran As OracleTransaction, ByVal impFileStageId As Long)
            ' 仕様書 P5「確定新規」の条件に基づくUPDATE文
            Const sql As String =
                        " UPDATE orders_stage " &
                        " SET status = 'PROCESSED' " &
                        " ,info_type = 'I' " &
                        " WHERE imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND order_type = 2 " &
                        " AND active_flag = 'Y'" &
                        " AND (info_type = 'I' OR info_type IS NULL)"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータ設定
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                ' 更新実行
                'Dim count As Integer = cmd.ExecuteNonQuery()
                cmd.ExecuteNonQuery()

            End Using
        End Sub

        ''' <summary>
        ''' 確定消込（客先発注Noでの消込）を実行する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        ''' <param name="impFileStageId">今回取込の IMP_FILE_STAGE_ID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub ExecuteOrderReconciliationByOrderNo(ByVal tran As OracleTransaction,
                                                        ByVal customerSettingId As Long,
                                                        ByVal impFileStageId As Long,
                                                        ByVal updatedAt As DateTime,
                                                        ByVal updateUserId As String,
                                                        ByVal updatepgId As String)

            '' 1. 同一の CUSTOMER_SETTING_ID と CUSTOMER_ORDER_NO を持つ過去データを特定
            '' 2. DEMAND_QTY を (過去 - 今回) で減数更新
            '' 3. 需要数が0になった場合は STATUS='RECONCILED', ACTIVE_FLAG='N' に更新

            Dim Sql As String =
                    " MERGE INTO orders_stage tgt " &
                    " USING ( " &
                    " WITH cur AS ( " &
                    " SELECT " &
                    " c.customer_setting_id, " &
                    " c.customer_order_no, " &
                    " SUM(c.demand_qty) AS demand_qty " &
                    " FROM orders_stage c " &
                    " WHERE c.imp_file_stage_id = :p_imp_file_stage_id " &
                    " AND c.order_type = 3 " &
                    " AND NVL(c.info_type, 'U') = 'U' " &
                    " And c.active_flag = 'Y' " &
                    " GROUP BY c.customer_setting_id, c.customer_order_no " &
                    " ), " &
                    " past AS ( " &
                    " SELECT p.rowid AS rid, p.customer_setting_id, p.customer_order_no, p.demand_qty " &
                    " FROM orders_stage p " &
                    " WHERE p.imp_file_id <> :p_imp_file_stage_id " &
                    " AND p.order_type = 2 " &
                    " AND p.active_flag = 'Y' " &
                    " AND p.customer_setting_id = :p_customer_setting_id " &
                    " ) " &
                    " SELECT  " &
                    " p.rid, " &
                    " p.demand_qty AS past_qty, " &
                    " c.demand_qty AS cur_qty " &
                    " FROM past p " &
                    " JOIN cur c ON c.customer_setting_id = p.customer_setting_id " &
                    " AND c.customer_order_no = p.customer_order_no " &
                    " ) m " &
                    " ON (tgt.rowid = m.rid) " &
                    " WHEN MATCHED THEN " &
                    " UPDATE SET " &
                    " tgt.demand_qty      = NVL(m.past_qty, 0) - NVL(m.cur_qty, 0), " &
                    " tgt.status          = CASE WHEN NVL(m.past_qty, 0) - NVL(m.cur_qty, 0) = 0 THEN 'RECONCILED' ELSE tgt.status END, " &
                    " tgt.active_flag     = CASE WHEN NVL(m.past_qty, 0) - NVL(m.cur_qty, 0) = 0 THEN 'N' ELSE tgt.active_flag END, " &
                    " tgt.updated_at      = :p_updated_at, " &
                    " tgt.updated_user_id = :p_user_id, " &
                    " tgt.updated_pg_id   = :p_updated_pg_id "

            Using cmd As New OracleCommand(Sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータ設定
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)


                cmd.ExecuteNonQuery()

            End Using
        End Sub

        ''' <summary>
        ''' 今回取込した納入指示データ取得
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="ImpFileStageId">処理中の取込ワークファイルID</param>
        Public Function GetNonyuSijiData(ByVal tran As OracleTransaction, ByVal ImpFileStageId As Long) As DataTable

            Dim dt As New DataTable()

            Const sql As String =
                        "  SELECT * FROM orders_stage " &
                        "  WHERE imp_file_stage_id = :p_imp_file_stage_id " &
                        "  AND order_type = 3 " &
                        "  AND active_flag = 'Y' " &
                        "  ORDER BY due_date ASC "

            Using cmd As New OracleCommand(sql, tran.Connection)

                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = ImpFileStageId

                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' 今回取込したレコードのステータス更新(納入指示 新規)を実行する
        ''' </summary>
        ''' <param name="tran">実行中のトランザクション</param>
        ''' <param name="impFileStageId">処理中の一時取込ファイルID</param>
        Public Sub UpdateNonyuSijiNewOrders(ByVal tran As OracleTransaction, ByVal impFileStageId As Long)
            ' 仕様書 P5「確定新規」の条件に基づくUPDATE文
            Const sql As String =
                        " UPDATE orders_stage " &
                        " SET status = 'PROCESSED' " &
                        " WHERE imp_file_stage_id = :p_imp_file_stage_id " &
                        " AND order_type = 3 " &
                        " AND (info_type = 'I' OR info_type IS NULL) " &
                        " AND active_flag = 'Y'"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータ設定
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Int64).Value = impFileStageId
                ' 更新実行
                Dim count As Integer = cmd.ExecuteNonQuery()

                ' ログ出力（必要に応じて）
                Console.WriteLine($"{count} 件の確定新規レコードを更新しました。")
            End Using
        End Sub

        '2026/05/26 酒井 フェーズ2 受注残対応
        ''' <summary>
        ''' 受注残消込処理を実行する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="customerSettingId">処理中の取引先設定ID</param>
        ''' <param name="impFileStageId">処理中の取込ワークファイルID</param>
        ''' <param name="reconcileType">消込条件（1:順次, 2:同月まで, 3:同月のみ）</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub BacklogForecast(ByVal tran As OracleTransaction,
                                        ByVal customerSettingId As Long,
                                        ByVal impFileStageId As Long,
                                        ByVal reconcileType As Integer,
                                        ByVal updatedAt As DateTime,
                                        ByVal updateUserId As String,
                                        ByVal updatepgId As String)


            ' --- 1. 今回取込データの集計（品目ごとの需要数の合計を取得、納期は最も古いものを採用） ---
            Dim curList As New Dictionary(Of String, OrderSummaryRow)

            Dim curSelect As String = ""
            Dim curWhere As String = ""
            Dim curGroupBy As String = ""
            Select Case reconcileType
                Case 1
                    curSelect = " MIN(pre_daily_delivery_date) AS earliest_pre_daily_delivery_date,"
                    curWhere = ""
                    curGroupBy = ""
                Case 2
                    curSelect = " MAX(pre_daily_delivery_date) AS earliest_pre_daily_delivery_date,"
                    curWhere = ""
                    curGroupBy = ""
                Case 3
                    curSelect = " TRUNC(pre_daily_delivery_date, 'MM') AS earliest_pre_daily_delivery_date,"
                    curWhere = ""
                    curGroupBy = " ,TRUNC(pre_daily_delivery_date, 'MM')"
                Case Else
                    curSelect = " MIN(pre_daily_delivery_date) AS earliest_pre_daily_delivery_date,"
                    curWhere = ""
                    curGroupBy = ""
            End Select

            Dim cursql As String = $"
                SELECT
                    item_no,
                    {curSelect}
                    SUM(stra_order_backlog) AS total_stra_order_backlog
                FROM order_backlog_view
                WHERE 1=1
                    {curWhere}
                GROUP BY item_no
                    {curGroupBy}
                ORDER BY item_no ASC"

            Using cmdCur As New OracleCommand(cursql, tran.Connection)
                cmdCur.Transaction = tran
                cmdCur.BindByName = True
                cmdCur.CommandType = CommandType.Text
                cmdCur.Parameters.Clear()

                Using dr As OracleDataReader = cmdCur.ExecuteReader()
                    While dr.Read()

                        Dim itemNo As String = dr("item_no").ToString()
                        Dim dueDate As DateTime = Convert.ToDateTime(dr("earliest_pre_daily_delivery_date"))

                        Dim row As New OrderSummaryRow With {
                            .ItemNo = itemNo,
                            .EarliestDueDate = dueDate,
                            .TotalDemandQty = Convert.ToDecimal(dr("total_stra_order_backlog"))
                        }
                        Dim dictKey As String = itemNo
                        If reconcileType = 3 Then
                            ' 品目No_202310 のような形式で月ごとに枠を管理する
                            dictKey = $"{itemNo}_{dueDate:yyyyMM}"
                        End If

                        curList.Add(dictKey, row)

                    End While
                End Using
            End Using

            ' --- 2. 今回取込データの内示データを古い順に取得
            Dim pastSql As String = $"
                    SELECT p.rowid, p.item_no, p.demand_qty, p.due_date
                    FROM orders_stage p
                    WHERE p.order_type = 1 
                      AND p.active_flag = 'Y' 
                      AND p.self_fcst_flag = 'N' 
                    ORDER BY p.item_no, p.due_date, p.rowid"

            Using cmdPast As New OracleCommand(pastSql, tran.Connection)
                cmdPast.Transaction = tran
                cmdPast.BindByName = True
                cmdPast.CommandType = CommandType.Text
                cmdPast.Parameters.Clear()

                Using drPast As OracleDataReader = cmdPast.ExecuteReader()
                    While drPast.Read()
                        Dim itemNo As String = drPast("item_no").ToString()
                        Dim forecastDate As DateTime = Convert.ToDateTime(drPast("due_date"))

                        Dim dictKey As String = itemNo
                        If reconcileType = 3 Then
                            dictKey = $"{itemNo}_{forecastDate:yyyyMM}"
                        End If

                        ' Dictionaryに該当品目（かつ該当月）の消込枠があるか確認
                        If curList.ContainsKey(dictKey) Then
                            Dim summary = curList(dictKey)

                            ' --- 消込対象かどうかの判定 (reconcileTypeによる比較) ---
                            Dim isTarget As Boolean = False
                            ' 年月のみで比較するための変数作成
                            Dim fcstMonth As DateTime = New DateTime(forecastDate.Year, forecastDate.Month, 1)
                            Dim curMonth As DateTime = New DateTime(summary.EarliestDueDate.Year, summary.EarliestDueDate.Month, 1)

                            Select Case reconcileType
                                Case 1 : isTarget = True ' 1:順次（全対象）
                                Case 2 : isTarget = (fcstMonth <= curMonth) ' 2:同月まで
                                Case 3 : isTarget = (fcstMonth = curMonth)  ' 3:同月のみ
                            End Select

                            ' 対象であり、かつまだ消込枠(注文合計)が残っている場合のみ処理
                            If isTarget AndAlso summary.TotalDemandQty > 0 Then
                                Dim rid As String = drPast("rowid").ToString()
                                Dim pastQty As Decimal = Convert.ToDecimal(drPast("demand_qty"))
                                Dim remainLimit As Decimal = summary.TotalDemandQty ' 現在の消込可能残数

                                Dim newDemandQty As Decimal = 0

                                ' --- 消込計算ロジック ---
                                If pastQty > remainLimit Then
                                    ' 内示残数の方が大きい場合：内示を一部減らし、消込枠を使い切る
                                    newDemandQty = pastQty - remainLimit
                                    summary.TotalDemandQty = 0
                                Else
                                    ' 消込枠の方が多い場合：この内示レコードを0にし、残った枠を次の納期分へ
                                    newDemandQty = 0
                                    summary.TotalDemandQty -= pastQty
                                End If

                                ' --- データベース更新 (rowid指定でピンポイント更新) ---
                                UpdateOrderRow(tran, rid, newDemandQty, updatedAt, updateUserId, updatepgId)
                            End If
                        End If
                    End While
                End Using

            End Using

        End Sub
        '--

        ''' <summary>
        ''' ログインユーザーが担当する加工済みデータの取込先設定ID(customer_setting_id)の一覧とデータ件数を取得する
        ''' </summary>
        ''' <param name="LoginUserId">ログインユーザーID</param>
        Public Function GetProcessedCustomerSettingIds(ByVal LoginUserId As String) As DataTable

            'sql =
            '             " SELECT os.customer_setting_id " &
            '             " ,os.customer_code " &
            '             " ,COUNT(*) as cnt " &
            '             " FROM orders_stage os " &
            '             " WHERE os.status = 'PROCESSED' " &
            '             " AND EXISTS ( " &
            '             "   SELECT 1 " &
            '             "   FROM customer_setting_mst csm " &
            '             "   WHERE csm.customer_setting_id = os.customer_setting_id " &
            '             "     AND UPPER(csm.prod_mgmt_user_id) = UPPER(:p_login_user_id) " &
            '             " ) " &
            '             " GROUP BY os.customer_setting_id, os.customer_code " &
            '             " ORDER BY os.customer_setting_id ASC "

            Dim dt As New DataTable()

            Dim sql As String = "SELECT os.customer_setting_id " &
                                 " ,os.customer_code " &
                                 " ,COUNT(*) As cnt " &
                                 " FROM orders_stage os " &
                                 " WHERE os.status = 'PROCESSED' "

            Dim isAdmin As Boolean = String.Equals(LoginUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)

            If Not isAdmin Then
                sql &= " AND EXISTS ( " &
                         "   SELECT 1 " &
                         "   FROM customer_setting_mst csm " &
                         "   WHERE csm.customer_setting_id = os.customer_setting_id " &
                         "     AND UPPER(csm.prod_mgmt_user_id) = UPPER(:p_login_user_id) " &
                         " ) " &
                         " GROUP BY os.customer_setting_id, os.customer_code " &
                         " ORDER BY os.customer_setting_id ASC "
            Else
                sql &= " GROUP BY os.customer_setting_id, os.customer_code " &
                       " ORDER BY os.customer_setting_id ASC "
            End If

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Dim pLoginUserId As String = If(String.IsNullOrWhiteSpace(LoginUserId), Nothing, LoginUserId.Trim())

                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Clear()

                    If Not isAdmin Then
                        ' パラメータ設定
                        cmd.Parameters.Add(":p_login_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(pLoginUserId, 9)
                    End If

                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using

                End Using

            End Using

            Return dt

        End Function

        ''' <summary>
        ''' ログインユーザーが担当する加工済みデータの取込先設定ID(customer_setting_id)単位の取込ファイルの一覧を取得する
        ''' </summary>
        ''' <param name="LoginUserId">ログインユーザーID</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        Public Function GetProcessedImpFileStageIds(ByVal LoginUserId As String, ByVal customerSettingId As Long) As DataTable

            Dim dt As New DataTable()

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Dim pLoginUserId As String = If(String.IsNullOrWhiteSpace(LoginUserId), Nothing, LoginUserId.Trim())

                'Dim sql As String =
                ' " SELECT os.imp_file_stage_id " &
                ' " FROM orders_stage os " &
                ' " WHERE os.status = 'PROCESSED' " &
                ' " AND os.imp_file_stage_id IS NOT NULL " &
                ' " AND EXISTS ( " &
                ' "   SELECT 1 " &
                ' "   FROM customer_setting_mst csm " &
                ' "   WHERE csm.customer_setting_id = os.customer_setting_id " &
                ' "     AND UPPER(csm.prod_mgmt_user_id) = UPPER(:p_login_user_id) " &
                ' "     AND os.customer_setting_id = :p_customer_setting_id " &
                ' " ) " &
                ' " GROUP BY os.imp_file_stage_id " &
                ' " ORDER BY os.imp_file_stage_id ASC "

                Dim sql As String = " SELECT os.imp_file_stage_id " &
                                    " FROM orders_stage os " &
                                    " WHERE os.status = 'PROCESSED' " &
                                    " AND os.imp_file_stage_id IS NOT NULL "

                Dim isAdmin As Boolean = String.Equals(LoginUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)

                If Not isAdmin Then
                    sql &= " AND EXISTS ( " &
                           "   SELECT 1 " &
                           "   FROM customer_setting_mst csm " &
                           "   WHERE csm.customer_setting_id = os.customer_setting_id " &
                           "     AND UPPER(csm.prod_mgmt_user_id) = UPPER(:p_login_user_id) " &
                           "     AND os.customer_setting_id = :p_customer_setting_id " &
                           " ) " &
                           " GROUP BY os.imp_file_stage_id " &
                           " ORDER BY os.imp_file_stage_id ASC "
                Else
                    sql &= " AND EXISTS ( " &
                           "   SELECT 1 " &
                           "   FROM customer_setting_mst csm " &
                           "   WHERE csm.customer_setting_id = os.customer_setting_id " &
                           "     AND os.customer_setting_id = :p_customer_setting_id " &
                           " ) " &
                           " GROUP BY os.imp_file_stage_id " &
                           " ORDER BY os.imp_file_stage_id ASC "
                End If

                Using cmd As New OracleCommand(sql, conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Clear()

                    If Not isAdmin Then
                        ' パラメータ設定
                        cmd.Parameters.Add(":p_login_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(pLoginUserId, 9)
                    End If

                    ' パラメータ設定
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId

                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using

                End Using

            End Using

            Return dt

        End Function

        ''' <summary>
        ''' ログインユーザーが担当する加工済みの受注ワークデータを削除する
        ''' </summary>
        ''' <param name="LoginUserId">ログインユーザーID</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        Public Sub DeleteProcessedOrdersRange(ByVal LoginUserId As String, ByVal customerSettingId As Long)

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    DeleteProcessedOrdersRange(tran, LoginUserId, customerSettingId)
                    tran.Commit()
                End Using
            End Using

        End Sub

        ''' <summary>
        ''' ログインユーザーが担当する加工済みの受注ワークデータを削除する
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="LoginUserId">ログインユーザーID</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        Public Function DeleteProcessedOrdersRange(ByVal tran As OracleTransaction, ByVal LoginUserId As String, ByVal customerSettingId As Long) As Integer

            Dim DeleteCount As Integer = 0

            Dim pLoginUserId As String = If(String.IsNullOrWhiteSpace(LoginUserId), Nothing, LoginUserId.Trim())

            Dim sql As String = " DELETE FROM orders_stage os " &
                                " WHERE os.status = 'PROCESSED' "

            Dim isAdmin As Boolean = String.Equals(LoginUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)

            If Not isAdmin Then
                '通常ユーザー
                sql &= " AND EXISTS ( " &
                       "     SELECT 1 " &
                       "     FROM customer_setting_mst csm " &
                       "     WHERE csm.customer_setting_id = os.customer_setting_id " &
                       "     AND UPPER(csm.prod_mgmt_user_id) = UPPER(:p_login_user_id) " &
                       "     AND os.customer_setting_id = :p_customer_setting_id " &
                       " ) "
            Else
                '管理者
                sql &= " AND EXISTS ( " &
                       "     SELECT 1 " &
                       "     FROM customer_setting_mst csm " &
                       "     WHERE csm.customer_setting_id = os.customer_setting_id " &
                       "     AND os.customer_setting_id = :p_customer_setting_id " &
                       " ) "
            End If

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()

                If Not isAdmin Then
                    ' パラメータ設定
                    cmd.Parameters.Add(":p_login_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(pLoginUserId, 9)
                End If

                ' パラメータ設定
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId

                ' 実行（一致するレコードを一括で更新）
                DeleteCount = cmd.ExecuteNonQuery()

            End Using

            Return DeleteCount

        End Function

        ''' <summary>
        ''' 指定した取引先設定IDのデータを正規テーブル(orders)に反映する
        ''' </summary>
        ''' <param name="tran">呼び出し元から引き継ぐトランザクション</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新実行ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub UpdateOrdersFromStage(ByVal tran As OracleTransaction,
                                            ByVal customerSettingId As Long,
                                            ByVal updatedAt As DateTime,
                                            ByVal updateUserId As String,
                                            ByVal updatepgId As String)

            ' ORDER_IDをキーにして、受注ワーク(orders_stage)から正規(orders)へ更新をかける
            'Const sql As String = "
            '        MERGE INTO orders tgt
            '        USING (
            '            SELECT 
            '                order_id, 
            '                customer_order_no,
            '                due_date,
            '                customer_item_no,
            '                demand_qty, 
            '                status, 
            '                active_flag
            '            FROM orders_stage
            '            WHERE imp_file_stage_id = :p_imp_file_stage_id
            '        ) src
            '        ON (tgt.order_id = src.order_id)
            '        WHEN MATCHED THEN
            '            UPDATE SET
            '                tgt.customer_order_no = src.customer_order_no,                
            '                tgt.due_date = src.due_date,
            '                tgt.customer_item_no = src.customer_item_no,
            '                tgt.demand_qty = src.demand_qty,
            '                tgt.status = src.status,
            '                tgt.active_flag = src.active_flag,
            '                tgt.updated_at = :p_updated_at,
            '                tgt.updated_user_id = :p_user_id,
            '                tgt.updated_pg_id = :p_updated_pg_id"
            Const sql As String = "
                    MERGE INTO orders tgt
                    USING (
                        SELECT 
                            order_id, 
                            customer_order_no,
                            due_date,
                            customer_item_no,
                            demand_qty, 
                            status, 
                            active_flag,
                            updated_at,
                            updated_user_id,
                            updated_pg_id
                        FROM orders_stage
                        WHERE customer_setting_id = :p_customer_setting_id
                    ) src
                    ON (tgt.order_id = src.order_id)
                    WHEN MATCHED THEN
                        UPDATE SET
                            tgt.customer_order_no = src.customer_order_no,                
                            tgt.due_date = src.due_date,
                            tgt.customer_item_no = src.customer_item_no,
                            tgt.demand_qty = src.demand_qty,
                            tgt.status = src.status,
                            tgt.active_flag = src.active_flag,
                            tgt.updated_at = src.updated_at,
                            tgt.updated_user_id = src.updated_user_id,
                            tgt.updated_pg_id = src.updated_pg_id"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータの設定
                'cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                'cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                'cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                'cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                ' 実行（一致するレコードを一括で更新）
                Dim updatedCount As Integer = cmd.ExecuteNonQuery()

            End Using
        End Sub

        ''' <summary>
        ''' 指定した取引先設定IDのデータを正規テーブル(orders)に反映する
        ''' </summary>
        ''' <param name="tran">呼び出し元から引き継ぐトランザクション</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新実行ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Function InsertOrdersFromStage(ByVal tran As OracleTransaction,
                                            ByVal customerSettingId As Long,
                                            ByVal updatedAt As DateTime,
                                            ByVal updateUserId As String,
                                            ByVal updatepgId As String) As Integer

            Dim insertedCount As Integer = 0

            ' 受注ワーク(orders_stage)から正規(orders)へデータを追加する
            Const sql As String =
                        "INSERT INTO orders (" &
                        "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  remarks, delivery_code, total_ship_qty, transport_method, " &
                        "  customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                        "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                        "  reconcile_type, imp_run_id, status, active_flag, " &
                        "  created_at, created_user_id, created_pg_id, " &
                        "  updated_at, updated_user_id, updated_pg_id, " &
                        "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                        ") SELECT " &
                        "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  remarks, delivery_code, total_ship_qty, transport_method, " &
                        "  customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_stage_id, " &
                        "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                        "  reconcile_type, imp_run_id, status, active_flag, " &
                        "  :p_updated_at, :p_user_id, :p_updated_pg_id, " &
                        "  :p_updated_at, :p_user_id, :p_updated_pg_id, " &
                        "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                        "  FROM orders_stage " &
                        "  WHERE customer_setting_id = :p_customer_setting_id " &
                        "  AND status = 'PROCESSED' " &
                        "  AND order_id IS NULL " &
                        "  AND active_flag = 'Y' "

            'Const sql As String =
            '            "INSERT INTO orders (" &
            '            "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  remarks, delivery_code, total_ship_qty, transport_method, " &
            '            "  customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id" &
            '            ") SELECT " &
            '            "  customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  remarks, delivery_code, total_ship_qty, transport_method, " &
            '            "  customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_stage_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id " &
            '            "  FROM orders_stage " &
            '            "  WHERE customer_setting_id = :p_customer_setting_id " &
            '            "  AND status = 'PROCESSED' " &
            '            "  AND order_id IS NULL " &
            '            "  AND active_flag = 'Y' "

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータの設定
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                ' 実行
                insertedCount = cmd.ExecuteNonQuery()

            End Using

            Return insertedCount

        End Function

        ''' <summary>
        ''' 指定した取引先設定IDのデータ受注テーブル(orders)のデータを受注履歴テーブル(orders_history)に反映する
        ''' </summary>
        ''' <param name="tran">呼び出し元から引き継ぐトランザクション</param>
        ''' <param name="customerSettingId">対象の取引先設定ID</param>
        ''' <param name="updatedAt">更新日時</param>
        ''' <param name="updateUserId">更新実行ユーザーID</param>
        ''' <param name="updatepgId">更新プログラムID</param>
        Public Sub InsertHistoryFromOrders(ByVal tran As OracleTransaction,
                                            ByVal customerSettingId As Long,
                                            ByVal updatedAt As DateTime,
                                            ByVal updateUserId As String,
                                            ByVal updatepgId As String)

            ' 受注履歴テーブル(orders_history)へ受注テーブル(orders)のデータを追加する
            Const sql As String =
                        "INSERT INTO orders_history (" &
                        "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
                        "  customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                        "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                        "  reconcile_type, imp_run_id, status, active_flag, " &
                        "  created_at, created_user_id, created_pg_id, " &
                        "  updated_at, updated_user_id, updated_pg_id, " &
                        "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                        ") SELECT " &
                        "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
                        "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
                        "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
                        "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
                        "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
                        "  customer_order_line_no, " &
                        "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
                        "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
                        "  reconcile_type, imp_run_id, status, active_flag, " &
                        "  created_at, created_user_id, created_pg_id, " &
                        "  updated_at, updated_user_id, updated_pg_id, " &
                        "  stra_order_qty, stra_ship_qty, stra_order_backlog " &
                        "  FROM orders " &
                        "  WHERE customer_setting_id = :p_customer_setting_id " &
                        "  AND status = 'PROCESSED' " &
                        "  AND active_flag = 'Y' " &
                        "  AND created_at = :p_updated_at " &
                        "  AND created_user_id = :p_user_id " &
                        "  AND created_pg_id = :p_updated_pg_id "
            'Const sql As String =
            '            "INSERT INTO orders_history (" &
            '            "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
            '            "  customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id" &
            '            ") SELECT " &
            '            "  order_id, customer_setting_id, customer_code, billing_to, customer_order_no, demand_status, ship_to, " &
            '            "  order_date, due_date, ship_scheduled_date, customer_item_no, item_no, " &
            '            "  demand_qty, demand_unit, currency_code, ship_stock_location, company_id, " &
            '            "  product_code, billing_standard, ship_process_type, delivery_instr_flag, " &
            '            "  order_no, remarks, delivery_code, total_ship_qty, ship_date, transport_method, " &
            '            "  customer_order_line_no, " &
            '            "  pre_daily_order_qty, pre_daily_delivery_date, imp_file_id, " &
            '            "  order_type, prorated_type, customer_info_type, info_type, self_fcst_flag, self_fcst_delete_flag, " &
            '            "  reconcile_type, imp_run_id, status, active_flag, " &
            '            "  created_at, created_user_id, created_pg_id, " &
            '            "  updated_at, updated_user_id, updated_pg_id " &
            '            "  FROM orders " &
            '            "  WHERE customer_setting_id = :p_customer_setting_id " &
            '            "  AND status = 'PROCESSED' " &
            '            "  AND active_flag = 'Y' "

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータの設定
                cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int64).Value = customerSettingId
                cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = updatedAt
                cmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(updateUserId, 9)
                cmd.Parameters.Add(":p_updated_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(updatepgId, 150)

                ' 実行
                Dim insertedCount As Integer = cmd.ExecuteNonQuery()

            End Using
        End Sub

        ''' <summary>
        ''' 正規データ更新 (ORDERSテーブルのIMP_FILE_IDを更新する)
        ''' </summary>
        ''' <param name="tran">トランザクション</param>
        ''' <param name="newImpFileId">INSERT(取込ファイル)処理にて取得した新しいIMP_FILE_ID</param>
        ''' <param name="impFileStageId">現在処理中の一時取込ファイルID</param>
        Public Sub UpdateOrdersImpFileId(ByVal tran As OracleTransaction,
                                         ByVal newImpFileId As Long,
                                         ByVal impFileStageId As Long)

            ' 抽出条件：IMP_FILE_STAGE_ID (一時取込ファイルID)
            ' 更新内容：IMP_FILE_ID をセット
            Const sql As String = "
                    UPDATE orders
                    SET imp_file_id = :p_new_imp_file_id
                    WHERE imp_file_id = :p_imp_file_stage_id"

            Using cmd As New OracleCommand(sql, tran.Connection)
                cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Clear()

                ' パラメータの設定
                ' ※newImpFileId は「正規データ追加」メソッドの戻り値を使用します
                cmd.Parameters.Add(":p_new_imp_file_id", OracleDbType.Long).Value = newImpFileId
                cmd.Parameters.Add(":p_imp_file_stage_id", OracleDbType.Long).Value = impFileStageId

                ' 実行
                Dim updatedCount As Integer = cmd.ExecuteNonQuery()

                ' ログ出力など（必要に応じて）
                ' Console.WriteLine($"{updatedCount} 件の受注データを更新しました。")
            End Using
        End Sub

        ' 受注ワーク一覧取得
        Public Function GetOrdersStage(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing,
            Optional ByVal status As String = Nothing,
            Optional ByVal prodMgmtUserId As String = Nothing,
            Optional ByVal activeFlag As String = Nothing
        ) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  order_id                   AS ""OrderId"", ")
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
            sb.AppendLine("  total_ship_qty             AS ""TotalShipQty"", ")
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
            'sb.AppendLine("  stra_order_qty             AS ""StraOrderQty"" ")
            'sb.AppendLine("  stra_ship_qty              AS ""StraShipQty"" ")
            'sb.AppendLine("  stra_order_backlog         AS ""StraOrderBacklog"" ")
            sb.AppendLine("FROM orders_stage_view ")
            sb.AppendLine("WHERE 1=1 ")

            Dim prm As New List(Of OracleParameter)()

            ' 文字列を安全にLIKEパターンへ（%と_をエスケープしてから %term% に）
            Dim pCustomerCode As String = Utils.BuildLikePattern(customerCode, LikeMode.Contains)
            Dim pCustomerName As String = Utils.BuildLikePattern(customerName, LikeMode.Contains)
            Dim pProfitCenter As String = Utils.BuildLikePattern(profitCenter, LikeMode.Contains)
            Dim pCustomerUnitName As String = Utils.BuildLikePattern(customerUnitName, LikeMode.Contains)
            Dim pProdMgmtUserId As String = Utils.BuildLikePattern(prodMgmtUserId, LikeMode.Contains)

            Dim pActiveFlag As String = If(String.IsNullOrWhiteSpace(activeFlag), Nothing, activeFlag.Trim())

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

        '2026/03/26 酒井 ed

        ''' <summary>
        '''  受注ファイル出力 出荷状況チェック PROD_PLAN_STAGE 更新(Pharse-2 )
        ''' </summary>
        ''' <returns></returns>
        Public Function ShippingStatusCheck(conn As OracleConnection, tran As OracleTransaction, customerSettingId As Long, customerOrderNo As String, itemNo As String, updateDate As Date, userId As String) As String
            ' PROD_PLAN_HISTORY（生産計画履歴）と
            ' PROD_PLAN_STAGE（生産計画ワーク)テーブル
            ' 1) 両テーブルにある下記の フィールド値 を外部から与えて抽出します。
            '   CUSTOMER_SETTING_ID（取引先設定ID）
            '   CUSTOMER_ORDER_NO(客先発注No)
            '   ITEM_NO（品目No）
            ' 2) 抽出した 対になる両レコードのうち  次のフィールドが下記の条件の時 
            '   PROD_PLAN_HISTORY が STATUS = 'EXPORTED'、ACTIVE_FLAG = 'Y'
            '   PROD_PLAN_STAGE が STATUS = 'POST_PLAN_DUE_SET'、ACTIVE_FLAG = 'Y'
            ' 3)PROD_PLAN_STAGE の 下記フィールドを更新します。
            '   ACTIVE_FLAG(有効フラグ) = 'N'
            '   UPDATED_AT(更新日時) = [処理開始日時]
            '   UPDATED_USER_ID(更新ユーザーID) = [ログインユーザーID]
            '   UPDATED_PG_ID(更新プログラムID) = 'OrderExport'

            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("MERGE INTO PROD_PLAN_STAGE target ")
            sb.AppendLine("USING ( ")
            sb.AppendLine("    SELECT  ")
            sb.AppendLine("        CUSTOMER_SETTING_ID, ")
            sb.AppendLine("        CUSTOMER_ORDER_NO, ")
            sb.AppendLine("        ITEM_NO ")
            sb.AppendLine("    FROM  ")
            sb.AppendLine("        PROD_PLAN_HISTORY ")
            sb.AppendLine("    WHERE  ")
            sb.AppendLine("        CUSTOMER_SETTING_ID = :p_customer_setting_id ")
            sb.AppendLine("        AND CUSTOMER_ORDER_NO = :p_customer_order_no ")
            sb.AppendLine("        AND ITEM_NO = :p_item_no ")
            sb.AppendLine("        AND STATUS = 'EXPORTED' ")
            sb.AppendLine("        AND ACTIVE_FLAG = 'Y' ")
            sb.AppendLine(") source ")
            sb.AppendLine("ON ( ")
            sb.AppendLine("    target.CUSTOMER_SETTING_ID = source.CUSTOMER_SETTING_ID ")
            sb.AppendLine("    AND target.CUSTOMER_ORDER_NO = source.CUSTOMER_ORDER_NO ")
            sb.AppendLine("    AND target.ITEM_NO = source.ITEM_NO ")
            sb.AppendLine("    AND target.STATUS = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("    AND target.ACTIVE_FLAG = 'Y' ")
            sb.AppendLine(") ")
            sb.AppendLine("WHEN MATCHED THEN ")
            sb.AppendLine("UPDATE SET ")
            sb.AppendLine("    target.ACTIVE_FLAG = 'N', ")
            sb.AppendLine("    target.UPDATED_AT = :p_date, ")
            sb.AppendLine("    target.UPDATED_USER_ID = :p_user_id, ")
            sb.AppendLine("    target.UPDATED_PG_ID = 'OrderExport' ")

            Try
                'Using conn As New OracleConnection(_connectionString)
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Parameters.Add(":p_customer_setting_id", OracleDbType.Int32).Value = customerSettingId
                    cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 40).Value = customerOrderNo
                    cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 45).Value = itemNo
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
        ''' 生産計画ワーク update
        ''' </summary>
        ''' <returns></returns>
        Public Function ProductionPlanWorkUpdate(conn As OracleConnection, tran As OracleTransaction, updateDate As Date, userId As String) As String

            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("UPDATE PROD_PLAN_STAGE ")
            sb.AppendLine("SET ")
            sb.AppendLine("    STATUS = 'EXPORTED', ")
            sb.AppendLine("    UPDATED_AT = :p_date, ")
            sb.AppendLine("    UPDATED_USER_ID = :p_user_id, ")
            sb.AppendLine("    UPDATED_PG_ID = 'OrderExport' ")
            sb.AppendLine("WHERE ")
            sb.AppendLine("    STATUS = 'POST_PLAN_DUE_SET' ")
            sb.AppendLine("    AND ACTIVE_FLAG = 'Y' ")

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
    ''' ORDERS 受け渡し用の行DTO（Repository向け）
    ''' </summary>
    Public Class OrdersStageRow

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
        Public Property ImpRunId As Long?                      ' IMP_RUN_ID NUMBER (10,0)
        Public Property Status As String                        ' STATUS VARCHAR2(20)
        Public Property ActiveFlag As String                    ' ACTIVE_FLAG CHAR(1)

        ' ====== 数値系 ======
        Public Property DemandQty As Long?                      ' DEMAND_QTY NUMBER(10,0)
        Public Property TotalShipQty As Decimal?                ' TOTAL_SHIP_QTY NUMBER(18,6)
        Public Property PreDailyOrderQty As Decimal?            ' PRE_DAILY_ORDER_QTY NUMBER(18,6)
        Public Property ImpFileStageId As Long?                      ' IMP_FILE_STAGE_ID NUMBER(10,0)
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
        Public Property CreatedAt As DateTime           ' CREATED_AT
        Public Property CreatedUserId As String         ' CREATED_USER_ID
        Public Property CreatedPgId As String           ' CREATED_PG_ID
        Public Property UpdatedAt As DateTime           ' UPDATED_AT
        Public Property UpdatedUserId As String         ' UPDATED_USER_ID
        Public Property UpdatedPgId As String           ' UPDATED_PG_ID

        ' Stage
        ' ====== 主キー／識別系 ======
        Public Property StageId As Long?                ' STAGE_ID
        Public Property ImpFilesStageId As Long?        ' IMP_FILE_STAGE_ID



        Sub New()

        End Sub
        ''' <summary>
        ''' Copy constructor OrdersStageRow -> OrdersStageRow
        ''' </summary>
        ''' <param name="org"></param>
        Sub New(ByVal org As OrdersStageRow)
            OrderId = org.OrderId
            CustomerSettingId = org.CustomerSettingId
            CustomerCode = org.CustomerCode
            BillingTo = org.BillingTo
            CustomerOrderNo = org.CustomerOrderNo
            DemandStatus = org.DemandStatus
            ShipTo = org.ShipTo
            CustomerItemNo = org.CustomerItemNo
            ItemNo = org.ItemNo
            DemandUnit = org.DemandUnit
            CurrencyCode = org.CurrencyCode
            ShipStockLocation = org.ShipStockLocation
            CompanyId = org.CompanyId
            ProductCode = org.ProductCode
            BillingStandard = org.BillingStandard
            ShipProcessType = org.ShipProcessType
            DeliveryInstrFlag = org.DeliveryInstrFlag
            OrderNo = org.OrderNo
            Remarks = org.Remarks
            DeliveryCode = org.DeliveryCode
            TransportMethod = org.TransportMethod
            CustomerOrderLineNo = org.CustomerOrderLineNo
            CustomerInfoType = org.CustomerInfoType
            InfoType = org.InfoType
            SelfFcstFlag = org.SelfFcstFlag
            SelfFcstDeleteFlag = org.SelfFcstDeleteFlag
            ImpRunId = org.ImpRunId
            Status = org.Status
            ActiveFlag = org.ActiveFlag
            DemandQty = org.DemandQty
            TotalShipQty = org.TotalShipQty
            PreDailyOrderQty = org.PreDailyOrderQty
            ImpFileId = org.ImpFileId
            OrderType = org.OrderType
            ProratedType = org.ProratedType
            ReconcileType = org.ReconcileType
            'Pharse2
            StraOrderQty = org.StraOrderQty
            StraShipQty = org.StraShipQty
            StraOrderBacklog = org.StraOrderBacklog
            'Pharse2
            OrderDate = org.OrderDate
            DueDate = org.DueDate
            ShipScheduledDate = org.ShipScheduledDate
            ShipDate = org.ShipDate
            ShipPlanDate = org.ShipPlanDate
            PreDailyDeliveryDate = org.PreDailyDeliveryDate
            CreatedAt = org.CreatedAt
            CreatedUserId = org.CreatedUserId
            CreatedPgId = org.CreatedPgId
            UpdatedAt = org.UpdatedAt
            UpdatedUserId = org.UpdatedUserId
            UpdatedPgId = org.UpdatedPgId
            StageId = org.StageId
            ImpFilesStageId = org.ImpFilesStageId
        End Sub
        ''' <summary>
        ''' Copy constructor OrdersRow -> OrdersStageRow
        ''' </summary>
        ''' <param name="org"></param>
        Sub New(ByVal org As OrdersRow)
            OrderId = org.OrderId
            CustomerSettingId = org.CustomerSettingId
            CustomerCode = org.CustomerCode
            BillingTo = org.BillingTo
            CustomerOrderNo = org.CustomerOrderNo
            DemandStatus = org.DemandStatus
            ShipTo = org.ShipTo
            CustomerItemNo = org.CustomerItemNo
            ItemNo = org.ItemNo
            DemandUnit = org.DemandUnit
            CurrencyCode = org.CurrencyCode
            ShipStockLocation = org.ShipStockLocation
            CompanyId = org.CompanyId
            ProductCode = org.ProductCode
            BillingStandard = org.BillingStandard
            ShipProcessType = org.ShipProcessType
            DeliveryInstrFlag = org.DeliveryInstrFlag
            OrderNo = org.OrderNo
            Remarks = org.Remarks
            DeliveryCode = org.DeliveryCode
            TransportMethod = org.TransportMethod
            CustomerOrderLineNo = org.CustomerOrderLineNo
            CustomerInfoType = org.CustomerInfoType
            InfoType = org.InfoType
            SelfFcstFlag = org.SelfFcstFlag
            SelfFcstDeleteFlag = org.SelfFcstDeleteFlag
            ImpRunId = org.ImpRunId
            Status = org.Status
            ActiveFlag = org.ActiveFlag
            DemandQty = org.DemandQty
            TotalShipQty = org.TotalShipQty
            PreDailyOrderQty = org.PreDailyOrderQty
            ImpFileId = org.ImpFileId
            OrderType = org.OrderType
            ProratedType = org.ProratedType
            ReconcileType = org.ReconcileType
            'Pharse2
            StraOrderQty = org.StraOrderQty
            StraShipQty = org.StraShipQty
            StraOrderBacklog = org.StraOrderBacklog
            'Pharse2
            OrderDate = org.OrderDate
            DueDate = org.DueDate
            ShipScheduledDate = org.ShipScheduledDate
            ShipDate = org.ShipDate
            ShipPlanDate = org.ShipPlanDate
            PreDailyDeliveryDate = org.PreDailyDeliveryDate
            CreatedAt = org.CreatedAt
            CreatedUserId = org.CreatedUserId
            CreatedPgId = org.CreatedPgId
            UpdatedAt = org.UpdatedAt
            UpdatedUserId = org.UpdatedUserId
            UpdatedPgId = org.UpdatedPgId
        End Sub

        ''' <summary>
        ''' OrderStageRow リストを OrdersRow リストに変換をする
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        Public Shared Function ToOrdersRow(src As List(Of OrdersStageRow)) As List(Of OrdersRow)

            Dim dst = New List(Of OrdersRow)
            For Each st In src

                dst.Add(ToOrdersRow(st))
            Next
            Return dst

        End Function

        ''' <summary>
        ''' OrderStageRow を OrdersRow に変換をする
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        Public Shared Function ToOrdersRow(src As OrdersStageRow) As OrdersRow
            Dim dst = New OrdersRow
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
