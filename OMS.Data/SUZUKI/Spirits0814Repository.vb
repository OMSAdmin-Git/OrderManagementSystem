Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data.SUZUKI
    Public Class Spirits0814Repository
#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region
        ''' <summary>
        ''' Spirits0814Row class をDBに追加する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As Spirits0814Row) As String

            Dim records = New List(Of Spirits0814Row)()
            records.Add(row)
            Return InsertRange(conn, tran, records)

        End Function

        ''' <summary>
        ''' Spirits0814Row class リストをDBに追加する (元のコード同等呼び出し
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of Spirits0814Row))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, records)
                    tran.Commit()
                End Using
            End Using
        End Sub
        ''' <summary>
        ''' Spirits0814Row record 追加
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of Spirits0814Row)) As String
            Dim errorMessage As String = ""
            If records Is Nothing Then Return "Spirits0814Row InsertRange no record data."
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO SUZUKI_SPIRITS_0814 (")
                sb.AppendLine("  info_type_code, doc_title_type, payment_method_word_type, client_code, ")
                sb.AppendLine("  reserve1, publication_date, publication_time, target_reference_date_type, target_reference_date, ")
                sb.AppendLine("  customer_item_no, item_no, customer_item_no_process_no, customer_item_name, critical_safety_parts_code, ")
                sb.AppendLine("  packaging_code, capacity, supplier_code, supplier_factory_code, supplier_shipping_location, ")
                sb.AppendLine("  supplier_name, delivery_code, delivery_factory_code, delivery_location, arrange_manager, ")
                sb.AppendLine("  purchase_manager, first_article_type, leadtime_type, leadtime, jersey_number, ")
                sb.AppendLine("  delivery_cycle, supply_manager, supply_primary_acceptance_location, supply_departure_location, reserve2, ")
                sb.AppendLine("  customer_order_no, customer_order_no_process_no1, customer_order_no_process_no2, supply_line, supply_process, ")
                sb.AppendLine("  delivery_type, order_data_type, delivery_date_type, order_qty_type, delivery_date, ")
                sb.AppendLine("  delivery_time, order_qty, imp_file_id, imp_run_id, status, ")
                sb.AppendLine("  active_flag, created_at, created_user_id, created_pg_id, updated_at, ")
                sb.AppendLine("  updated_user_id, updated_pg_id ")

                sb.AppendLine(") VALUES (")
                sb.AppendLine(" :p_info_type_code, :p_doc_title_type, :p_payment_method_word_type, :p_client_code, ")
                sb.AppendLine(" :p_reserve1, :p_publication_date, :p_publication_time, :p_target_reference_date_type, :p_target_reference_date, ")
                sb.AppendLine(" :p_customer_item_no, :p_item_no, :p_customer_item_no_process_no, :p_customer_item_name, :p_critical_safety_parts_code, ")
                sb.AppendLine(" :p_packaging_code, :p_capacity, :p_supplier_code, :p_supplier_factory_code, :p_supplier_shipping_location, ")
                sb.AppendLine(" :p_supplier_name, :p_delivery_code, :p_delivery_factory_code, :p_delivery_location, :p_arrange_manager, ")
                sb.AppendLine(" :p_purchase_manager, :p_first_article_type, :p_leadtime_type, :p_leadtime, :p_jersey_number, ")
                sb.AppendLine(" :p_delivery_cycle, :p_supply_manager, :p_supply_primary_acceptance_location, :p_supply_departure_location, :p_reserve2, ")
                sb.AppendLine(" :p_customer_order_no, :p_customer_order_no_process_no1, :p_customer_order_no_process_no2, :p_supply_line, :p_supply_process, ")
                sb.AppendLine(" :p_delivery_type, :p_order_data_type, :p_delivery_date_type, :p_order_qty_type, :p_delivery_date, ")
                sb.AppendLine(" :p_delivery_time, :p_order_qty, :p_imp_file_id, :p_imp_run_id, :p_status, ")
                sb.AppendLine(" :p_active_flag, :p_created_at, :p_created_user_id, :p_created_pg_id, :p_updated_at, ")
                sb.AppendLine(" :p_updated_user_id, :p_updated_pg_id ")


                sb.AppendLine(")")

                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.Transaction = tran
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text

                    For Each r In records
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add(":p_info_type_code", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.InfoTypeCode, 4)
                        cmd.Parameters.Add(":p_doc_title_type", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.DocTitleType, 2)
                        cmd.Parameters.Add(":p_payment_method_word_type", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.PaymentMethodWordType, 2)
                        cmd.Parameters.Add(":p_client_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.ClientCode, 12)
                        cmd.Parameters.Add(":p_reserve1", OracleDbType.Varchar2, 16).Value = SafeVarchar(r.Reserve1, 16)
                        cmd.Parameters.Add(":p_publication_date", OracleDbType.Date).Value = SafeNullable(r.PublicationDate)
                        cmd.Parameters.Add(":p_publication_time", OracleDbType.Int16).Value = SafeNullable(r.PublicationTime)
                        cmd.Parameters.Add(":p_target_reference_date_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.TargetReferenceDateType, 1)
                        cmd.Parameters.Add(":p_target_reference_date", OracleDbType.Varchar2, 8).Value = SafeVarchar(r.TargetReferenceDate, 8)
                        cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 15).Value = SafeVarchar(r.CustomerItemNo, 15)
                        cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ItemNo, 45)
                        cmd.Parameters.Add(":p_customer_item_no_process_no", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CustomerItemNoProcessNo, 2)
                        cmd.Parameters.Add(":p_customer_item_name", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.CustomerItemName, 12)
                        cmd.Parameters.Add(":p_critical_safety_parts_code", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.CriticalSafetyPartsCode, 1)
                        cmd.Parameters.Add(":p_packaging_code", OracleDbType.Varchar2, 5).Value = SafeVarchar(r.PackagingCode, 5)
                        cmd.Parameters.Add(":p_capacity", OracleDbType.Int16).Value = SafeNullable(r.Capacity)
                        cmd.Parameters.Add(":p_supplier_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.SupplierCode, 12)
                        cmd.Parameters.Add(":p_supplier_factory_code", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.SupplierFactoryCode, 3)
                        cmd.Parameters.Add(":p_supplier_shipping_location", OracleDbType.Varchar2, 8).Value = SafeVarchar(r.SupplierShippingLocation, 8)
                        cmd.Parameters.Add(":p_supplier_name", OracleDbType.Varchar2, 50).Value = SafeVarchar(r.SupplierName, 50)
                        cmd.Parameters.Add(":p_delivery_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.DeliveryCode, 12)
                        cmd.Parameters.Add(":p_delivery_factory_code", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.DeliveryFactoryCode, 3)
                        cmd.Parameters.Add(":p_delivery_location", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.DeliveryLocation, 4)
                        cmd.Parameters.Add(":p_arrange_manager", OracleDbType.Varchar2, 7).Value = SafeVarchar(r.PurchaseManager, 7)
                        cmd.Parameters.Add(":p_purchase_manager", OracleDbType.Varchar2, 7).Value = SafeVarchar(r.PurchaseManager, 7)
                        cmd.Parameters.Add(":p_first_article_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.FirstArticleType, 1)
                        cmd.Parameters.Add(":p_leadtime_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.LeadtimeType, 1)
                        cmd.Parameters.Add(":p_leadtime", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.Leadtime, 3)
                        cmd.Parameters.Add(":p_jersey_number", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.JerseyNumber, 3)
                        cmd.Parameters.Add(":p_delivery_cycle", OracleDbType.Varchar2, 5).Value = SafeVarchar(r.DeliveryCycle, 5)
                        cmd.Parameters.Add(":p_supply_manager", OracleDbType.Varchar2, 5).Value = SafeVarchar(r.SupplyManager, 5)
                        cmd.Parameters.Add(":p_supply_primary_acceptance_location", OracleDbType.Varchar2, 3).Value = SafeVarchar(r.SupplyPrimaryAcceptanceLocation, 3)
                        cmd.Parameters.Add(":p_supply_departure_location", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.SupplyDepartureLocation, 4)
                        cmd.Parameters.Add(":p_reserve2", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.Reserve2, 4)
                        cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 13).Value = SafeVarchar(r.CustomerOrderNo, 13)
                        cmd.Parameters.Add(":p_customer_order_no_process_no1", OracleDbType.Varchar2, 13).Value = SafeVarchar(r.CustomerOrderNoProcessNo1, 13)
                        cmd.Parameters.Add(":p_customer_order_no_process_no2", OracleDbType.Varchar2, 13).Value = SafeVarchar(r.CustomerOrderNoProcessNo2, 13)
                        cmd.Parameters.Add(":p_supply_line", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.SupplyLine, 4)
                        cmd.Parameters.Add(":p_supply_process", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.SupplyProcess, 4)
                        cmd.Parameters.Add(":p_delivery_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.DeliveryDateType, 1)
                        cmd.Parameters.Add(":p_order_data_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.DeliveryDateType, 1)
                        cmd.Parameters.Add(":p_delivery_date_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.DeliveryDateType, 1)
                        cmd.Parameters.Add(":p_order_qty_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.OrderQtyType, 1)
                        cmd.Parameters.Add(":p_delivery_date", OracleDbType.Date).Value = SafeNullable(r.DeliveryDate)
                        cmd.Parameters.Add(":p_delivery_time", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.DeliveryTime, 4)
                        cmd.Parameters.Add(":p_order_qty", OracleDbType.Int64).Value = SafeNullable(r.OrderQty)
                        cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = SafeNullable(r.ImpFileId)
                        cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = SafeNullable(r.ImpRunId)
                        cmd.Parameters.Add(":p_status", OracleDbType.Char, 20).Value = SafeVarchar(r.Status, 20)
                        cmd.Parameters.Add(":p_active_flag", OracleDbType.Char, 1).Value = SafeVarchar(r.ActiveFlag, 1)
                        cmd.Parameters.Add(":p_created_at", OracleDbType.Date).Value = SafeNullable(r.CreatedAt)
                        cmd.Parameters.Add(":p_created_user_id", OracleDbType.Varchar2, 9).Value = SafeVarchar(r.CreatedUserId, 9)
                        cmd.Parameters.Add(":p_created_pg_id", OracleDbType.Varchar2, 150).Value = SafeVarchar(r.CreatedPgId, 150)
                        cmd.Parameters.Add(":p_updated_at", OracleDbType.Date).Value = SafeNullable(r.UpdatedAt)
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

        End Function


    End Class
    Public Class Spirits0814Row
        Public Property InfoTypeId As Long                              ' 情報区分ID			NUMBER		10,0
        Public Property InfoTypeCode As String                          ' 情報区分コード		VARCHAR2	4 CHAR
        Public Property DocTitleType As String                          ' 書面タイトル区分		VARCHAR2	2 CHAR
        Public Property PaymentMethodWordType As String                 ' 支払方法等文言区分	VARCHAR2	2 CHAR
        Public Property ClientCode As String                            ' 発注者コード			VARCHAR2	12 CHAR
        Public Property Reserve1 As String                              ' 予備1					VARCHAR2	16 CHAR
        Public Property PublicationDate As Date?                        ' 発行日				DATE		
        Public Property PublicationTime As Integer?                     ' 発行時刻				NUMBER		4,0
        Public Property TargetReferenceDateType As String               ' 対象基準日区分		VARCHAR2	1 CHAR
        Public Property TargetReferenceDate As String                   ' 対象基準日			VARCHAR2	8 CHAR
        Public Property CustomerItemNo As String                        ' 部品番号				VARCHAR2	15 CHAR
        Public Property ItemNo As String                                ' 品目No				VARCHAR2	45 CHAR
        Public Property CustomerItemNoProcessNo As String               ' 部品番号識別-1		VARCHAR2	2 CHAR
        Public Property CustomerItemName As String                      ' 部品名称				VARCHAR2	30 CHAR
        Public Property CriticalSafetyPartsCode As String               ' 重要保安部品コード	VARCHAR2	1 CHAR
        Public Property PackagingCode As String                         ' 荷姿コード			VARCHAR2	5 CHAR
        Public Property Capacity As Integer?                            ' 収容数				NUMBER		5,0
        Public Property SupplierCode As String                          ' 仕入先コード			VARCHAR2	12 CHAR
        Public Property SupplierFactoryCode As String                   ' 仕入先工場コード		VARCHAR2	3 CHAR
        Public Property SupplierShippingLocation As String              ' 仕入先出荷場所		VARCHAR2	8 CHAR
        Public Property SupplierName As String                          ' 仕入先名称			VARCHAR2	50 CHAR
        Public Property DeliveryCode As String                          ' 納入先コード			VARCHAR2	12 CHAR
        Public Property DeliveryFactoryCode As String                   ' 納入先工場コード		VARCHAR2	3 CHAR
        Public Property DeliveryLocation As String                      ' 納入場所				VARCHAR2	4 CHAR
        Public Property ArrangeManager As String                        ' 手配担当				VARCHAR2	7 CHAR
        Public Property PurchaseManager As String                       ' 購買担当				VARCHAR2	7 CHAR
        Public Property FirstArticleType As String                      ' 初物区分				VARCHAR2	1 CHAR
        Public Property LeadtimeType As String                          ' 先行時間区分			VARCHAR2	1 CHAR
        Public Property Leadtime As String                              ' 先行時間				VARCHAR2	3 CHAR
        Public Property JerseyNumber As String                          ' 背番号				VARCHAR2	3 CHAR
        Public Property DeliveryCycle As String                         ' サイクル				VARCHAR2	5 CHAR
        Public Property SupplyManager As String                         ' 支給担当				VARCHAR2	5 CHAR
        Public Property SupplyPrimaryAcceptanceLocation As String       ' 支給品一次受入場所	VARCHAR2	3 CHAR
        Public Property SupplyDepartureLocation As String               ' 支給品出庫場所		VARCHAR2	4 CHAR
        Public Property Reserve2 As String                              ' 予備2					VARCHAR2	4 CHAR
        Public Property CustomerOrderNo As String                       ' 発行番号				VARCHAR2	13 CHAR
        Public Property CustomerOrderNoProcessNo1 As String             ' 発行番号識別-1		VARCHAR2	13 CHAR
        Public Property CustomerOrderNoProcessNo2 As String             ' 発行番号識別-2		VARCHAR2	13 CHAR
        Public Property SupplyLine As String                            ' 供給ライン			VARCHAR2	4 CHAR
        Public Property SupplyProcess As String                         ' 供給工程				VARCHAR2	10 CHAR
        Public Property DeliveryType As String                          ' 納入指示区分			VARCHAR2	5 CHAR
        Public Property OrderDataType As String                         ' 注文データ区分		VARCHAR2	1 CHAR
        Public Property DeliveryDateType As String                      ' 納入日区分			VARCHAR2	1 CHAR
        Public Property OrderQtyType As String                          ' 数量区分				VARCHAR2	1 CHAR
        Public Property DeliveryDate As Date?                           ' 納入指示日			DATE		
        Public Property DeliveryTime As String                          ' 納入指示時刻			VARCHAR2	4 CHAR
        Public Property OrderQty As Long?                               ' 納入指示数			NUMBER		8,0
        Public Property ImpFileId As Long?                              ' 取込ファイルID		NUMBER		10,0
        Public Property ImpRunId As Long?                               ' 取込実行ID			NUMBER		10,0
        Public Property Status As String                                ' ステータス			VARCHAR2	20
        Public Property ActiveFlag As String                            ' 有効フラグ			CHAR		1
        Public Property CreatedAt As Date?                              ' 登録日時				DATE		
        Public Property CreatedUserId As String                         ' 登録ユーザーID		VARCHAR2	9
        Public Property CreatedPgId As String                           ' 登録プログラムID		VARCHAR2	150
        Public Property UpdatedAt As Date?                              ' 更新日時				DATE		
        Public Property UpdatedUserId As String                         ' 更新ユーザーID		VARCHAR2	9
        Public Property UpdatedPgId As String                           ' 更新プログラムID		VARCHAR2	150

    End Class
End Namespace
