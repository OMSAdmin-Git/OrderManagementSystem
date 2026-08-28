Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data.SUZUKI


    Public Class Spirits0501And0502Repository
#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region
        ''' <summary>
        ''' Spirits0501And0502Row class をDBに追加する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As Spirits0501And0502Row) As String

            Dim records = New List(Of Spirits0501And0502Row)()
            records.Add(row)
            Return InsertRange(conn, tran, records)

        End Function

        ''' <summary>
        ''' Spirits0501And0502Row class リストをDBに追加する (元のコード同等呼び出し
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of Spirits0501And0502Row))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, records)
                    tran.Commit()
                End Using
            End Using
        End Sub
        ''' <summary>
        ''' Spirits0501And0502Row record 追加
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of Spirits0501And0502Row)) As String
            Dim errorMessage As String = ""
            If records Is Nothing Then Return "Spirits0501And0502Row InsertRange no record data."
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO SUZUKI_SPIRITS_0501AND0502 (")
                sb.AppendLine("  nfo_type_code, doc_title_type, payment_method_word_type, client_code, ")
                sb.AppendLine("  contractor_code, contractor_office_code, publication_date, publication_time, target_reference_date_type, ")
                sb.AppendLine("  target_reference_date, customer_item_no, item_no, customer_item_no_process_no, customer_item_name, ")
                sb.AppendLine("  suppliers_code, reserve1, arrange_manager, purchase_manager, reserve2, ")
                sb.AppendLine("  order_data_type, delivery_date_type, order_qty_type, delivery_date, order_qty, ")
                sb.AppendLine("  imp_file_id, imp_run_id, active_flag, created_at, created_user_id, ")
                sb.AppendLine("  created_pg_id, updated_at, updated_user_id, updated_pg_id ")
                sb.AppendLine(") VALUES (")
                sb.AppendLine(" :p_info_type_code, :p_doc_title_type, :p_payment_method_word_type, :p_client_code, ")
                sb.AppendLine(" :p_contractor_code, :p_contractor_office_code, :p_publication_date, :p_publication_time, :p_target_reference_date_type, ")
                sb.AppendLine(" :p_target_reference_date, :p_customer_item_no, :p_item_no, :p_customer_item_no_process_no, :p_customer_item_name, ")
                sb.AppendLine(" :p_suppliers_code, :p_reserve1, :p_arrange_manager, :p_purchase_manager, :p_reserve2, ")
                sb.AppendLine(" :p_order_data_type, :p_delivery_date_type, :p_order_qty_type, :p_delivery_date, :p_order_qty, ")
                sb.AppendLine(" :p_imp_file_id, :p_imp_run_id, :p_active_flag, :p_created_at, :p_created_user_id, ")
                sb.AppendLine(" :p_created_pg_id, :p_updated_at, :p_updated_user_id, :p_updated_pg_id ")
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
                        cmd.Parameters.Add(":p_contractor_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.ContractorCode, 12)
                        cmd.Parameters.Add(":p_contractor_office_code", OracleDbType.Varchar2, 4).Value = SafeVarchar(r.ContractorOfficeCode, 4)
                        cmd.Parameters.Add(":p_publication_date", OracleDbType.Date).Value = SafeNullable(r.PublicationDate)
                        cmd.Parameters.Add(":p_publication_time", OracleDbType.Int16).Value = SafeNullable(r.PublicationTime)
                        cmd.Parameters.Add(":p_target_reference_date_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.TargetReferenceDateType, 1)
                        cmd.Parameters.Add(":p_target_reference_date", OracleDbType.Varchar2, 8).Value = SafeVarchar(r.TargetReferenceDate, 8)
                        cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 15).Value = SafeVarchar(r.CustomerItemNo, 15)
                        cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ItemNo, 45)
                        cmd.Parameters.Add(":p_customer_item_no_process_no", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CustomerItemNoProcessNo, 2)
                        cmd.Parameters.Add(":p_customer_item_name", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.CustomerItemName, 12)
                        cmd.Parameters.Add(":p_suppliers_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.SuppliersCode, 12)
                        cmd.Parameters.Add(":p_reserve1", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.Reserve1, 12)
                        cmd.Parameters.Add(":p_arrange_manager", OracleDbType.Varchar2, 7).Value = SafeVarchar(r.ArrangeManager, 7)
                        cmd.Parameters.Add(":p_purchase_manager", OracleDbType.Varchar2, 7).Value = SafeVarchar(r.PurchaseManager, 7)
                        cmd.Parameters.Add(":p_reserve2", OracleDbType.Varchar2, 8).Value = SafeVarchar(r.Reserve2, 8)
                        cmd.Parameters.Add(":p_order_data_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.OrderDataType, 1)
                        cmd.Parameters.Add(":p_delivery_date_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.DeliveryDateType, 1)
                        cmd.Parameters.Add(":p_order_qty_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.OrderQtyType, 1)
                        cmd.Parameters.Add(":p_delivery_date", OracleDbType.Date).Value = SafeNullable(r.DeliveryDate)
                        cmd.Parameters.Add(":p_order_qty", OracleDbType.Int64).Value = SafeNullable(r.OrderQty)
                        cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = SafeNullable(r.ImpFileId)
                        cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = SafeNullable(r.ImpRunId)
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

    Public Class Spirits0501And0502Row
        Public Property InfoTypeId As Long                  ' 情報区分ID			NUMBER			10,0
        Public Property InfoTypeCode As String              ' 情報区分コード		VARCHAR2		4 CHAR
        Public Property DocTitleType As String              ' 書面タイトル区分		VARCHAR2		2 CHAR
        Public Property PaymentMethodWordType As String     ' 支払方法等文言区分	VARCHAR2		2 CHAR
        Public Property ClientCode As String                ' 発注者コード			VARCHAR2		12 CHAR
        Public Property ContractorCode As String            ' 受注者コード			VARCHAR2		12 CHAR
        Public Property ContractorOfficeCode As String      ' 受注者事業所コード	VARCHAR2		4 CHAR
        Public Property PublicationDate As Date?            ' 発行日				DATE			
        Public Property PublicationTime As Integer?          ' 発行時刻				NUMBER			4,0
        Public Property TargetReferenceDateType As String   ' 対象基準日区分		VARCHAR2		1 CHAR
        Public Property TargetReferenceDate As String       ' 対象基準日			VARCHAR2		8 CHAR
        Public Property CustomerItemNo As String            ' 部品番号				VARCHAR2		15 CHAR
        Public Property ItemNo As String                    ' 品目No				VARCHAR2		45 CHAR
        Public Property CustomerItemNoProcessNo As String   ' 部品番号識別-1		VARCHAR2		2 CHAR
        Public Property CustomerItemName As String          ' 部品名称				VARCHAR2		30 CHAR
        Public Property SuppliersCode As String             ' 仕入先コード			VARCHAR2		12 CHAR
        Public Property Reserve1 As String                  ' 予備					VARCHAR2		12 CHAR
        Public Property ArrangeManager As String            ' 手配担当				VARCHAR2		7 CHAR
        Public Property PurchaseManager As String           ' 購買担当				VARCHAR2		7 CHAR
        Public Property Reserve2 As String                  ' 予備					VARCHAR2		8 CHAR
        Public Property OrderDataType As String             ' 注文データ区分		VARCHAR2		1 CHAR
        Public Property DeliveryDateType As String          ' 納入日区分			VARCHAR2		1 CHAR
        Public Property OrderQtyType As String              ' 数量区分				VARCHAR2		1 CHAR
        Public Property DeliveryDate As Date?               ' 納入指示日			DATE			
        Public Property OrderQty As Long?                   ' 所要数量				NUMBER			8,0
        Public Property ImpFileId As Long?                  ' 取込ファイルID		NUMBER			10,0
        Public Property ImpRunId As Long?                   ' 取込実行ID			NUMBER			10,0
        Public Property ActiveFlag As String                ' 有効フラグ			CHAR			1
        Public Property CreatedAt As Date?                  ' 登録日時				DATE			
        Public Property CreatedUserId As String             ' 登録ユーザーID		VARCHAR2		9
        Public Property CreatedPgId As String               ' 登録プログラムID		VARCHAR2		150
        Public Property UpdatedAt As Date?                  ' 更新日時				DATE			
        Public Property UpdatedUserId As String             ' 更新ユーザーID		VARCHAR2		9
        Public Property UpdatedPgId As String               ' 更新プログラムID		VARCHAR2		150
    End Class
End Namespace