Imports System.Data
Imports System.Text
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Namespace OMS.Data.SUZUKI
    Public Class Spirits663NAnd663SAnd66Repository
#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

        ''' <summary>
        ''' Spirits663NAnd663SAnd66Row class をDBに追加する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="row"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function Insert(conn As OracleConnection, tran As OracleTransaction, row As Spirits663NAnd663SAnd66Row) As String

            Dim records = New List(Of Spirits663NAnd663SAnd66Row)()
            records.Add(row)
            Return InsertRange(conn, tran, records)

        End Function

        ''' <summary>
        ''' Spirits663NAnd663SAnd66Row class リストをDBに追加する (元のコード同等呼び出し
        ''' </summary>
        ''' <param name="records"></param>
        Public Sub InsertRange(records As IEnumerable(Of Spirits663NAnd663SAnd66Row))

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    InsertRange(conn, tran, records)
                    tran.Commit()
                End Using
            End Using
        End Sub
        ''' <summary>
        ''' Spirits663NAnd663SAnd66Row record 追加
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="records"></param>
        ''' <returns>string: Number:xxxx Errormessage</returns>
        Public Function InsertRange(conn As OracleConnection, tran As OracleTransaction, records As List(Of Spirits663NAnd663SAnd66Row)) As String
            Dim errorMessage As String = ""
            If records Is Nothing Then Return "OrdersRow InsertRange no record data."
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine($"INSERT INTO SUZUKI_SPIRITS_663NAND663SAND664T (")
                sb.AppendLine("  info_type_code, doc_title_type, payment_method_word_type, client_code, ")
                sb.AppendLine("  contractor_code, contractor_office_code, publication_date, target_reference_date_type, target_reference_date, ")
                sb.AppendLine("  supplier_code, delivery_form_type, customer_item_no, item_no, process_type, ")
                sb.AppendLine("  customer_item_no_process_no1, customer_order_no, pickup_instructions_times, old_delivery_date, delivery_date, ")
                sb.AppendLine("  old_order_qty, order_qty, cahnge_reason, reserve, imp_file_id, ")
                sb.AppendLine("  imp_run_id, status, active_flag, created_at, created_user_id, ")
                sb.AppendLine("  created_pg_id, updated_at, updated_user_id, updated_pg_id ")
                sb.AppendLine(") VALUES (")
                sb.AppendLine(" :p_info_type_code, :p_doc_title_type, :p_payment_method_word_type, :p_client_code, ")
                sb.AppendLine(" :p_contractor_code, :p_contractor_office_code, :p_publication_date, :p_target_reference_date_type, :p_target_reference_date, ")
                sb.AppendLine(" :p_supplier_code, :p_delivery_form_type, :p_customer_item_no, :p_item_no, :p_process_type, ")
                sb.AppendLine(" :p_customer_item_no_process_no1, :p_customer_order_no, :p_pickup_instructions_times, :p_old_delivery_date, :p_delivery_date, ")
                sb.AppendLine(" :p_old_order_qty, :p_order_qty, :p_cahnge_reason, :p_reserve, :p_imp_file_id, ")
                sb.AppendLine(" :p_imp_run_id, :p_status, :p_active_flag, :p_created_at, :p_created_user_id, ")
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
                        cmd.Parameters.Add(":p_target_reference_date_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.TargetReferenceDateType, 1)
                        cmd.Parameters.Add(":p_target_reference_date", OracleDbType.Varchar2, 8).Value = SafeVarchar(r.TargetReferenceDate, 8)

                        cmd.Parameters.Add(":p_supplier_code", OracleDbType.Varchar2, 12).Value = SafeVarchar(r.SupplierCode, 12)
                        cmd.Parameters.Add(":p_delivery_form_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.DeliveryFormType, 1)
                        cmd.Parameters.Add(":p_customer_item_no", OracleDbType.Varchar2, 15).Value = SafeVarchar(r.CustomerItemNo, 15)
                        cmd.Parameters.Add(":p_item_no", OracleDbType.Varchar2, 45).Value = SafeVarchar(r.ItemNo, 45)
                        cmd.Parameters.Add(":p_process_type", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.ProcessType, 1)

                        cmd.Parameters.Add(":p_customer_item_no_process_no1", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CustomerItemNoProcessNo1, 2)
                        cmd.Parameters.Add(":p_customer_order_no", OracleDbType.Varchar2, 13).Value = SafeVarchar(r.CustomerOrderNo, 13)
                        cmd.Parameters.Add(":p_pickup_instructions_times", OracleDbType.Varchar2, 1).Value = SafeVarchar(r.PickupInstructionsTimes, 1)
                        cmd.Parameters.Add(":p_old_delivery_date", OracleDbType.Date).Value = SafeNullable(r.OldDeliveryDate)
                        cmd.Parameters.Add(":p_delivery_date", OracleDbType.Date).Value = SafeNullable(r.DeliveryDate)

                        cmd.Parameters.Add(":p_old_order_qty", OracleDbType.Int64).Value = SafeNullable(r.OldOrderQty)
                        cmd.Parameters.Add(":p_order_qty", OracleDbType.Int64).Value = SafeNullable(r.OrderQty)
                        cmd.Parameters.Add(":p_cahnge_reason", OracleDbType.Varchar2, 2).Value = SafeVarchar(r.CahngeReason, 2)
                        cmd.Parameters.Add(":p_reserve", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.Reserve, 20)
                        cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = SafeNullable(r.ImpFileId)

                        cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = SafeNullable(r.ImpRunId)
                        cmd.Parameters.Add(":p_status", OracleDbType.Varchar2, 20).Value = SafeVarchar(r.Status, 20)
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

    ''' <summary>
    ''' Spirits663NAnd663SAnd66 受け渡し用の行DTO（Repository向け）
    ''' </summary>
    Public Class Spirits663NAnd663SAnd66Row

        Public Property InfoTypeId As Long                          ' 情報区分ID          NUMBER      10,0
        Public Property InfoTypeCode As String                      ' 情報区分コード      VARCHAR2    4 CHAR
        Public Property DocTitleType As String                      ' 書面タイトル区分    VARCHAR2    2 CHAR
        Public Property PaymentMethodWordType As String             ' 支払方法等文言区分  VARCHAR2    2 CHAR
        Public Property ClientCode As String                        ' 発注者コード        VARCHAR2    12 CHAR
        Public Property ContractorCode As String                    ' 受注者コード        VARCHAR2    12 CHAR
        Public Property ContractorOfficeCode As String              ' 受注者事業所コード  VARCHAR2    4 CHAR
        Public Property PublicationDate As Date?                    ' 発行日              DATE        
        Public Property TargetReferenceDateType As String           ' 対象基準日区分      VARCHAR2    1 CHAR
        Public Property TargetReferenceDate As String               ' 対象基準日          VARCHAR2    8 CHAR
        Public Property SupplierCode As String                      ' 仕入先コード        VARCHAR2    12 CHAR
        Public Property DeliveryFormType As String                  ' 完成購中区分        VARCHAR2    1 CHAR
        Public Property CustomerItemNo As String                    ' 部品番号            VARCHAR2    15 CHAR
        Public Property ItemNo As String                            ' 品目No              VARCHAR2    45 CHAR
        Public Property ProcessType As String                       ' 工程区分            VARCHAR2    1 CHAR
        Public Property CustomerItemNoProcessNo1 As String          ' 部品番号識別-1      VARCHAR2    2 CHAR
        Public Property CustomerOrderNo As String                   ' 発行番号            VARCHAR2    13 CHAR
        Public Property PickupInstructionsTimes As String           ' 引取指示回数        VARCHAR2    1 CHAR
        Public Property OldDeliveryDate As Date?                    ' 旧納入指示日        DATE        
        Public Property DeliveryDate As Date?                       ' 納入指示日          DATE        
        Public Property OldOrderQty As Long?                        ' 旧納入指示数        NUMBER      7,0
        Public Property OrderQty As Long?                           ' 納入指示数          NUMBER      7,0
        Public Property CahngeReason As String                      ' 変更理由            VARCHAR2    2 CHAR
        Public Property Reserve As String                           ' 予備                VARCHAR2    20 CHAR
        Public Property ImpFileId As Long?                          ' 取込ファイルID      NUMBER      10,0
        Public Property ImpRunId As Long?                           ' 取込実行ID          NUMBER      10,0
        Public Property Status As String                            ' ステータス          VARCHAR2    20
        Public Property ActiveFlag As String                        ' 有効フラグ          CHAR        1
        Public Property CreatedAt As Date?                          ' 登録日時            DATE        
        Public Property CreatedUserId As String                     ' 登録ユーザーID      VARCHAR2    9
        Public Property CreatedPgId As String                       ' 登録プログラムID    VARCHAR2    150
        Public Property UpdatedAt As Date?                          ' 更新日時            DATE        
        Public Property UpdatedUserId As String                     ' 更新ユーザーID      VARCHAR2    9
        Public Property UpdatedPgId As String                       ' 更新プログラムID    VARCHAR2    150
    End Class

End Namespace