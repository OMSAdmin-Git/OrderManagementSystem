Imports System.Data
Imports System.Text
Imports DocumentFormat.OpenXml.Bibliography
Imports OMS.Common
Imports OMS.Data.SectmRepository
Imports OMS.Data.UsrDeffIdfRepository
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data

    Public Class ShproutmRepository
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        ''' <summary>
        ''' 輸送リードタイム取得
        ''' </summary>
        ''' <param name="custmerCode"></param>
        ''' <param name="deliveryCode"></param>
        ''' <param name="customerItemNo"></param>
        ''' <returns></returns>
        Public Function GetTransferLeadTime(custmerCode As String, deliveryCode As String, customerItemNo As String) As Int16
            Dim dt As New DataTable()
            Dim lt = -1
            Try
                Using conn As New OracleConnection(_connectionString)
                    conn.Open()
                    Using tran As OracleTransaction = conn.BeginTransaction()

                        Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                            cmd.CommandText = "WITH target_input AS (
                                              -- 1.2. 外部から値を与え、取引先コードと納入先コードを連結して「取引先コード1」を作成
                                              SELECT 
                                                :p_custmerCode || :p_deliveryCode AS customerDeliveryCode,
                                                :p_customerItemNo AS target_item_no
                                              FROM dual
                                            ),
                                            first_product AS (
                                              -- 3. PRDSLSODRM から顧客番号と客先品目Noが一致する先頭の製品コードを取得
                                              SELECT p.fprdcd
                                              FROM prdslsodrm p
                                              JOIN target_input ti ON p.fcustcd = ti.customerDeliveryCode 
                                                                  AND p.fcustitemno = ti.target_item_no
                                              ORDER BY p.rowid -- 先頭を特定するソート（主キー等への変更を推奨）
                                              FETCH FIRST 1 ROW ONLY
                                            ),
                                            first_warehouse AS (
                                              -- 4. ITEMM から製品コードが一致する先頭の出荷在庫場所を取得
                                              SELECT i.fprmwhcd
                                              FROM itemm i
                                              JOIN first_product fp ON i.fitemno = fp.fprdcd
                                              ORDER BY i.rowid -- 先頭を特定するソート
                                              FETCH FIRST 1 ROW ONLY
                                            )
                                            -- 5. SHPROUTM から条件に一致する全レコードを抽出し、優先順位の昇順でソート
                                            SELECT s.*
                                            FROM shproutm s
                                            JOIN target_input ti ON s.fshptocd = ti.customerDeliveryCode
                                            JOIN first_warehouse fw ON s.fprmwhcd = fw.fprmwhcd
                                            ORDER BY s.fpriority ASC 
                                            FETCH FIRST 1 ROW ONLY "
                            cmd.Parameters.Add(":p_custmerCode", OracleDbType.Char, 25).Value = custmerCode
                            cmd.Parameters.Add(":p_deliveryCode", OracleDbType.Char, 25).Value = deliveryCode
                            cmd.Parameters.Add(":p_customerItemNo", OracleDbType.Char, 45).Value = customerItemNo

                            Using reader As OracleDataReader = cmd.ExecuteReader()
                                dt.Load(reader)
                            End Using

                            If (dt.Rows.Count > 0) Then
                                Dim dtr = ToClass(dt.Rows(0))
                                lt = dtr.TransferLeadTime
                            End If

                        End Using
                        tran.Commit()
                    End Using
                    'conn.Close()
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try
            Return lt

        End Function

        ''' <summary>
        ''' 輸送リードタイム取得
        ''' </summary>
        ''' <param name="shipTo"></param>
        ''' <param name="priority"></param>
        ''' <returns></returns>
        Public Function GetTransferLeadTime(shipTo As String, priority As Integer) As Int16
            Dim dtrs = GetUsrShproutm(shipTo, priority)
            If (dtrs.Rows.Count = 0) Then
                Return 0
            Else
                Dim dtr = ToClass(dtrs.Rows(0))
                Return dtr.TransferLeadTime
            End If

        End Function

        ''' <summary>
        ''' 納期計算 (shipTo、priorityより 出荷ルートマスター 取得)
        ''' </summary>
        ''' <param name="shipTo"></param>
        ''' <param name="priority"></param>
        ''' <returns></returns>
        Public Function GetUsrShproutm(shipTo As String, priority As Integer) As DataTable
            Dim dt As New DataTable()
            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    dt = GetUsrShproutm(conn, tran, shipTo, priority)
                    tran.Commit()
                End Using
            End Using
            Return dt
        End Function
        ''' <summary>
        ''' 納期計算 (shipTo、priorityより 出荷ルートマスター 取得)
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="shipTo"></param>
        ''' <param name="priority"></param>
        ''' <returns></returns>
        Public Function GetUsrShproutm(conn As OracleConnection, tran As OracleTransaction, shipTo As String, priority As Integer) As DataTable

            Dim dt As New DataTable()
            'SECTM.FUPPSECT ≠ 'F'の場合（電子機器）	                     
            'SHPROUTM.FSHPTOCD（出荷ルートマスター.出荷先）	= 処理中のSHIP_TO（出荷先）	
            'かつ	
            'SHPROUTM.FPRIORITY（出荷ルートマスター.優先順位）	= 1	
            'SECTM.FUPPSECT = 'F'の場合（ハーネス）	
            'SHPROUTM.FSHPTOCD（出荷ルートマスター.出荷先） = 処理中のSHIP_TO（出荷先）
            'かつ
            'SHPROUTM.FPRIORITY（出荷ルートマスター.優先順位） = 2
            Try

                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}

                    ' FshpToCd string VARCHAR2(25 BYTE)
                    ' Fpriority Int16 NUMBER(3,0)
                    cmd.CommandText = "
                        SELECT 
                             FshpToCd  As ""ShipToCd"",  
                             Fpriority As ""Priority"",  
                             Ftranlt   As ""TransferLeadTimme""  
                        FROM Shproutm
                        WHERE FshpToCd  = :p_shipTo 
                        AND Fpriority   = :p_priority  "
                    cmd.Parameters.Add(":p_shipTo", OracleDbType.Varchar2, 45).Value = SafeVarchar(shipTo, 25)
                    cmd.Parameters.Add(":p_priority", OracleDbType.Int16).Value = priority
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim s = ex.Message
            End Try
            Return dt

        End Function

        ''' <summary>
        ''' DataRow to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As ShproutmRow

            Dim udir = New ShproutmRow
            udir.ShioToCd = dt.Field(Of String)("ShipToCd")
            udir.Priority = dt.Field(Of Int16)("Priority")
            udir.TransferLeadTime = dt.Field(Of Int16)("TransferLeadTimme")
            Return udir
        End Function

        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of ShproutmRow)

            Dim osrs = New List(Of ShproutmRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

    End Class
    ''' <summary>
    ''' 必要な項目のみ
    ''' </summary>
    Public Class ShproutmRow
        'SHPROUTM	FSHPTOCD	出荷先      VARCHAR2(25)
        'SHPROUTM	FPRIORITY	優先順位    NUMBER(3,0)
        'SHPROUTM	FTRANLT	    輸送L/T     NUMBER(3,0)

        Public Property ShioToCd As String
        Public Property Priority As Int16
        Public Property TransferLeadTime As Int16
    End Class
End Namespace