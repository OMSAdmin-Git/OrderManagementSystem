Imports System.Data
Imports System.Text
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