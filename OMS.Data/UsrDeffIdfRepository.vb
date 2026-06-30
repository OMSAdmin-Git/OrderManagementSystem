Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
Imports DataTable = System.Data.DataTable

Namespace OMS.Data
    Public Class UsrDeffIdfRepository
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
        ''' <summary>
        ''' 品揃えリードタイムy取得
        ''' </summary>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetAssortLeadTime(customerCode As String, profitCenter As String, customerItemNumber As String) As Decimal
            Dim dtrs = GetUsrDeffIdf(customerCode, profitCenter, customerItemNumber)
            If (dtrs.Rows.Count = 0) Then
                Return 0
            Else
                Dim dtr = ToClass(dtrs.Rows(0))
                Return dtr.AssortLeadTime
            End If

        End Function

        ''' <summary>
        ''' 納期計算 (customerCodeより 品目マスタ.(A)品揃リードタイム 取得)
        ''' </summary>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetUsrDeffIdf(customerCode As String, profitCenter As String, customerItemNumber As String) As DataTable
            Dim dt As New DataTable()
            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    dt = GetUsrDeffIdf(conn, tran, customerCode, profitCenter, customerItemNumber)
                    tran.Commit()
                End Using
                conn.Close()
            End Using
            Return dt
        End Function
        ''' <summary>
        ''' 納期計算 (customerCodeより 品目マスタ.(A)品揃リードタイム 取得)
        ''' </summary>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetUsrDeffIdf(conn As OracleConnection, tran As OracleTransaction, customerCode As String, profitCenter As String, customerItemNumber As String) As DataTable

            Dim dt As New DataTable()
            'USRDEFFLDF.FUSRSTR10 (品目マスタ.(A)顧客) = 処理中のCUSTOMER_CODE
            'かつ
            'USRDEFFLDF.FUSRSTR1 (品目マスタ.(A)PC) = 処理中のPROFIT_CENTER
            'かつ
            'USRDEFFLDF.FUSRSTR18（ユーザー定義マスタ.客先品番）= 処理中のCUSTOMER_ITEM_NO
            Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                cmd.CommandText = "
                        SELECT 
                             fusrdec1 As ""AssortLeadTime""  
                        FROM usrdeffldf
                        WHERE fusrstr10 = :p_customerCode 
                        AND fusrstr1    = :p_profitCenter 
                        AND fusrstr18   = :p_customerItemNumber "
                cmd.Parameters.Add(":p_customerCode", OracleDbType.Varchar2, 45).Value = SafeVarchar(customerCode, 45)
                cmd.Parameters.Add(":p_profitCenter", OracleDbType.Varchar2, 45).Value = SafeVarchar(profitCenter, 45)
                cmd.Parameters.Add(":p_customerItemNumber", OracleDbType.Varchar2, 45).Value = SafeVarchar(customerItemNumber, 45)
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
            Return dt

        End Function

        ''' <summary>
        ''' DataRow to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As UsrDeffIdfRow

            Dim udir = New UsrDeffIdfRow
            udir.AssortLeadTime = dt.Field(Of Decimal)("AssortLeadTime")
            Return udir
        End Function

        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of UsrDeffIdfRow)

            Dim osrs = New List(Of UsrDeffIdfRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

        ''' <summary>
        ''' 必要な項目のみ
        ''' </summary>
        Public Class UsrDeffIdfRow
            Public Property AssortLeadTime As Decimal     '品目マスタ.(A)品揃リードタイム NUMBER(18,6)
        End Class
    End Class
End Namespace