
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data.UsrDeffIdfRepository
Imports Oracle.ManagedDataAccess.Client
Imports DataTable = System.Data.DataTable

Namespace OMS.Data
    Public Class SectmRepository
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
        ''' <summary>
        ''' 納期計算 (customerCodeより 出荷ルートマスター.上位部署 取得)
        ''' </summary>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetUpSection(customerCode As String) As String
            Dim dtrs = GetSectm(customerCode)
            If (dtrs.Rows.Count <> 1) Then
                Return 0
            Else
                Dim dtr = ToClass(dtrs.Rows(0))
                Return dtr.UpSection
            End If

        End Function
        ''' <summary>
        ''' 納期計算 (customerCodeより 出荷ルートマスター.上位部署 取得)
        ''' </summary>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetSectm(customerCode As String) As DataTable
            Dim dt As New DataTable()
            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    dt = GetSectm(conn, tran, customerCode)
                    tran.Commit()
                End Using
            End Using
            Return dt
        End Function

        ''' <summary>
        ''' 納期計算 (customerCodeより 出荷ルートマスター 取得)
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="customerCode"></param>
        ''' <returns></returns>
        Public Function GetSectm(conn As OracleConnection, tran As OracleTransaction, customerCode As String) As DataTable

            Dim dt As New DataTable()
            Try

                Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                    cmd.CommandText = "
                        SELECT 
                             fuppsect As ""AppSection""  
                        FROM sectm
                        WHERE fsectcd = :p_customerCode "
                    cmd.Parameters.Add(":p_customerCode", OracleDbType.Varchar2, 45).Value = SafeVarchar(customerCode, 45)
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
        Public Function ToClass(dt As DataRow) As Sectm

            Dim smr = New Sectm
            smr.UpSection = dt.Field(Of String)("AppSection")
            Return smr
        End Function
        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of Sectm)

            Dim osrs = New List(Of Sectm)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

        ''' <summary>
        ''' 必要な項目のみ
        ''' </summary>
        Public Class Sectm
            Public Property UpSection As String     'FUPPSECT 上位部署 VARCHAR2(25)
        End Class

    End Class
End Namespace
