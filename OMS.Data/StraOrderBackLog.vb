
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
Imports Microsoft.VisualBasic.ApplicationServices
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client
' <summary>
' 受注残 関連
' </summary>
Namespace OMS.Data
    Public Class StraOrderBackLog

#Region "フィールド・コンストラクタ"
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub
#End Region

        ''' <summary>
        ''' UPDATE(受注残更新)
        ''' </summary>
        ''' <returns></returns>
        Public Function UpdateOrderBack(ud As Date, uid As String, pid As String) As String

            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("MERGE INTO ORDERS ODR ")
            sb.AppendLine("USING  ")
            sb.AppendLine("(SELECT ")
            sb.AppendLine("    FCUSREQQTY, ")
            sb.AppendLine("    FSHPQTY, ")
            sb.AppendLine("    FODRREMQTY ")
            sb.AppendLine(" FROM ")
            sb.AppendLine("    A_S_ODRREMQTY ")
            sb.AppendLine(" WHERE  ")
            sb.AppendLine("    FIMPPATN = 1 ")
            sb.AppendLine(") ASO ")
            sb.AppendLine("ON( ")
            sb.AppendLine("    ODR.ACTIVE_FLAG = 'Y'                      AND ")
            sb.AppendLine("    ODR.ORDER_TYPE  = 3                        AND ")
            sb.AppendLine("    ODR.STATUS = 'EXPORTED'                    AND ")
            sb.AppendLine("    ODR.CUSTOMER_CODE = ASO.FCUSTCD            AND ")
            sb.AppendLine("    ODR.CUSTOMER_ORDER_NO = ASO.FCUSTODRNO     AND ")
            sb.AppendLine("    ODR.CUSTOMER_ORDER_LINE_NO = ASO.FCUSTLINE AND ")
            sb.AppendLine("    ODR.PRE_DAILY_DELIVERY_DATE = ASO.FUSRDTE3 AND ")
            sb.AppendLine("    ODR.ITEM_NO = ASO.FITEMNO ")
            sb.AppendLine(") ")
            sb.AppendLine("WHEN MATCHED THEN  ")
            sb.AppendLine("UPDATE SET ")
            sb.AppendLine("    ODR.STRA_ORDER_QTY = ASO.FCUSREQQTY, ")
            sb.AppendLine("    ODR.STRA_SHIP_QTY = ASO.FSHPQTY, ")
            sb.AppendLine("    ODR.STRA_ORDER_BACKLOG = ASO.FODRREMQTY, ")
            sb.AppendLine($"    ODR.UPDATED_AT = {ud}, ")
            sb.AppendLine($"    ODR.UPDATED_USER_ID = {uid}, ")
            sb.AppendLine($"    ODR.UPDATED_PG_ID = {pid} ")

            Try
                Using conn As New OracleConnection(_connectionString)
                    Using cmd As New OracleCommand(sb.ToString(), conn)
                        conn.Open()
                        Using tran As OracleTransaction = conn.BeginTransaction()
                            Dim cnt = cmd.ExecuteNonQuery()
                            tran.Commit()
                        End Using
                        conn.Close()
                    End Using
                End Using

            Catch ex As Exception
                errors = ex.Message
            End Try

            Return errors

        End Function

        ''' <summary>
        ''' UPDATE(受注残更新)
        ''' </summary>
        ''' <returns></returns>
        Public Function UpdateOrderBackZero(ud As Date, uid As String, pid As String) As String

            Dim sb As New StringBuilder()
            Dim errors = ""

            sb.AppendLine("UPDATE ORDERS ODR ")
            sb.AppendLine("SET ODR.STRA_ORDER_BACKLOG = 0, ")
            sb.AppendLine("    ODR.STATUS = 'SHIPPED' ")
            sb.AppendLine($"    ODR.UPDATED_AT = {ud.ToString()}, ")
            sb.AppendLine($"    ODR.UPDATED_USER_ID = {uid}, ")
            sb.AppendLine($"    ODR.UPDATED_PG_ID = {pid} ")
            sb.AppendLine("WHERE ODR.ACTIVE_FLAG = 'Y' ")
            sb.AppendLine("AND  ODR.ORDER_TYPE = 3 ")
            sb.AppendLine("AND  ODR.STATUS = 'EXPORTED' ")
            sb.AppendLine("AND NOT EXISTS ( ")
            sb.AppendLine("    SELECT 1 FROM A_S_ODRREMQTY ASO ")
            sb.AppendLine("    WHERE ASO.FIMPPATN = 1 ")
            sb.AppendLine("      AND ASO.FCUSTCD = ODR.CUSTOMER_CODE ")
            sb.AppendLine("      AND ASO.FCUSTODRNO = ODR.CUSTOMER_ORDER_NO ")
            sb.AppendLine("      AND ASO.FCUSTLINE = ODR.CUSTOMER_ORDER_LINE_NO ")
            sb.AppendLine("      AND ASO.FUSRDTE3 = ODR.PRE_DAILY_DELIVERY_DATE ")
            sb.AppendLine("      AND ASO.FITEMNO = ODR.ITEM_NO	")
            sb.AppendLine("    ) ")
            Try
                Using conn As New OracleConnection(_connectionString)
                    Using cmd As New OracleCommand(sb.ToString(), conn)
                        conn.Open()
                        Using tran As OracleTransaction = conn.BeginTransaction()
                            Dim cnt = cmd.ExecuteNonQuery()
                            tran.Commit()
                        End Using
                        conn.Close()
                    End Using
                End Using

            Catch ex As Exception
                errors = ex.Message
            End Try

            Return errors
        End Function

    End Class
End Namespace