
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


        Public Function UpdateOrderBack() As String
            Return ""
        End Function

        Public Function UpdateOrderBackZero() As String
            Return ""
        End Function




    End Class
End Namespace