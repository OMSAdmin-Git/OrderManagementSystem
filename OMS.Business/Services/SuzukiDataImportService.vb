Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Imports OMS.Common
Imports OMS.Data

Namespace Services

    Public Class SuzukiDataImportService

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        ''' <summary>
        ''' 受注登録 (Order Registration) Stage 2 Entry Point
        ''' </summary>
        Public Function ExecuteOrderRegistration(customerSettingId As Long, folderType As Integer, impRunId As Long, userId As String) As Integer
            Dim loggerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log")
            Dim logger As New Logger(loggerPath)
            logger.Write($"[SUZUKI_ORDER_IMPORT] Stage 2 Order Registration started for CustomerSettingId={customerSettingId}, FolderType={folderType}, ImpRunId={impRunId}")
            
            ' Placeholder for Stage 2 processing
            
            Return 0
        End Function

    End Class

End Namespace
