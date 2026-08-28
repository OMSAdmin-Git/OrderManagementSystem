Imports System.Configuration
Imports System.IO
Imports OMS.Business.Services
Imports OMS.Common
Imports OMS.Data

Namespace SUZUKI_SPIRITS_DATA_PRE_IMPORT

    Module Program

        Sub Main(args As String())
            Console.WriteLine("==============================================")
            Console.WriteLine(" SUZUKI_SPIRITS_DATA_PRE_IMPORT Console Runner ")
            Console.WriteLine("==============================================")

            Dim userId As String = "SUZUKI-AT"
            Dim isBatch As Boolean = True
            Dim targetFolder As String = String.Empty
            Dim overrideCustomerSettingId As Long = 0
            Dim profitCenter As String = String.Empty

            ' コマンドライン引数パース
            For i As Integer = 0 To args.Length - 1
                Dim arg = args(i).ToLower()
                If arg = "/batch" Then
                    isBatch = True
                ElseIf arg = "/manual" Then
                    isBatch = False
                ElseIf arg.StartsWith("/user:") Then
                    userId = args(i).Substring(6).Trim()
                ElseIf arg.StartsWith("/folder:") Then
                    targetFolder = args(i).Substring(8).Trim()
                ElseIf arg.StartsWith("/customersettingid:") Then
                    Long.TryParse(args(i).Substring(19).Trim(), overrideCustomerSettingId)
                ElseIf arg.StartsWith("/profitcenter:") Then
                    profitCenter = args(i).Substring(14).Trim()
                End If
            Next

            ' DBから設定値を取得
            Dim customerSettingId As Long = If(overrideCustomerSettingId > 0, overrideCustomerSettingId, GetSuzukiCustomerSettingId(profitCenter))
            Dim folderType As Integer = 4 ' 4: 混在

            ' フォルダ未指定の場合、DB またはローカルデフォルトから取得
            If String.IsNullOrWhiteSpace(targetFolder) Then
                targetFolder = GetSuzukiFolderPath(customerSettingId)
            End If

            Console.WriteLine($"[INFO] Mode: {If(isBatch, "BATCH (/batch)", "MANUAL (/manual)")}")
            Console.WriteLine($"[INFO] User ID: {userId}")
            Console.WriteLine($"[INFO] Target Folder: {targetFolder}")
            Console.WriteLine($"[INFO] Customer Setting ID: {customerSettingId}")

            If String.IsNullOrWhiteSpace(targetFolder) OrElse Not Directory.Exists(targetFolder) Then
                Console.WriteLine($"[WARN] Target folder does not exist or is not specified: '{targetFolder}'")
                Environment.ExitCode = 1
                Return
            End If

            Try
                Dim service As New SuzukiPreImportService()
                Dim count As Integer = service.ExecuteImport(targetFolder, customerSettingId, folderType, userId, isBatch)

                Console.WriteLine($"[SUCCESS] Processed {count} file(s) successfully.")
                Environment.ExitCode = 0
            Catch ex As Exception
                Console.WriteLine($"[ERROR] Import failed: {ex.Message}")
                Environment.ExitCode = 1
            End Try
        End Sub

        Private Function GetSuzukiCustomerSettingId(profitCenter As String) As Long
            Try
                Dim connStr = Utils.GetConnectionString()
                Dim sql As String
                If String.IsNullOrWhiteSpace(profitCenter) Then
                    sql = "SELECT customer_setting_id FROM customer_setting_mst WHERE customer_code = '5455' AND profit_center IS NULL AND customer_unit_id IS NULL AND active_flag = 'Y'"
                Else
                    sql = $"SELECT customer_setting_id FROM customer_setting_mst WHERE customer_code = '5455' AND profit_center = '{profitCenter}' AND customer_unit_id IS NULL AND active_flag = 'Y'"
                End If

                Using conn As New Oracle.ManagedDataAccess.Client.OracleConnection(connStr)
                    conn.Open()
                    Using cmd As New Oracle.ManagedDataAccess.Client.OracleCommand(sql, conn)
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                            Return Convert.ToInt64(result)
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try
            Return 0 ' Default fallback
        End Function

        Private Function GetSuzukiFolderPath(customerSettingId As Long) As String
            ' 1. DB (FOLDER_MST) から設定IDに基づくフォルダパスを取得 (フォルダ区分 = 4:混在)
            Try
                Dim connStr = Utils.GetConnectionString()
                Dim sql = "SELECT folder_path FROM folder_mst WHERE customer_setting_id = :p_id AND folder_type = 4 AND active_flag = 'Y'"
                Using conn As New Oracle.ManagedDataAccess.Client.OracleConnection(connStr)
                    conn.Open()
                    Using cmd As New Oracle.ManagedDataAccess.Client.OracleCommand(sql, conn)
                        cmd.Parameters.Add(":p_id", Oracle.ManagedDataAccess.Client.OracleDbType.Int64).Value = customerSettingId
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                            Dim pathVal = result.ToString()
                            If Not String.IsNullOrWhiteSpace(pathVal) AndAlso Directory.Exists(pathVal) Then
                                Return pathVal
                            End If
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try

            ' 2. App.config の SuzukiFolderPath を取得
            Dim configPath = ConfigurationManager.AppSettings("SuzukiFolderPath")
            If Not String.IsNullOrWhiteSpace(configPath) AndAlso Directory.Exists(configPath) Then
                Return configPath
            End If

            ' 3. ローカル相対パスのフォールバック (開発・テスト用)
            Dim samplePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "CSV_Example")
            If Directory.Exists(samplePath) Then
                Return Path.GetFullPath(samplePath)
            End If

            Return "C:\ASTI\DATA\SUZUKI\IN"
        End Function

    End Module

End Namespace
