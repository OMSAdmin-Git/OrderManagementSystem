Imports System.Configuration
Imports System.IO
Imports OMS.Business.Services
Imports OMS.Common
Imports OMS.Data

Namespace SUZUKI_SPIRITS_DATA_IMPORT

    Module SuzukiSpiritsDataImport

        Sub Main(args As String())
            Dim logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs")
            If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
            Dim loggerPath = Path.Combine(logDir, $"SuzukiImport_{DateTime.Now:yyyyMMdd}.log")
            Dim logger As New Logger(loggerPath)

            Console.WriteLine("==============================================")
            Console.WriteLine(" SUZUKI SPIRITS データ取込バッチ処理 ")
            Console.WriteLine("==============================================")
            logger.Write("==============================================")
            logger.Write(" SUZUKI SPIRITS データ取込バッチ処理を開始します ")
            logger.Write("==============================================")

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

            Dim modeText = If(isBatch, "バッチ実行 (/batch)", "手動実行 (/manual)")
            Console.WriteLine($"[INFO] 動作モード: {modeText}")
            Console.WriteLine($"[INFO] 実行ユーザーID: {userId}")
            Console.WriteLine($"[INFO] 取込対象フォルダ: {targetFolder}")
            Console.WriteLine($"[INFO] 取引先設定ID: {customerSettingId}")

            logger.Write($"[INFO] 動作モード: {modeText}, 実行ユーザーID: {userId}, 取込対象フォルダ: {targetFolder}, 取引先設定ID: {customerSettingId}")

            Try
                Dim preCount As Integer = 0

                ' 1. Stage 1 (Pre-Import): フォルダが存在する場合はファイル取込を実行
                If Not String.IsNullOrWhiteSpace(targetFolder) AndAlso Directory.Exists(targetFolder) Then
                    Console.WriteLine("[INFO] 工程 1: 取込準備 (Pre-Import) を開始します...")
                    Dim preService As New SuzukiPreImportService()
                    preCount = preService.ExecuteImport(targetFolder, customerSettingId, folderType, userId, isBatch)
                    Console.WriteLine($"[INFO] 取込準備 (Pre-Import) 完了: {preCount} 件のファイルを正常に処理しました。")
                Else
                    Console.WriteLine($"[INFO] 取込対象フォルダが未指定、または存在しません: '{targetFolder}'。取込準備 (Pre-Import) をスキップします。")
                End If

                ' 2. Stage 2 (Order Registration & Reconciliation) -> 受注取込 (Import)
                ' バッチモードで Pre-Import が実行された場合 (preCount > 0)、ExecuteImport 内で既に
                ' 当該 impRunId を対象に ExecuteOrderRegistration が実行されているため、重複実行を防止する。
                ' 未処理ファイルが 0 件だがステージングに未登録データが存在する場合のみ、明示的に登録を実行する。
                If Not isBatch OrElse preCount = 0 Then
                    Console.WriteLine("[INFO] 工程 2: 受注取込・本登録 (Import) および STRA納期設定を開始します...")
                    Dim importService As New SuzukiDataImportService()
                    Dim registeredCount As Integer = importService.ExecuteOrderRegistration(customerSettingId, folderType, 0, userId, isBatch)
                    Console.WriteLine($"[SUCCESS] 受注取込 (Import) が正常に完了しました。登録件数: {registeredCount} 件")
                Else
                    Console.WriteLine("[INFO] 工程 2: バッチ実行フローに基づき、取込準備 (Pre-Import) から受注取込 (Import)・STRA納期設定まで自動実行されました。")
                End If

                Console.WriteLine("[SUCCESS] 全てのバッチ処理工程 (Pre-Import / Import / STRA納期設定) が正常に完了しました。")
                logger.Write("[SUCCESS] 全てのバッチ処理工程 (Pre-Import / Import / STRA納期設定) が正常に完了しました。")
                Environment.ExitCode = 0
            Catch ex As Exception
                Console.WriteLine($"[ERROR] 取込処理に失敗しました: {ex.Message}")
                logger.Write($"[ERROR] 取込処理に失敗しました: {ex.Message}")
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
