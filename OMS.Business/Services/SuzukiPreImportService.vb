Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Text
Imports OMS.Common
Imports OMS.Data
Imports OMS.Data.SUZUKI
Imports Oracle.ManagedDataAccess.Client

Namespace Services

    ''' <summary>
    ''' スズキ (SPIRITS) データ取込前処理 サービスエンジン
    ''' </summary>
    Public Class SuzukiPreImportService

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        Public Sub New()
            Me.New(Utils.GetConnectionString())
        End Sub

        ''' <summary>
        ''' 取込処理メインエントリ
        ''' </summary>
        Public Function ExecuteImport(folderPath As String, customerSettingId As Long, folderType As Integer, Optional userId As String = "SUZUKI-AT", Optional isBatch As Boolean = True, Optional reconcileFlag As String = "", Optional fcstReconcileFlag As String = "", Optional ByRef webErrors As List(Of String) = Nothing, Optional ByRef outStageRows As List(Of ImpFilesStageRow) = Nothing, Optional ByRef totalFilesCount As Integer = 0, Optional ByRef validFilesCount As Integer = 0, Optional ByRef errorFilesCount As Integer = 0) As Integer
            Dim loggerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log")
            Dim logger As New Logger(loggerPath)

            ' ユーザーIDの厳密チェック (Strict User ID check)
            If String.IsNullOrWhiteSpace(userId) Then
                If isBatch Then
                    userId = "SUZUKI-AT"
                Else
                    logger.Write("[SUZUKI_IMPORT] User ID is empty. Aborting execution.")
                    If webErrors IsNot Nothing Then
                        webErrors.Add("ユーザーIDが取得できませんでした。ログイン状態を確認してください。")
                    End If
                    Return 0
                End If
            End If

            logger.Write($"[SUZUKI_IMPORT] Start processing folder={folderPath} user={userId} isBatch={isBatch}")

            If String.IsNullOrWhiteSpace(folderPath) OrElse Not Directory.Exists(folderPath) Then
                logger.Write($"[SUZUKI_IMPORT] Target folder does not exist: {folderPath}")
                Return 0
            End If

            ' 2. ユーザー専用のWORKフォルダ及びCOMPLETEDフォルダを確保
            Dim workRoot = ConfigurationManager.AppSettings("WorkFolderRoot")
            If String.IsNullOrWhiteSpace(workRoot) Then workRoot = "C:\ASTI\DATA\WORK"
            Dim completedRoot = ConfigurationManager.AppSettings("CompletedFolderRoot")
            If String.IsNullOrWhiteSpace(completedRoot) Then completedRoot = "C:\ASTI\DATA\COMPLETED"

            Dim customerCode = GetCustomerCode(customerSettingId)
            Dim userWorkDir = Path.Combine(workRoot, userId, customerCode, folderType.ToString())
            Utils.EnsureDirectory(userWorkDir)
            Dim userCompletedDir = Path.Combine(completedRoot, userId, customerCode, folderType.ToString())
            Utils.EnsureDirectory(userCompletedDir)

            ' 1. 取込対象フォルダ（source folder）内の CSV / Excel / TXT を検索（サブフォルダは除外）
            Dim sourceFiles = If(Directory.Exists(folderPath),
                Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly) _
                    .Where(Function(f)
                        Dim ext = Path.GetExtension(f).ToLowerInvariant()
                        Return ext = ".csv" OrElse ext = ".xlsx" OrElse ext = ".txt"
                    End Function).ToArray(),
                Array.Empty(Of String)())

            If sourceFiles.Length = 0 Then
                logger.Write("[SUZUKI_IMPORT] No import target files (.csv, .xlsx, .txt) found in source folder.")
                totalFilesCount = 0
                validFilesCount = 0
                errorFilesCount = 0
                Return 0
            End If

            ' 2. 今回発見されたファイルのみ WORK へ移動し、処理対象リストに格納
            Dim workFiles As New List(Of String)()
            For Each src In sourceFiles
                Dim fName = Path.GetFileName(src)
                Dim dest = Path.Combine(userWorkDir, fName)
                If Not src.Equals(dest, StringComparison.OrdinalIgnoreCase) Then
                    If File.Exists(dest) Then File.Delete(dest)
                    File.Move(src, dest)
                End If
                workFiles.Add(dest)
            Next

            totalFilesCount = workFiles.Count
            validFilesCount = 0
            errorFilesCount = 0

            Dim processCount As Integer = 0
            Dim runRepo As New ImpRunRepository(_connectionString)
            Dim stageRepo As New ImpFilesStageRepository(_connectionString)
            Dim impRunId As Long = 0

            ' DBトランザクション実行
            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                
                ' Step 4: Import Run Tracking Initialization
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Try
                        Dim runRow As New ImpRunRow(DateTime.Now, "RUNNING", userId, "OrderImport(Execute)")
                        runRepo.Insert(conn, tran, runRow)
                        If runRow.ImpRunId.HasValue Then
                            impRunId = runRow.ImpRunId.Value
                        End If
                        tran.Commit()
                    Catch ex As Exception
                        tran.Rollback()
                        logger.Write($"[SUZUKI_IMPORT] Failed to insert IMP_RUN: {ex.Message}")
                        Return 0
                    End Try
                End Using

                For Each workFilePath In workFiles
                    Try
                        Dim fileName = Path.GetFileName(workFilePath)
                        Dim fileExt = Path.GetExtension(workFilePath).ToLowerInvariant()

                        ' --- ASTI追加内示 / Excel ファイル (.xlsx) のステージング処理 ---
                        If fileExt = ".xlsx" Then
                            Using tran As OracleTransaction = conn.BeginTransaction()
                                Try
                                    Dim rFlag = If(Not String.IsNullOrEmpty(reconcileFlag), reconcileFlag, If(isBatch, "Y", "N"))
                                    Dim fFlag = If(Not String.IsNullOrEmpty(fcstReconcileFlag), fcstReconcileFlag, If(isBatch, "Y", "N"))

                                    Dim stageRow As New ImpFilesStageRow() With {
                                        .CustomerSettingId = customerSettingId,
                                        .FolderType = CShort(folderType),
                                        .FolderPath = folderPath,
                                        .FileName = fileName,
                                        .StagedFolderPath = userWorkDir,
                                        .StagedFileName = fileName,
                                        .ReconcileFlag = rFlag,
                                        .FcstReconcileFlag = fFlag,
                                        .HandFlag = "Y",
                                        .Status = "DISCOVERED",
                                        .CreatedAt = DateTime.Now,
                                        .CreatedUserId = userId,
                                        .CreatedPgId = "OrderImport(Stage)",
                                        .UpdatedAt = DateTime.Now,
                                        .UpdatedUserId = userId,
                                        .UpdatedPgId = "OrderImport(Stage)"
                                    }
                                    Dim stageInsErr = stageRepo.Insert(conn, tran, stageRow)
                                    If Not String.IsNullOrEmpty(stageInsErr) Then
                                        Throw New Exception($"取込ファイルワーク登録エラー: {stageInsErr}")
                                    End If
                                    tran.Commit()

                                    validFilesCount += 1
                                    processCount += 1
                                    If outStageRows IsNot Nothing Then
                                        outStageRows.Add(stageRow)
                                    End If
                                    logger.Write($"[SUZUKI_IMPORT] Successfully staged Excel file (HAND_FLAG=Y): {fileName}")
                                Catch ex As Exception
                                    tran.Rollback()
                                    errorFilesCount += 1
                                    logger.Write($"[SUZUKI_IMPORT] Error staging Excel file {fileName}: {ex.Message}")
                                End Try
                            End Using
                            Continue For
                        End If

                        ' --- CSV / TXT ファイル処理 (SPIRITS EDI / 取込処理) ---
                        ' 種別判定
                        Dim diagMsg As String = ""
                        Dim infoCode = SuzukiCsvParser.PeekInfoTypeCode(workFilePath, diagMsg)
                        If String.IsNullOrEmpty(infoCode) OrElse infoCode.Length <> 4 Then
                            errorFilesCount += 1
                            logger.Write($"[SUZUKI_IMPORT] Invalid or missing INFO_TYPE_CODE in file: {fileName}. Detail: {diagMsg}")

                            Dim userErrMsg = $"【{fileName}】対象データが1件もありません。"
                            If webErrors IsNot Nothing Then
                                webErrors.Add(userErrMsg)
                            End If

                            Dim errList As New List(Of String)()
                            errList.Add("対象データが1件もありません。")
                            If Not String.IsNullOrEmpty(diagMsg) Then
                                errList.Add(diagMsg)
                            End If
                            ExportErrorsToCsv(folderPath, fileName, errList, isBatch)

                            MoveBackErrorFile(workFilePath, folderPath, fileName, userId)
                            Continue For
                        End If

                        Dim impFileStageId As Long = 0

                        Using tran As OracleTransaction = conn.BeginTransaction()
                            Try
                                Dim rFlag = If(Not String.IsNullOrEmpty(reconcileFlag), reconcileFlag, If(isBatch, "Y", "N"))
                                Dim fFlag = If(Not String.IsNullOrEmpty(fcstReconcileFlag), fcstReconcileFlag, If(isBatch, "Y", "N"))

                                ' Step 5: Work File Staging Registration
                                Dim stageRow As New ImpFilesStageRow() With {
                                    .CustomerSettingId = customerSettingId,
                                    .FolderType = CShort(folderType),
                                    .FolderPath = folderPath,
                                    .FileName = fileName,
                                    .StagedFolderPath = userWorkDir,
                                    .StagedFileName = fileName,
                                    .ReconcileFlag = rFlag,
                                    .FcstReconcileFlag = fFlag,
                                    .HandFlag = "N",
                                    .Status = "DISCOVERED",
                                    .CreatedAt = DateTime.Now,
                                    .CreatedUserId = userId,
                                    .CreatedPgId = "OrderImport(Stage)",
                                    .UpdatedAt = DateTime.Now,
                                    .UpdatedUserId = userId,
                                    .UpdatedPgId = "OrderImport(Stage)"
                                }
                                Dim stageInsErr = stageRepo.Insert(conn, tran, stageRow)
                                If Not String.IsNullOrEmpty(stageInsErr) Then
                                    Throw New Exception($"取込ファイルワーク登録エラー: {stageInsErr}")
                                End If
                                
                                ' Fetch the generated ID
                                Dim fetchedRow = stageRepo.GetImpFilesStageFilename(conn, tran, fileName, folderPath)
                                If fetchedRow IsNot Nothing AndAlso fetchedRow.ImpFileStageId > 0 Then
                                    impFileStageId = fetchedRow.ImpFileStageId
                                End If

                                ' パース & インサート
                                ProcessFileByInfoCode(conn, tran, workFilePath, infoCode, userId, impRunId, impFileStageId)

                                ' 品番更新 (PRDSLSODRM)
                                UpdateAstiPartNumbers(conn, tran, infoCode, impFileStageId)

                                ' 無効化チェック (ACTIVE_FLAG = 'N')
                                UpdateActiveFlags(conn, tran, infoCode)

                                ' エラーチェック (ASTI品番未設定行)
                                Dim errorRows = FindPartMatchingErrors(conn, tran, infoCode, workFilePath, impFileStageId)
                                If errorRows.Count > 0 Then
                                    errorFilesCount += 1

                                    ' 有効件数の確認 (Check valid rows count)
                                    Dim validRowsCount As Integer = 0
                                    Dim tableName = GetTargetTableName(infoCode)
                                    If Not String.IsNullOrEmpty(tableName) Then
                                        Dim sqlCount = $"SELECT COUNT(1) FROM {tableName} WHERE imp_file_id = :p_imp_file_id AND active_flag = 'Y' AND item_no IS NOT NULL"
                                        Using countCmd As New OracleCommand(sqlCount, conn)
                                            countCmd.Transaction = tran
                                            countCmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileStageId
                                            Dim cnt = countCmd.ExecuteScalar()
                                            If cnt IsNot Nothing AndAlso Not DBNull.Value.Equals(cnt) Then
                                                validRowsCount = Convert.ToInt32(cnt)
                                            End If
                                        End Using
                                    End If

                                    logger.Write($"[SUZUKI_IMPORT] Found {errorRows.Count} part matching errors in {fileName} (valid rows: {validRowsCount})")

                                    If validRowsCount = 0 Then
                                        ' --- Rule 4: 全件がASTI品番エラーの場合 ---
                                        ' エラーリストの最後に「対象データが1件もありません。」を追加
                                        errorRows.Add("対象データが1件もありません。")
                                        ExportErrorsToCsv(folderPath, fileName, errorRows, isBatch, webErrors)

                                        ' workに移動したファイルは元のフォルダへ戻す (※workフォルダには残らない)
                                        MoveBackErrorFile(workFilePath, folderPath, fileName, userId)

                                        ' スズキ特殊テーブルの該当データは削除 (ロールバック)
                                        tran.Rollback()
                                        Continue For
                                    Else
                                        ' --- Rule 3: 一部の行がASTI品番エラーの場合 ---
                                        ' 1. エラーリストへ追加 (CSV出力 & 画面エラーリスト)
                                        ExportErrorsToCsv(folderPath, fileName, errorRows, isBatch, webErrors)

                                        ' 2. 取込元フォルダにファイル名にuseridとtimestampをつけてコピー (※コピーなのでworkフォルダに残る)
                                        CopyErrorBackupFile(workFilePath, folderPath, fileName, userId)

                                        ' 3. スズキ特殊テーブルの該当データは削除する
                                        DeleteErrorRows(conn, tran, infoCode, impFileStageId)
                                    End If
                                Else
                                    ' --- Rule 2: 全件正常の場合 ---
                                    logger.Write($"[SUZUKI_IMPORT] File staged successfully in WORK: {fileName}")
                                End If

                                ' 正常データおよびIMP_FILES_STAGEを確定
                                tran.Commit()
                                validFilesCount += 1
                                processCount += 1

                                If outStageRows IsNot Nothing Then
                                    outStageRows.Add(stageRow)
                                End If
                            Catch ex As Exception
                                tran.Rollback()
                                errorFilesCount += 1
                                logger.Write($"[SUZUKI_IMPORT] DB Error processing {fileName}: {ex.Message}")
                                Dim errList As New List(Of String)()
                                errList.Add($"処理エラー: {ex.Message}")
                                ExportErrorsToCsv(folderPath, fileName, errList, isBatch, webErrors)
                                MoveBackErrorFile(workFilePath, folderPath, fileName, userId)
                            End Try
                        End Using
                    Catch ex As Exception
                        logger.Write($"[SUZUKI_IMPORT] File Error: {ex.Message}")
                    End Try
                Next

                ' Step 10: Transition to Order Registration (Batch Mode Only)
                If isBatch Then
                    Dim importService As New SuzukiDataImportService(_connectionString)
                    importService.ExecuteOrderRegistration(customerSettingId, folderType, impRunId, userId)
                End If

                ' 実行管理更新
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Try
                        runRepo.UpdateRange(impRunId, "COMPLETED", DateTime.Now, processCount, 0, 0, "")
                        tran.Commit()
                    Catch ex As Exception
                        tran.Rollback()
                    End Try
                End Using

            End Using

            Return processCount
        End Function

        Private Function ProcessFileByInfoCode(conn As OracleConnection, tran As OracleTransaction, filePath As String, infoCode As String, userId As String, impRunId As Long, impFileStageId As Long) As Integer
            Dim insertedCount As Integer = 0
            Select Case infoCode
                Case "0501", "0502"
                    Dim repo As New Spirits0501And0502Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0501And0502(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0600", "0630"
                    Dim repo As New Spirits0600And0630Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.Parse0600And0630(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0602"
                    Dim repo As New Spirits0602Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.Parse0602(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0650"
                    Dim repo As New Spirits0650Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0650(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0651"
                    Dim repo As New Spirits0651Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0651(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0740"
                    Dim repo As New Spirits0740Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0740(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0813"
                    Dim repo As New Spirits0813Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0813(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "0814"
                    Dim repo As New Spirits0814Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0814(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "6604", "6634"
                    Dim repo As New Spirits6604And6634Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits6604And6634(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "6624"
                    Dim repo As New Spirits6624Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits6624(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count

                Case "663N", "663S", "664T"
                    Dim repo As New Spirits663NAnd663SAnd66Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits663NAnd663SAnd664T(filePath)
                    If rows.Count = 0 Then Throw New Exception("対象データが1件もありません。")
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Stage)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Stage)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)
                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
                    insertedCount = rows.Count
            End Select
            Return insertedCount
        End Function

        Private Function GetTargetTableName(infoCode As String) As String
            Select Case infoCode
                Case "0600", "0630" : Return "SUZUKI_SPIRITS_0600AND0630"
                Case "0602" : Return "SUZUKI_SPIRITS_0602"
                Case "0501", "0502" : Return "SUZUKI_SPIRITS_0501AND0502"
                Case "0650" : Return "SUZUKI_SPIRITS_0650"
                Case "0651" : Return "SUZUKI_SPIRITS_0651"
                Case "0740" : Return "SUZUKI_SPIRITS_0740"
                Case "0813" : Return "SUZUKI_SPIRITS_0813"
                Case "0814" : Return "SUZUKI_SPIRITS_0814"
                Case "6604", "6634" : Return "SUZUKI_SPIRITS_6604AND6634"
                Case "6624" : Return "SUZUKI_SPIRITS_6624"
                Case "663N", "663S", "664T" : Return "SUZUKI_SPIRITS_663NAND663SAND664T"
                Case Else : Return String.Empty
            End Select
        End Function

        Public Shared Function GetCustomerItemNoColumnIndex(infoCode As String) As Integer
            Select Case infoCode
                Case "0501", "0502"
                    Return 10
                Case "0600", "0630"
                    Return 10
                Case "0602"
                    Return 8
                Case "0650"
                    Return 8
                Case "0651"
                    Return 8
                Case "0740"
                    Return 6
                Case "0813"
                    Return 9
                Case "0814"
                    Return 9
                Case "6604", "6634"
                    Return 9
                Case "6624"
                    Return 7
                Case "663N", "663S", "664T"
                    Return 11
                Case Else
                    Return 10
            End Select
        End Function

        Private Sub UpdateAstiPartNumbers(conn As OracleConnection, tran As OracleTransaction, infoCode As String, impFileStageId As Long)
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return

            Dim isZSuffix As Boolean = {"6604", "6624", "6634", "663N", "663S", "664T"}.Contains(infoCode)

            Dim sql As New StringBuilder()
            sql.AppendLine($"UPDATE {tableName} t")
            sql.AppendLine("SET t.item_no = (")
            sql.AppendLine("  SELECT TRIM(p.FPRDCD) FROM PRDSLSODRM p")
            sql.AppendLine("  WHERE TRIM(p.FCUSTCD) = '5455'")
            sql.AppendLine("    AND TRIM(p.FCUSTITEMNO) = TRIM(t.customer_item_no)")

            If isZSuffix Then
                sql.AppendLine("    AND UPPER(TRIM(p.FPRDCD)) LIKE '%Z'")
            Else
                sql.AppendLine("    AND UPPER(TRIM(p.FPRDCD)) NOT LIKE '%Z'")
            End If

            sql.AppendLine("    FETCH FIRST 1 ROWS ONLY")
            sql.AppendLine(")")
            sql.AppendLine("WHERE t.active_flag = 'Y' AND t.item_no IS NULL")
            If impFileStageId > 0 Then
                sql.AppendLine("  AND t.imp_file_id = :p_imp_file_id")
            End If

            Using cmd As New OracleCommand(sql.ToString(), conn)
                cmd.Transaction = tran
                If impFileStageId > 0 Then
                    cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileStageId
                End If
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub UpdateActiveFlags(conn As OracleConnection, tran As OracleTransaction, infoCode As String)
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return

            Dim sql As New StringBuilder()
            sql.AppendLine($"UPDATE {tableName} old_t")
            sql.AppendLine("SET old_t.active_flag = 'N'")
            sql.AppendLine("WHERE old_t.active_flag = 'Y'")
            sql.AppendLine($"  AND EXISTS (")
            sql.AppendLine($"    SELECT 1 FROM {tableName} new_t")
            sql.AppendLine($"    WHERE new_t.customer_item_no = old_t.customer_item_no")

            If infoCode = "0740" Then
                sql.AppendLine("      AND new_t.acceptance_date = old_t.acceptance_date")
                sql.AppendLine("      AND new_t.acceptance_time = old_t.acceptance_time")
            ElseIf {"6604", "6624", "6634", "663N", "663S", "664T"}.Contains(infoCode) Then
                sql.AppendLine("      AND new_t.publication_date = old_t.publication_date")
            Else
                sql.AppendLine("      AND new_t.publication_date = old_t.publication_date")
                sql.AppendLine("      AND new_t.publication_time = old_t.publication_time")
            End If

            sql.AppendLine($"      AND new_t.created_at > old_t.created_at")
            sql.AppendLine($"  )")

            Using cmd As New OracleCommand(sql.ToString(), conn)
                cmd.Transaction = tran
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Function FindPartMatchingErrors(conn As OracleConnection, tran As OracleTransaction, infoCode As String, csvFilePath As String, impFileStageId As Long) As List(Of String)
            Dim result As New List(Of String)()
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return result

            Dim unmatchedItems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim sql = $"SELECT DISTINCT customer_item_no FROM {tableName} WHERE active_flag = 'Y' AND item_no IS NULL"
            If impFileStageId > 0 Then
                sql &= " AND imp_file_id = :p_imp_file_id"
            End If

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                If impFileStageId > 0 Then
                    cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileStageId
                End If
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then
                            unmatchedItems.Add(reader.GetString(0).Trim())
                        End If
                    End While
                End Using
            End Using

            If unmatchedItems.Count > 0 AndAlso File.Exists(csvFilePath) Then
                Dim colIdx = GetCustomerItemNoColumnIndex(infoCode)
                Dim lines = SuzukiCsvParser.ReadLinesAutoEncoding(csvFilePath)

                For lineIdx As Integer = 0 To lines.Length - 1
                    Dim line = lines(lineIdx)
                    If String.IsNullOrWhiteSpace(line) Then Continue For

                    Dim cols = SuzukiCsvParser.SplitCsvLine(line)
                    Dim custItemNo = SuzukiCsvParser.CleanCol(cols, colIdx)
                    If Not String.IsNullOrEmpty(custItemNo) AndAlso unmatchedItems.Contains(custItemNo) Then
                        Dim csvRowNumber = lineIdx + 1
                        result.Add($"Row({csvRowNumber}): 品目No及び製品コードが取得できません。")
                    End If
                Next
            End If

            Return result
        End Function

        Private Sub ExportErrorsToCsv(sourceFolder As String, originalFileName As String, errors As List(Of String), isBatch As Boolean, Optional ByRef webErrors As List(Of String) = Nothing)
            Try
                ' Add to webErrors list if passed from Web UI caller
                If webErrors IsNot Nothing AndAlso errors IsNot Nothing Then
                    For Each errItem In errors
                        webErrors.Add($"{originalFileName}: {errItem}")
                    Next
                End If

                Dim errorDir = Path.Combine(sourceFolder, "エラーリスト")
                Utils.EnsureDirectory(errorDir)

                Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim csvPath = Path.Combine(errorDir, $"ErrorList_{Path.GetFileNameWithoutExtension(originalFileName)}_{timeStamp}.csv")

                Using sw As New StreamWriter(csvPath, False, Encoding.GetEncoding("shift-jis"))
                    sw.WriteLine(Chr(34) & "行" & Chr(34) & "," & Chr(34) & "エラー内容" & Chr(34))
                    For i As Integer = 0 To errors.Count - 1
                        Dim rowNum = (i + 1).ToString()
                        Dim errMsg = If(errors(i), "").Replace(Chr(34), Chr(34) & Chr(34))
                        sw.WriteLine(Chr(34) & rowNum & Chr(34) & "," & Chr(34) & errMsg & Chr(34))
                    Next
                End Using
            Catch ex As Exception
                Dim loggerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log")
                Dim logger As New Logger(loggerPath)
                logger.Write($"[SUZUKI_IMPORT] Error exporting CSV: {ex.Message}")
            End Try
        End Sub

        Private Sub MoveBackErrorFile(workFilePath As String, sourceFolder As String, originalFileName As String, userId As String)
            Try
                Dim nameNoExt = Path.GetFileNameWithoutExtension(originalFileName)
                Dim ext = Path.GetExtension(originalFileName)
                Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim newFileName = $"{nameNoExt}_{userId}_{timeStamp}{ext}"
                Dim targetPath = Path.Combine(sourceFolder, newFileName)
                File.Copy(workFilePath, targetPath, True)
                If File.Exists(targetPath) Then
                    File.Delete(workFilePath)
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub CopyErrorBackupFile(workFilePath As String, sourceFolder As String, originalFileName As String, userId As String)
            Try
                Dim nameNoExt = Path.GetFileNameWithoutExtension(originalFileName)
                Dim ext = Path.GetExtension(originalFileName)
                Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim newFileName = $"{nameNoExt}_{userId}_{timeStamp}{ext}"
                Dim targetPath = Path.Combine(sourceFolder, newFileName)
                File.Copy(workFilePath, targetPath, True)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub DeleteErrorRows(conn As OracleConnection, tran As OracleTransaction, infoCode As String, impFileStageId As Long)
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return

            Dim sql = $"DELETE FROM {tableName} WHERE active_flag = 'Y' AND item_no IS NULL"
            If impFileStageId > 0 Then
                sql &= " AND imp_file_id = :p_imp_file_id"
            End If

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                If impFileStageId > 0 Then
                    cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileStageId
                End If
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 取引先設定IDから取引先コードを取得する (Get Customer Code by Customer Setting ID)
        ''' </summary>
        Private Function GetCustomerCode(customerSettingId As Long) As String
            Try
                Using conn As New OracleConnection(_connectionString)
                    conn.Open()
                    Dim sql = "SELECT customer_code FROM customer_setting_mst WHERE customer_setting_id = :p_id AND active_flag = 'Y'"
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                            Return res.ToString().Trim()
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try
            Return "5455" ' スズキのデフォルト取引先コード (Default Suzuki customer code fallback)
        End Function

    End Class

End Namespace
