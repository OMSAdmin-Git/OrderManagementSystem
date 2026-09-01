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

            ' 2. ユーザー専用のWORKフォルダを確保（WORK/[UserID]/[取引先CD]/[FolderType]）
            Dim workRoot = ConfigurationManager.AppSettings("WorkFolderRoot")
            If String.IsNullOrWhiteSpace(workRoot) Then workRoot = "C:\ASTI\DATA\WORK"
            Dim customerCode = GetCustomerCode(customerSettingId)
            Dim userWorkDir = Path.Combine(workRoot, userId, customerCode, folderType.ToString())
            Utils.EnsureDirectory(userWorkDir)

            ' 1. 取込対象フォルダ（source folder）内の CSV を検索（サブフォルダは除外）
            Dim sourceFiles = If(Directory.Exists(folderPath), Directory.GetFiles(folderPath, "*.csv", SearchOption.TopDirectoryOnly), Array.Empty(Of String)())

            If sourceFiles.Length = 0 Then
                logger.Write("[SUZUKI_IMPORT] No CSV files found in source folder.")
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

                        ' 種別判定
                        Dim infoCode = SuzukiCsvParser.PeekInfoTypeCode(workFilePath)
                        If String.IsNullOrEmpty(infoCode) OrElse infoCode.Length <> 4 Then
                            logger.Write($"[SUZUKI_IMPORT] Invalid or missing INFO_TYPE_CODE in file: {fileName}")
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
                                stageRepo.Insert(conn, tran, stageRow)
                                
                                ' Fetch the generated ID
                                Dim fetchedRow = stageRepo.GetImpFilesStageFilename(conn, tran, fileName, folderPath)
                                If fetchedRow IsNot Nothing AndAlso fetchedRow.ImpFileStageId > 0 Then
                                    impFileStageId = fetchedRow.ImpFileStageId
                                End If

                                ' パース & インサート
                                ProcessFileByInfoCode(conn, tran, workFilePath, infoCode, userId, impRunId, impFileStageId)

                                ' 品番更新 (PRDSLSODRM)
                                UpdateAstiPartNumbers(conn, tran, infoCode)

                                ' 無効化チェック (ACTIVE_FLAG = 'N')
                                UpdateActiveFlags(conn, tran, infoCode)

                                ' エラーチェック
                                Dim errorRows = FindPartMatchingErrors(conn, tran, infoCode, workFilePath)
                                If errorRows.Count > 0 Then
                                    errorFilesCount += 1
                                    logger.Write($"[SUZUKI_IMPORT] Found {errorRows.Count} part matching errors in {fileName}")
                                    ExportErrorsToCsv(folderPath, fileName, errorRows, isBatch, webErrors)
                                    MoveBackErrorFile(workFilePath, folderPath, fileName, userId)
                                    DeleteErrorRows(conn, tran, infoCode)
                                Else
                                    validFilesCount += 1
                                End If

                                tran.Commit()
                                processCount += 1
                                logger.Write($"[SUZUKI_IMPORT] Successfully imported file: {fileName}")

                                ' ワークフォルダのファイル後処理 (Stage 1: エラー時もWORKフォルダから削除せず保持)
                                If File.Exists(workFilePath) Then
                                    Try
                                        ' 発見されたステージング行を結果リストに追加
                                        If outStageRows IsNot Nothing Then
                                            outStageRows.Add(stageRow)
                                        End If
                                    Catch ex As Exception
                                        logger.Write($"[SUZUKI_IMPORT] Post-processing error for {fileName}: {ex.Message}")
                                    End Try
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

        Private Sub ProcessFileByInfoCode(conn As OracleConnection, tran As OracleTransaction, filePath As String, infoCode As String, userId As String, impRunId As Long, impFileStageId As Long)
            Select Case infoCode
                Case "0501", "0502"
                    Dim repo As New Spirits0501And0502Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0501And0502(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0600", "0630"
                    Dim repo As New Spirits0600And0630Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.Parse0600And0630(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0602"
                    Dim repo As New Spirits0602Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.Parse0602(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0650"
                    Dim repo As New Spirits0650Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0650(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0651"
                    Dim repo As New Spirits0651Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0651(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0740"
                    Dim repo As New Spirits0740Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0740(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "0813"
                    Dim repo As New Spirits0813Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0813(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)




                Case "0814"
                    Dim repo As New Spirits0814Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits0814(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "6604", "6634"
                    Dim repo As New Spirits6604And6634Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits6604And6634(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "6624"
                    Dim repo As New Spirits6624Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits6624(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)

                Case "663N", "663S", "664T"
                    Dim repo As New Spirits663NAnd663SAnd66Repository(_connectionString)
                    Dim rows = SuzukiCsvParser.ParseSpirits663NAnd663SAnd664T(filePath)
                    For Each r In rows
                        r.ImpRunId = impRunId
                        r.ImpFileId = impFileStageId
                        r.CreatedUserId = userId
                        r.CreatedPgId = "OrderImport(Execute)"
                        r.CreatedAt = DateTime.Now
                        r.UpdatedUserId = userId
                        r.UpdatedPgId = "OrderImport(Execute)"
                        r.UpdatedAt = DateTime.Now
                    Next
                    Dim dbErr = repo.InsertRange(conn, tran, rows)

                    If Not String.IsNullOrEmpty(dbErr) Then Throw New Exception(dbErr)
            End Select
        End Sub

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

        Private Sub UpdateAstiPartNumbers(conn As OracleConnection, tran As OracleTransaction, infoCode As String)
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return

            Dim isZSuffix As Boolean = {"6604", "6624", "6634", "663N", "663S", "664T"}.Contains(infoCode)

            Dim sql As New StringBuilder()
            sql.AppendLine($"UPDATE {tableName} t")
            sql.AppendLine("SET t.item_no = (")
            sql.AppendLine("  SELECT p.FPRDCD FROM PRDSLSODRM p")
            sql.AppendLine("  WHERE p.FCUSTCD = '5455'")
            sql.AppendLine("    AND TRIM(p.FCUSTITEMNO) = TRIM(t.customer_item_no)")

            If isZSuffix Then
                sql.AppendLine("    AND UPPER(p.FPRDCD) LIKE '%Z'")
            Else
                sql.AppendLine("    AND UPPER(p.FPRDCD) NOT LIKE '%Z'")
            End If

            sql.AppendLine("    FETCH FIRST 1 ROWS ONLY")
            sql.AppendLine(")")
            sql.AppendLine("WHERE t.active_flag = 'Y' AND t.item_no IS NULL")

            Using cmd As New OracleCommand(sql.ToString(), conn)
                cmd.Transaction = tran
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

        Private Function FindPartMatchingErrors(conn As OracleConnection, tran As OracleTransaction, infoCode As String, csvFilePath As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return result

            Dim unmatchedItems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim sql = $"SELECT DISTINCT customer_item_no FROM {tableName} WHERE active_flag = 'Y' AND item_no IS NULL"
            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then
                            unmatchedItems.Add(reader.GetString(0).Trim())
                        End If
                    End While
                End Using
            End Using

            If unmatchedItems.Count > 0 AndAlso File.Exists(csvFilePath) Then
                Dim colIdx As Integer
                Select Case infoCode
                    Case "0600", "0630", "0650"
                        colIdx = 10
                    Case "0602", "0651", "0740"
                        colIdx = 8
                    Case "0814", "6604", "6624", "6634"
                        colIdx = 9
                    Case "663N", "663S", "664T"
                        colIdx = 7
                    Case "0501", "0502"
                        colIdx = 11
                    Case "0813"
                        colIdx = 6

                    Case Else
                        colIdx = 10
                End Select

                Dim lines = File.ReadAllLines(csvFilePath, Encoding.GetEncoding("shift-jis"))
                For lineIdx As Integer = 0 To lines.Length - 1
                    Dim line = lines(lineIdx)
                    If String.IsNullOrWhiteSpace(line) Then Continue For

                    Dim cols = line.Split(","c)
                    If colIdx < cols.Length Then
                        Dim custItemNo = cols(colIdx).Trim(" "c, """"c)
                        If Not String.IsNullOrEmpty(custItemNo) AndAlso unmatchedItems.Contains(custItemNo) Then
                            Dim csvRowNumber = lineIdx + 1
                            result.Add($"Row({csvRowNumber}): 品目No及び製品コードが取得できません。")
                        End If
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
                Dim baseName = Path.GetFileNameWithoutExtension(originalFileName)
                Dim match = Text.RegularExpressions.Regex.Match(baseName, "^(.*?)(?:_[a-zA-Z0-9\-]+_\d{8}_\d{6})+$")
                If match.Success Then
                    baseName = match.Groups(1).Value
                End If

                Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim newFileName = $"{baseName}_{userId}_{timeStamp}{Path.GetExtension(originalFileName)}"
                Dim targetPath = Path.Combine(sourceFolder, newFileName)
                File.Copy(workFilePath, targetPath, True)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub DeleteErrorRows(conn As OracleConnection, tran As OracleTransaction, infoCode As String)
            Dim tableName = GetTargetTableName(infoCode)
            If String.IsNullOrEmpty(tableName) Then Return

            Dim sql = $"DELETE FROM {tableName} WHERE active_flag = 'Y' AND item_no IS NULL"
            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
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
