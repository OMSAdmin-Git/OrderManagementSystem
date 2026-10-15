Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Oracle.ManagedDataAccess.Client
Imports OMS.Common
Imports OMS.Data
Imports OMS.Data.SUZUKI

Namespace Services

    ''' <summary>
    ''' スズキ SPIRITS データ本取込・受注登録サービス (Stage 2: SUZUKI_SPIRITS_DATA_IMPORT)
    ''' </summary>
    Public Class SuzukiDataImportService

        Private ReadOnly _connectionString As String
        Private ReadOnly _orderStageRepo As OrderStageRepository
        Private ReadOnly _calender As CalenderRepository

        Public Sub New()
            Me.New(Utils.GetConnectionString())
        End Sub

        Public Sub New(connectionString As String)
            _connectionString = connectionString
            _orderStageRepo = New OrderStageRepository(_connectionString)
            _calender = New CalenderRepository(_connectionString)
        End Sub

        ''' <summary>
        ''' 【Stage 2 バッチ実行】受注登録・洗替・消込・本番保存・STRA納期設定オーケストレーション (工程1〜8)
        ''' </summary>
        ''' <param name="customerSettingId">取引先設定ID (初期値: 5455 PC未設定)</param>
        ''' <param name="folderType">フォルダ区分 (4: 混在)</param>
        ''' <param name="impRunId">取込実行ID</param>
        ''' <param name="userId">ユーザーID</param>
        ''' <param name="isBatch">バッチフラグ</param>
        ''' <returns>登録件数</returns>
        Public Function ExecuteOrderRegistration(customerSettingId As Long, folderType As Integer, impRunId As Long, userId As String, Optional isBatch As Boolean = True) As Integer
            Dim logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs")
            If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
            Dim loggerPath = Path.Combine(logDir, $"SuzukiImport_{DateTime.Now:yyyyMMdd}.log")
            Dim logger As New Logger(loggerPath)
            logger.Write($"[SUZUKI_IMPORT] 受注取込 (Import) を開始しました (取引先設定ID: {customerSettingId}, フォルダ種別: {folderType}, 実行管理ID: {impRunId})")

            Dim pgId As String = "SUZUKI_SPIRITS_DATA_IMPORT"
            Dim totalRegistered As Integer = 0
            Dim folderPath As String = GetFolderPath(customerSettingId, folderType)
            Dim committedSettings As New List(Of Long)()

            Using conn As New OracleConnection(_connectionString)
                conn.Open()

                Using tran As OracleTransaction = conn.BeginTransaction()
                    Try
                        ' 【工程 1】受注ワーク削除 (DELETE ①)
                        logger.Write("[SUZUKI_IMPORT] 工程 1: ORDERS_STAGE の既存ワークデータをクリアしています...")
                        _orderStageRepo.DeleteOrdersStageForSuzuki(tran, customerSettingId, folderType)

                        ' 【工程 2 & 3】スズキ専用テーブルよりデータ抽出 & ORDERS_STAGE 行オブジェクト作成
                        logger.Write("[SUZUKI_IMPORT] 工程 2 & 3: スズキ専用ステージングテーブルより有効データを抽出中...")
                        Dim rowsForTemp2 As New List(Of OrdersStageRow)()
                        Dim errors As New List(Of String)()

                        ExtractStagingRows(conn, tran, customerSettingId, folderType, 0, impRunId, userId, pgId, rowsForTemp2)
                        logger.Write($"[SUZUKI_IMPORT] ステージングデータ {rowsForTemp2.Count} 件を抽出しました。")

                        ' 【工程 4】プロフィットセンター (PC) 解決 & 特殊コード加算・更新 (0630, 663N, 663S)
                        logger.Write("[SUZUKI_IMPORT] 工程 4: プロフィットセンター(PC)解決および増減・変更データ(0630, 663N, 663S)を反映中...")
                        Dim adjustmentStageIds As New HashSet(Of Long)()
                        ResolveProfitCenters(conn, tran, rowsForTemp2, customerSettingId)
                        ApplySpecialAdjustments(conn, tran, 0, impRunId, rowsForTemp2, userId, pgId, adjustmentStageIds)
                        ResolveProfitCenters(conn, tran, rowsForTemp2, customerSettingId)

                        If rowsForTemp2.Count = 0 Then
                            logger.Write("[SUZUKI_IMPORT] 処理対象の有効データが存在しないため、登録処理をスキップします。")
                            tran.Commit()
                            Return 0
                        End If

                        ' 【工程 5】エラーチェック
                        logger.Write("[SUZUKI_IMPORT] 工程 5: ワークデータの整合性チェックを実行中...")
                        ValidateStagedRows(rowsForTemp2, customerSettingId, errors)

                        If errors.Count > 0 Then
                            logger.Write($"[SUZUKI_IMPORT] データ検証エラーが {errors.Count} 件発生しました。")
                            If isBatch Then
                                ExportErrorsToCsv(folderPath, errors)
                            End If
                            tran.Rollback()
                            Throw New Exception($"Stage 2 Validation Error: {String.Join("; ", errors.Take(5))}")
                        End If

                        ' 【工程 6】受注ワーク登録 & 洗い替え・消込処理
                        logger.Write("[SUZUKI_IMPORT] 工程 6: ORDERS_STAGE への展開および消込処理を実行中...")
                        Dim insertedCount As Integer = _orderStageRepo.InsertRange(tran, rowsForTemp2)
                        logger.Write($"[SUZUKI_IMPORT] ORDERS_STAGE へ {insertedCount} 件のデータを登録しました。")

                        ' 解決された各取引先設定IDごとに正規データ読み込み & 内示洗い替え・消込
                        committedSettings = rowsForTemp2.Select(Function(r) r.CustomerSettingId).Distinct().ToList()
                        For Each csId In committedSettings
                            _orderStageRepo.InsertStageFromOrders(tran, csId)

                            ' 出荷先変換 (SUZUKI_SPIRITS_CONVERSION_MST)
                            _orderStageRepo.ConvertShipToForSuzuki(tran, csId)

                            ' 内示加工 (0602除外、0600/0813/6604洗い替え)
                            _orderStageRepo.ReplaceNaijiRelationSuzuki(tran, 0, csId, DateTime.Now, userId, pgId)

                            ' 受注残加工
                            Dim rFlag As String = GetRuleFlag(conn, tran, csId, folderType, "RECONCILE_FLAG")
                            If rFlag = "Y" Then
                                Dim rType As Integer = GetRuleReconcileType(conn, tran, csId, folderType)
                                _orderStageRepo.BacklogForecast(tran, csId, 0, rType, DateTime.Now, userId, pgId)
                            End If

                            ' 納入指示加工: 打切 / 取消
                            _orderStageRepo.UpdateClese(tran, 0, 3)
                            _orderStageRepo.UpdateCancel(tran, 0, 3)

                            ' 納入指示加工: 内示消込 (0650/0651->0600, 0814->0813, 6624->6604)
                            Dim fcstFlag As String = GetRuleFlag(conn, tran, csId, folderType, "FCST_RECONCILE_FLAG")
                            If fcstFlag = "Y" Then
                                _orderStageRepo.ReconcileForecastSuzuki(tran, csId, 0, DateTime.Now, userId, pgId)
                            End If

                            ' 納入指示 新規確定
                            _orderStageRepo.UpdateKakuteiNewOrdersSuzuki(tran, 0)
                        Next

                        ' 【工程 7】本番保存 (INSERT ORDERS / INSERT ORDER_HISTORY)
                        logger.Write("[SUZUKI_IMPORT] 工程 7: 受注マスタ (ORDERS) および履歴 (ORDER_HISTORY) へ本登録中...")
                        For Each csId In committedSettings
                            _orderStageRepo.UpdateOrdersFromStage(tran, csId, DateTime.Now, userId, pgId)
                            Dim saveCount As Integer = _orderStageRepo.InsertOrdersFromStage(tran, csId, DateTime.Now, userId, pgId)
                            _orderStageRepo.InsertHistoryFromOrders(tran, csId, DateTime.Now, userId, pgId)
                            _orderStageRepo.DeleteProcessedOrdersRange(tran, userId, csId)

                            ' 【バッチ用クリーンパージ】保存完了したバッチ対象の PROCESSED レコードを確実にステージから削除
                            If isBatch Then
                                Using delCmd As New OracleCommand("DELETE FROM orders_stage WHERE customer_code = '5455' AND (imp_run_id = :p_run_id OR created_user_id = :p_user_id) AND status = 'PROCESSED'", conn)
                                    delCmd.Transaction = tran
                                    delCmd.Parameters.Add(":p_run_id", OracleDbType.Int64).Value = impRunId
                                    delCmd.Parameters.Add(":p_user_id", OracleDbType.Varchar2).Value = userId
                                    delCmd.ExecuteNonQuery()
                                End Using
                            End If

                            totalRegistered += saveCount
                        Next

                        ' 受注ファイル取込ページ [保存]ボタンと同等の処理 (SPIRITSデータ受注登録 仕様書 P.9-10)
                        ' 処理対象の IMP_FILE_STAGE_ID を IMP_FILES へ昇格し、IMP_FILES_STAGE から削除する
                        Dim processedStageIds = rowsForTemp2.
                            Where(Function(r) r.ImpFileStageId.HasValue).
                            Select(Function(r) r.ImpFileStageId.Value).
                            Distinct().
                            ToList()

                        ' 【特殊コード対応】0630, 663N, 663S 等の増減・変更ファイルIDも決済対象に含める
                        For Each adjId In adjustmentStageIds
                            If Not processedStageIds.Contains(adjId) Then
                                processedStageIds.Add(adjId)
                            End If
                        Next

                        ' 処理対象のファイル名を収集 (MoveWorkFilesToCompleted で登録済みファイルのみを選択的に移動するため)
                        ' ※ Oracle ORA-01795 (IN句の要素数1000個上限) 対策として500件ずつチャンク分割して取得
                        Dim processedFileNames As New List(Of String)()
                        If processedStageIds.Count > 0 Then
                            For i As Integer = 0 To processedStageIds.Count - 1 Step 500
                                Dim chunk = processedStageIds.Skip(i).Take(500).ToList()
                                Dim idListStr = String.Join(",", chunk)
                                Using cmdNames As New OracleCommand($"SELECT DISTINCT NVL(staged_file_name, file_name) AS fname FROM imp_files_stage WHERE imp_file_stage_id IN ({idListStr})", conn)
                                    cmdNames.Transaction = tran
                                    Using rNames = cmdNames.ExecuteReader()
                                        While rNames.Read()
                                            If Not rNames.IsDBNull(0) Then
                                                Dim fn = rNames(0).ToString().Trim()
                                                If Not String.IsNullOrEmpty(fn) AndAlso Not processedFileNames.Contains(fn) Then
                                                    processedFileNames.Add(fn)
                                                End If
                                            End If
                                        End While
                                    End Using
                                End Using
                            Next
                        End If

                        Dim stageRepo As New ImpFilesStageRepository(_connectionString)
                        For Each stgId In processedStageIds
                            ' status を 'PARSED' に更新 (InsertImpFileFromStage の抽出条件が status = 'PARSED' のため)
                            stageRepo.Update(conn, tran, kImpFileStageId:=stgId, status:="PARSED", updatedAt:=DateTime.Now)

                            ' 1. IMP_FILES へ昇格登録
                            Dim newImpFileId As Long = stageRepo.InsertImpFileFromStage(tran, stgId, DateTime.Now, userId, pgId)

                            ' 2. IMP_FILES_STAGE から削除
                            stageRepo.DeleteImpFileStageRange(tran, stgId)

                            ' 3. ORDERS.IMP_FILE_ID を更新
                            _orderStageRepo.UpdateOrdersImpFileId(tran, newImpFileId, stgId)

                            ' 4. スズキ専用テーブルのステータス更新 (STATUS = 'REGISTERED')
                            _orderStageRepo.UpdateSuzukiStagingStatusRegistered(tran, stgId, userId, pgId)
                        Next

                        ' スズキ専用テーブルのステータス更新 (STATUS = 'REGISTERED')
                        _orderStageRepo.UpdateSuzukiStagingStatusRegisteredByRunId(tran, impRunId, userId, pgId)

                        ' WORKファイルのCOMPLETEDへの移動 (登録対象ファイルのみ選択的に移動)
                        MoveWorkFilesToCompleted(conn, tran, impRunId, folderPath, userId, folderType, processedFileNames)

                        tran.Commit()
                        logger.Write($"[SUZUKI_IMPORT] 受注本番テーブルへの登録が正常に完了しました (登録件数: {totalRegistered} 件)。")

                    Catch ex As Exception
                        tran.Rollback()
                        logger.Write($"[SUZUKI_IMPORT] 受注取込処理中にエラーが発生しました: {ex.Message}")
                        Throw
                    End Try
                End Using

                ' 【工程 8】STRA納期設定 & IMP_RUN 完了更新 (バッチ実行のみ)
                If isBatch Then
                    logger.Write("[SUZUKI_IMPORT] 工程 8: 後続の STRA納期設定処理を実行中...")
                    ExecuteStraDeliveryDateSetting(conn, committedSettings, userId, logger)
                End If

            End Using

            Return totalRegistered
        End Function

        ''' <summary>
        ''' 【Stage 2 画面実行】OrderImport.aspx (取込実行ボタン) より呼び出され、ORDERS_STAGEへの変換・投入・洗替を行う
        ''' </summary>
        Public Function ExecuteStageImportToOrdersStage(conn As OracleConnection,
                                                       tran As OracleTransaction,
                                                       customerSettingId As Long,
                                                       customerCode As String,
                                                       impfilestageId As Long,
                                                       folderType As Integer,
                                                       userId As String,
                                                       pgId As String,
                                                       errors As List(Of String),
                                                       rowsForTemp2 As List(Of OrdersStageRow)) As Boolean
            Try
                ' 1. ORDERS_STAGE 物理削除 (DELETE ①)
                _orderStageRepo.DeleteOrdersStageForSuzuki(tran, customerSettingId, folderType)

                ' 2. スズキ専用テーブルよりデータ抽出
                Dim tempRows As New List(Of OrdersStageRow)()
                ExtractStagingRows(conn, tran, customerSettingId, folderType, impfilestageId, 0, userId, pgId, tempRows)

                ' 3. プロフィットセンター (PC) 解決 & 特殊コード加算・更新 (0630, 663N, 663S)
                ResolveProfitCenters(conn, tran, tempRows, customerSettingId)
                ApplySpecialAdjustments(conn, tran, impfilestageId, 0, tempRows, userId, pgId)
                ResolveProfitCenters(conn, tran, tempRows, customerSettingId)

                If tempRows.Count = 0 Then
                    Return False
                End If

                ' 4. エラーチェック
                ValidateStagedRows(tempRows, customerSettingId, errors)
                If errors.Count > 0 Then
                    Return False
                End If

                ' 5. 出荷先変換 (SUZUKI_SPIRITS_CONVERSION_MST)
                ApplyShipToConversionInMemory(conn, tran, tempRows)

                ' 6. 呼び出し元の rowsForTemp2 に追加
                rowsForTemp2.AddRange(tempRows)
                Return True

            Catch ex As Exception
                errors.Add($"スズキ取込エラー: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' スズキ専用テーブルから有効レコードを抽出し OrdersStageRow に変換する
        ''' </summary>
        Private Sub ExtractStagingRows(conn As OracleConnection,
                                      tran As OracleTransaction,
                                      customerSettingId As Long,
                                      folderType As Integer,
                                      impFileId As Long,
                                      impRunId As Long,
                                      userId As String,
                                      pgId As String,
                                      outRows As List(Of OrdersStageRow))

            Dim billingTo As String = ""
            Dim currencyCode As String = ""
            Dim errMsg As String = ""
            _orderStageRepo.GetBillingTo("5455", billingTo, errMsg)
            _orderStageRepo.GetCurrencyCode("5455", currencyCode, errMsg)

            Dim dateCache As New Dictionary(Of Date, Date)()
            Dim unitCache As New Dictionary(Of String, (DemandUnit As String, ShipWh As String))()

            ' 消込条件区分(reconcile_type)の取得 (デフォルト: 1)
            Dim ruleReconcileType As Short = 1
            Dim reconcileErrMsg As String = ""
            Dim tempReconcile As Integer = 1
            If _orderStageRepo.GetReconcileType(customerSettingId, CLng(folderType), tempReconcile, reconcileErrMsg) Then
                ruleReconcileType = CShort(tempReconcile)
            End If

            ' 1. 0600 (確定受注データ・月次内示)
            LoadSpirits0600(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 2. 0602 (確定受注データ・追加内示)
            LoadSpirits0602(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 3. 0650 (納入指示データ・通常)
            LoadSpirits0650(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 4. 0651 (納入指示データ・臨時)
            LoadSpirits0651(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 5. 0813 (支給品情報データ)
            ' 【2026/09/25 仕様変更】0813はSPIRITS仕様より除外されたため、ORDERS_STAGEへの展開を行いません。
            ' LoadSpirits0813(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 6. 0814 (支給品納入指示データ)
            LoadSpirits0814(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 7. 6604 (注文データ・内示)
            LoadSpirits6604(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)

            ' 8. 6624 (注文データ・確定/指示)
            LoadSpirits6624(conn, tran, impFileId, impRunId, customerSettingId, billingTo, currencyCode, dateCache, unitCache, userId, pgId, ruleReconcileType, outRows)
        End Sub

        ''' <summary>
        ''' SPIRITSデータマッピング仕様に基づく 取引先情報区分 (CUSTOMER_INFO_TYPE) の変換
        ''' </summary>
        Public Shared Function ConvertCustomerInfoType(infoCode As String, orderDataType As String, orderNotes As String) As String
            Select Case infoCode
                Case "0600", "6604"
                    Select Case (If(orderDataType, "")).Trim()
                        Case "1", "2" : Return "新規"
                        Case "3" : Return "変更"
                        Case Else : Return ""
                    End Select
                Case "0602"
                    Select Case (If(orderDataType, "")).Trim()
                        Case "2" : Return "新規"
                        Case "3" : Return "変更"
                        Case Else : Return ""
                    End Select
                Case "0651"
                    If Not String.IsNullOrEmpty(orderNotes) AndAlso orderNotes.Length >= 17 Then
                        If orderNotes.Substring(16, 1) <> "0" Then
                            Return "打切"
                        End If
                    End If
                    Return ""
                Case Else
                    Return ""
            End Select
        End Function

        ''' <summary>
        ''' 取引先情報区分から 情報区分 (INFO_TYPE) を解決
        ''' INFO_TYPE_MST を優先参照し、未定義または取得失敗時は標準コード (新規:I, 変更:U, 打切:N) にフォールバック
        ''' </summary>
        Public Function ResolveInfoType(customerSettingId As Long, customerInfoType As String) As String
            Return ResolveInfoTypeStatic(_orderStageRepo, customerSettingId, customerInfoType)
        End Function

        Public Shared Function ResolveInfoTypeStatic(repo As OrderStageRepository, customerSettingId As Long, customerInfoType As String) As String
            If String.IsNullOrWhiteSpace(customerInfoType) Then Return ""

            Dim cleanType = customerInfoType.Trim()
            If repo IsNot Nothing AndAlso customerSettingId > 0 Then
                Try
                    Dim resolved As String = ""
                    Dim errMsg As String = ""
                    If repo.GetInfoType(customerSettingId, cleanType, resolved, errMsg) AndAlso Not String.IsNullOrEmpty(resolved) Then
                        Return resolved.Trim()
                    End If
                Catch
                    ' fallback to standard code
                End Try
            End If

            Select Case cleanType
                Case "新規" : Return "I"
                Case "変更" : Return "U"
                Case "打切" : Return "N"
                Case Else : Return ""
            End Select
        End Function

        ''' <summary>
        ''' SPIRITSデータマッピング仕様に基づく 分割区分 (PRORATED_TYPE) の変換
        ''' 1: 日割り, 2: 日割以外
        ''' </summary>
        Public Shared Function ConvertProratedType(infoCode As String, deliveryDateType As String) As Short
            Dim dt = (If(deliveryDateType, "")).Trim()
            Select Case infoCode
                Case "0600"
                    Select Case dt
                        Case "1" : Return 2
                        Case "5" : Return 1
                        Case Else : Return 1
                    End Select

                Case "6604"
                    Select Case dt
                        Case "1", "4" : Return 2
                        Case "5" : Return 1
                        Case Else : Return 1
                    End Select

                Case "0650", "0651", "0813", "0814"
                    Select Case dt
                        Case "5", "6" : Return 1
                        Case Else : Return 1
                    End Select

                Case "0602", "6624"
                    Select Case dt
                        Case "5" : Return 1
                        Case Else : Return 1
                    End Select

                Case Else
                    Return 1
            End Select
        End Function

#Region "スズキ専用テーブル別 データ抽出メソッド"

        Private Sub LoadSpirits0600(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0600AND0630 WHERE info_type_code = '0600' AND active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0600", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")
                        Dim orderDataType = If(r.IsDBNull(r.GetOrdinal("order_data_type")), "", r("order_data_type").ToString())
                        Dim custInfoType = ConvertCustomerInfoType("0600", orderDataType, "")
                        Dim resolvedInfoType = ResolveInfoType(customerSettingId, custInfoType)

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = "",
                            .DemandStatus = "F",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "O",
                            .DeliveryInstrFlag = "N",
                            .TotalShipQty = Nothing,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 1,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = custInfoType,
                            .InfoType = resolvedInfoType,
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = If(r.IsDBNull(r.GetOrdinal("target_reference_date_type")), "", r("target_reference_date_type").ToString()),
                            .TargetReferenceDate = If(r.IsDBNull(r.GetOrdinal("target_reference_date")), "", r("target_reference_date").ToString()),
                            .InfoTypeCode = "0600",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString()),
                            .ProductionMonthType = If(r.IsDBNull(r.GetOrdinal("production_month_type")), "", r("production_month_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits0602(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0602 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0602", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")
                        Dim orderDataType = If(r.IsDBNull(r.GetOrdinal("order_data_type")), "", r("order_data_type").ToString())
                        Dim custInfoType = ConvertCustomerInfoType("0602", orderDataType, "")
                        Dim resolvedInfoType = ResolveInfoType(customerSettingId, custInfoType)

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "F",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "O",
                            .DeliveryInstrFlag = "N",
                            .TotalShipQty = Nothing,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 1,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = custInfoType,
                            .InfoType = resolvedInfoType,
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .TargetReferenceDateType = "",
                            .TargetReferenceDate = "",
                            .InfoTypeCode = "0602",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString()),
                            .ProductionMonthType = If(r.IsDBNull(r.GetOrdinal("production_month_type")), "", r("production_month_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits0650(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0650 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0650", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "O",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "K",
                            .DeliveryInstrFlag = "Y",
                            .TotalShipQty = 0,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 3,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = "",
                            .InfoType = "",
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = "",
                            .TargetReferenceDate = "",
                            .InfoTypeCode = "0650",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString()),
                            .ProductionMonthType = If(r.IsDBNull(r.GetOrdinal("production_month_type")), "", r("production_month_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits0651(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0651 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0651", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")
                        Dim orderNotes = If(r.IsDBNull(r.GetOrdinal("order_notes")), "", r("order_notes").ToString())
                        Dim custInfoType = ConvertCustomerInfoType("0651", "", orderNotes)
                        Dim resolvedInfoType = ResolveInfoType(customerSettingId, custInfoType)

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "O",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "K",
                            .DeliveryInstrFlag = "Y",
                            .TotalShipQty = 0,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 3,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = custInfoType,
                            .InfoType = resolvedInfoType,
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = "",
                            .TargetReferenceDate = "",
                            .InfoTypeCode = "0651",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString()),
                            .ProductionMonthType = If(r.IsDBNull(r.GetOrdinal("production_month_type")), "", r("production_month_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits0813(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0813 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0813", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "F",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "O",
                            .DeliveryInstrFlag = "N",
                            .TotalShipQty = Nothing,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 1,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = "0813",
                            .InfoType = "0",
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = If(r.IsDBNull(r.GetOrdinal("target_reference_date_type")), "", r("target_reference_date_type").ToString()),
                            .TargetReferenceDate = If(r.IsDBNull(r.GetOrdinal("target_reference_date")), "", r("target_reference_date").ToString()),
                            .InfoTypeCode = "0813",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits0814(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0814 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("0814", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "O",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "K",
                            .DeliveryInstrFlag = "Y",
                            .TotalShipQty = 0,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 3,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = "",
                            .InfoType = "",
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = If(r.IsDBNull(r.GetOrdinal("target_reference_date_type")), "", r("target_reference_date_type").ToString()),
                            .TargetReferenceDate = If(r.IsDBNull(r.GetOrdinal("target_reference_date")), "", r("target_reference_date").ToString()),
                            .InfoTypeCode = "0814",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits6604(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_6604AND6634 WHERE info_type_code = '6604' AND active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("6604", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")
                        Dim orderDataType = If(r.IsDBNull(r.GetOrdinal("order_data_type")), "", r("order_data_type").ToString())
                        Dim custInfoType = ConvertCustomerInfoType("6604", orderDataType, "")
                        Dim resolvedInfoType = ResolveInfoType(customerSettingId, custInfoType)

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = "",
                            .DemandStatus = "F",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "O",
                            .DeliveryInstrFlag = "N",
                            .TotalShipQty = Nothing,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 1,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = custInfoType,
                            .InfoType = resolvedInfoType,
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .TargetReferenceDateType = If(r.IsDBNull(r.GetOrdinal("target_reference_date_type")), "", r("target_reference_date_type").ToString()),
                            .TargetReferenceDate = If(r.IsDBNull(r.GetOrdinal("target_reference_date")), "", r("target_reference_date").ToString()),
                            .InfoTypeCode = "6604",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

        Private Sub LoadSpirits6624(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, customerSettingId As Long, billingTo As String, currencyCode As String, dateCache As Dictionary(Of Date, Date), unitCache As Dictionary(Of String, (DemandUnit As String, ShipWh As String)), userId As String, pgId As String, ruleReconcileType As Short, outRows As List(Of OrdersStageRow))
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_6624 WHERE active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim itemNo = r("item_no").ToString()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        Dim dueDate = If(delDate.HasValue, GetWorkingDueDateCached(delDate.Value, dateCache), CType(Nothing, Date?))
                        Dim unitWh = GetUnitAndWhCached(itemNo, unitCache)
                        Dim dateType = If(r.IsDBNull(r.GetOrdinal("delivery_date_type")), "", r("delivery_date_type").ToString())
                        Dim prorated = ConvertProratedType("6624", dateType)
                        Dim orderQty = GetSafeInt64(r, "order_qty")

                        outRows.Add(New OrdersStageRow With {
                            .CustomerSettingId = customerSettingId,
                            .CustomerCode = "5455",
                            .BillingTo = billingTo,
                            .CustomerOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString()),
                            .DemandStatus = "O",
                            .ShipTo = If(r.IsDBNull(r.GetOrdinal("delivery_code")), "", r("delivery_code").ToString()),
                            .OrderDate = If(r.IsDBNull(r.GetOrdinal("publication_date")), Nothing, CType(r("publication_date"), Date?)),
                            .DueDate = dueDate,
                            .CustomerItemNo = If(r.IsDBNull(r.GetOrdinal("customer_item_no")), "", r("customer_item_no").ToString()),
                            .ItemNo = itemNo,
                            .DemandQty = orderQty,
                            .DemandUnit = unitWh.DemandUnit,
                            .CurrencyCode = currencyCode,
                            .ShipStockLocation = unitWh.ShipWh,
                            .CompanyId = "1000",
                            .ProductCode = itemNo,
                            .BillingStandard = "S",
                            .ShipProcessType = "K",
                            .DeliveryInstrFlag = "Y",
                            .TotalShipQty = 0,
                            .TransportMethod = "2",
                            .PreDailyOrderQty = If(orderQty.HasValue, Convert.ToDecimal(orderQty.Value), CType(Nothing, Decimal?)),
                            .PreDailyDeliveryDate = delDate,
                            .ImpFileStageId = If(r.IsDBNull(r.GetOrdinal("imp_file_id")), Nothing, Convert.ToInt64(r("imp_file_id"))),
                            .OrderType = 3,
                            .ReconcileType = ruleReconcileType,
                            .ProratedType = prorated,
                            .CustomerInfoType = "",
                            .InfoType = "",
                            .SelfFcstFlag = "N",
                            .SelfFcstDeleteFlag = "N",
                            .ImpRunId = If(r.IsDBNull(r.GetOrdinal("imp_run_id")), Nothing, Convert.ToInt64(r("imp_run_id"))),
                            .Status = "IMPORTED",
                            .ActiveFlag = "Y",
                            .CreatedAt = DateTime.Now,
                            .CreatedUserId = userId,
                            .CreatedPgId = pgId,
                            .UpdatedAt = DateTime.Now,
                            .UpdatedUserId = userId,
                            .UpdatedPgId = pgId,
                            .DeliveryTime = If(r.IsDBNull(r.GetOrdinal("delivery_time")), Nothing, Convert.ToDecimal(r("delivery_time"))),
                            .ContainerCapacity = If(r.IsDBNull(r.GetOrdinal("capacity")), Nothing, Convert.ToDecimal(r("capacity"))),
                            .InitialFlag = If(r.IsDBNull(r.GetOrdinal("first_article_type")), "", r("first_article_type").ToString()),
                            .TargetReferenceDateType = "",
                            .TargetReferenceDate = "",
                            .InfoTypeCode = "6624",
                            .DeliveryType = If(r.IsDBNull(r.GetOrdinal("delivery_type")), "", r("delivery_type").ToString())
                        })
                    End While
                End Using
            End Using
        End Sub

#End Region

#Region "プロフィットセンター解決 & 特殊コード加算・更新"

        ''' <summary>
        ''' 【Stage 2 工程4】プロフィットセンター (PC) 解決
        ''' USRDEFFLDF (FTABLEID='ITEMM', FRECKEY=ProductCode) から FUSRSTR1 を取得し、
        ''' CUSTOMER_SETTING_MST (CUSTOMER_CODE='5455', PROFIT_CENTER=FUSRSTR1) から CUSTOMER_SETTING_ID を特定・更新
        ''' </summary>
        Private Sub ResolveProfitCenters(conn As OracleConnection, tran As OracleTransaction, rows As List(Of OrdersStageRow), defaultSettingId As Long)
            Dim pcCache As New Dictionary(Of String, String)()
            Dim csMap As New Dictionary(Of String, Long)()

            Using cmd As New OracleCommand("SELECT customer_setting_id, profit_center FROM customer_setting_mst WHERE customer_code = '5455' AND active_flag = 'Y' AND profit_center IS NOT NULL", conn)
                cmd.Transaction = tran
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim csId = Convert.ToInt64(r("customer_setting_id"))
                        Dim pc = r("profit_center").ToString().Trim()
                        If Not String.IsNullOrEmpty(pc) AndAlso Not csMap.ContainsKey(pc) Then
                            csMap(pc) = csId
                        End If
                    End While
                End Using
            End Using

            For Each row In rows
                Dim pCode = If(String.IsNullOrEmpty(row.ProductCode), row.ItemNo, row.ProductCode)
                If String.IsNullOrEmpty(pCode) Then Continue For

                Dim pc As String = ""
                If pcCache.ContainsKey(pCode) Then
                    pc = pcCache(pCode)
                Else
                    Using cmd As New OracleCommand("SELECT FUSRSTR1 FROM USRDEFFLDF WHERE FTABLEID = 'ITEMM' AND FRECKEY = :p_key", conn)
                        cmd.Transaction = tran
                        cmd.Parameters.Add(":p_key", OracleDbType.Varchar2).Value = pCode
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                            pc = res.ToString().Trim()
                        End If
                    End Using
                    pcCache(pCode) = pc
                End If

                If Not String.IsNullOrEmpty(pc) AndAlso csMap.ContainsKey(pc) Then
                    row.CustomerSettingId = csMap(pc)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 【Stage 2 工程4】特殊コード加算・更新 (0630, 663N, 663S)
        ''' </summary>
        Private Sub ApplySpecialAdjustments(conn As OracleConnection,
                                           tran As OracleTransaction,
                                           impFileId As Long,
                                           impRunId As Long,
                                           rows As List(Of OrdersStageRow),
                                           userId As String,
                                           pgId As String,
                                           Optional outAdjustmentStageIds As HashSet(Of Long) = Nothing)
            Dim dateCache As New Dictionary(Of Date, Date)()

            ' 1. 0630: 増減 (0600に対する数量加算)
            ApplyAdjustment0630(conn, tran, impFileId, impRunId, rows, userId, pgId, outAdjustmentStageIds)

            ' 2. 663N: 納期変更 (6624に対する納期変更)
            ApplyAdjustment663N(conn, tran, impFileId, impRunId, rows, dateCache, userId, pgId, outAdjustmentStageIds)

            ' 3. 663S: 数量変更 (6624に対する数量変更)
            ApplyAdjustment663S(conn, tran, impFileId, impRunId, rows, userId, pgId, outAdjustmentStageIds)
        End Sub

        Private Sub ApplyAdjustment0630(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, rows As List(Of OrdersStageRow), userId As String, pgId As String, Optional outAdjustmentStageIds As HashSet(Of Long) = Nothing)
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_0600AND0630 WHERE info_type_code = '0630' AND active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim adjFileId = GetSafeInt64(r, "imp_file_id")
                        If adjFileId.HasValue AndAlso adjFileId.Value > 0 AndAlso outAdjustmentStageIds IsNot Nothing Then
                            outAdjustmentStageIds.Add(adjFileId.Value)
                        End If

                        Dim itemNo = r("item_no").ToString().Trim()
                        Dim delDate = GetSafeDate(r, "delivery_date")
                        If Not delDate.HasValue Then Continue While
                        Dim trdType = If(r.IsDBNull(r.GetOrdinal("target_reference_date_type")), "", r("target_reference_date_type").ToString().Trim())
                        Dim trd = If(r.IsDBNull(r.GetOrdinal("target_reference_date")), "", r("target_reference_date").ToString().Trim())
                        Dim addQty = If(GetSafeInt64(r, "order_qty"), 0L)

                        Dim target = rows.FirstOrDefault(Function(x) x.InfoTypeCode = "0600" AndAlso
                                                                     x.ItemNo = itemNo AndAlso
                                                                     x.PreDailyDeliveryDate.HasValue AndAlso
                                                                     x.PreDailyDeliveryDate.Value.Date = delDate.Value.Date AndAlso
                                                                     (x.TargetReferenceDateType Is Nothing OrElse x.TargetReferenceDateType.Trim() = trdType) AndAlso
                                                                     (x.TargetReferenceDate Is Nothing OrElse x.TargetReferenceDate.Trim() = trd))

                        If target IsNot Nothing Then
                            target.DemandQty = (If(target.DemandQty, 0L)) + addQty
                            target.PreDailyOrderQty = target.DemandQty
                        Else
                            Using cmdOrd As New OracleCommand("SELECT * FROM orders WHERE customer_code = '5455' AND info_type_code = '0600' AND item_no = :p_item_no AND TRUNC(pre_daily_delivery_date) = :p_del_date AND NVL(TRIM(target_reference_date_type), ' ') = NVL(:p_trd_type, ' ') AND NVL(TRIM(target_reference_date), ' ') = NVL(:p_trd, ' ') AND active_flag = 'Y' AND status IN ('PROCESSED', 'RECONCILED', 'IMPORTED')", conn)
                                cmdOrd.Transaction = tran
                                cmdOrd.BindByName = True
                                cmdOrd.Parameters.Add(":p_item_no", OracleDbType.Varchar2).Value = itemNo
                                cmdOrd.Parameters.Add(":p_del_date", OracleDbType.Date).Value = delDate.Value.Date
                                cmdOrd.Parameters.Add(":p_trd_type", OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(trdType), " ", trdType)
                                cmdOrd.Parameters.Add(":p_trd", OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(trd), " ", trd)
                                Using ordR = cmdOrd.ExecuteReader()
                                    If ordR.Read() Then
                                        Dim fallbackRow = CreateOrdersStageRowFromOrdersReader(ordR, userId, pgId)
                                        fallbackRow.DemandQty = (If(fallbackRow.DemandQty, 0L)) + addQty
                                        fallbackRow.PreDailyOrderQty = fallbackRow.DemandQty
                                        If impFileId > 0 Then
                                            fallbackRow.ImpFileStageId = impFileId
                                            fallbackRow.ImpFileId = impFileId
                                        ElseIf adjFileId.HasValue AndAlso adjFileId.Value > 0 Then
                                            fallbackRow.ImpFileStageId = adjFileId.Value
                                            fallbackRow.ImpFileId = adjFileId.Value
                                        End If
                                        If impRunId > 0 Then
                                            fallbackRow.ImpRunId = impRunId
                                        End If
                                        rows.Add(fallbackRow)
                                    End If
                                End Using
                            End Using
                        End If
                    End While
                End Using
            End Using
        End Sub

        Private Sub ApplyAdjustment663N(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, rows As List(Of OrdersStageRow), dateCache As Dictionary(Of Date, Date), userId As String, pgId As String, Optional outAdjustmentStageIds As HashSet(Of Long) = Nothing)
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_663NAND663SAND664T WHERE info_type_code = '663N' AND active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim adjFileId = GetSafeInt64(r, "imp_file_id")
                        If adjFileId.HasValue AndAlso adjFileId.Value > 0 AndAlso outAdjustmentStageIds IsNot Nothing Then
                            outAdjustmentStageIds.Add(adjFileId.Value)
                        End If

                        Dim custOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString().Trim())
                        Dim newDelDate = GetSafeDate(r, "delivery_date")
                        If Not newDelDate.HasValue Then Continue While
                        Dim newDueDate = GetWorkingDueDateCached(newDelDate.Value, dateCache)

                        Dim target = rows.FirstOrDefault(Function(x) x.InfoTypeCode = "6624" AndAlso x.CustomerOrderNo IsNot Nothing AndAlso x.CustomerOrderNo.Trim() = custOrderNo)
                        If target IsNot Nothing Then
                            target.DueDate = newDueDate
                            target.PreDailyDeliveryDate = newDelDate.Value
                        Else
                            Using cmdOrd As New OracleCommand("SELECT * FROM orders WHERE customer_code = '5455' AND info_type_code = '6624' AND TRIM(customer_order_no) = :p_ord_no AND active_flag = 'Y'", conn)
                                cmdOrd.Transaction = tran
                                cmdOrd.BindByName = True
                                cmdOrd.Parameters.Add(":p_ord_no", OracleDbType.Varchar2).Value = custOrderNo
                                Using ordR = cmdOrd.ExecuteReader()
                                    If ordR.Read() Then
                                        Dim fallbackRow = CreateOrdersStageRowFromOrdersReader(ordR, userId, pgId)
                                        fallbackRow.DueDate = newDueDate
                                        fallbackRow.PreDailyDeliveryDate = newDelDate
                                        If impFileId > 0 Then
                                            fallbackRow.ImpFileStageId = impFileId
                                            fallbackRow.ImpFileId = impFileId
                                        ElseIf adjFileId.HasValue AndAlso adjFileId.Value > 0 Then
                                            fallbackRow.ImpFileStageId = adjFileId.Value
                                            fallbackRow.ImpFileId = adjFileId.Value
                                        End If
                                        If impRunId > 0 Then
                                            fallbackRow.ImpRunId = impRunId
                                        End If
                                        rows.Add(fallbackRow)
                                    End If
                                End Using
                            End Using
                        End If
                    End While
                End Using
            End Using
        End Sub

        Private Sub ApplyAdjustment663S(conn As OracleConnection, tran As OracleTransaction, impFileId As Long, impRunId As Long, rows As List(Of OrdersStageRow), userId As String, pgId As String, Optional outAdjustmentStageIds As HashSet(Of Long) = Nothing)
            Dim sql = "SELECT * FROM SUZUKI_SPIRITS_663NAND663SAND664T WHERE info_type_code = '663S' AND active_flag = 'Y' AND (status IS NULL OR status <> 'REGISTERED')"
            If impFileId > 0 Then sql &= " AND imp_file_id = :p_imp_file_id"
            If impRunId > 0 Then sql &= " AND imp_run_id = :p_imp_run_id"

            Using cmd As New OracleCommand(sql, conn)
                cmd.Transaction = tran
                cmd.BindByName = True
                If impFileId > 0 Then cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileId
                If impRunId > 0 Then cmd.Parameters.Add(":p_imp_run_id", OracleDbType.Int64).Value = impRunId

                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim adjFileId = GetSafeInt64(r, "imp_file_id")
                        If adjFileId.HasValue AndAlso adjFileId.Value > 0 AndAlso outAdjustmentStageIds IsNot Nothing Then
                            outAdjustmentStageIds.Add(adjFileId.Value)
                        End If

                        Dim custOrderNo = If(r.IsDBNull(r.GetOrdinal("customer_order_no")), "", r("customer_order_no").ToString().Trim())
                        Dim newQty = If(GetSafeInt64(r, "order_qty"), 0L)

                        Dim target = rows.FirstOrDefault(Function(x) x.InfoTypeCode = "6624" AndAlso x.CustomerOrderNo IsNot Nothing AndAlso x.CustomerOrderNo.Trim() = custOrderNo)
                        If target IsNot Nothing Then
                            target.DemandQty = newQty
                            target.PreDailyOrderQty = Convert.ToDecimal(newQty)
                        Else
                            Using cmdOrd As New OracleCommand("SELECT * FROM orders WHERE customer_code = '5455' AND info_type_code = '6624' AND TRIM(customer_order_no) = :p_ord_no AND active_flag = 'Y'", conn)
                                cmdOrd.Transaction = tran
                                cmdOrd.BindByName = True
                                cmdOrd.Parameters.Add(":p_ord_no", OracleDbType.Varchar2).Value = custOrderNo
                                Using ordR = cmdOrd.ExecuteReader()
                                    If ordR.Read() Then
                                        Dim fallbackRow = CreateOrdersStageRowFromOrdersReader(ordR, userId, pgId)
                                        fallbackRow.DemandQty = newQty
                                        fallbackRow.PreDailyOrderQty = Convert.ToDecimal(newQty)
                                        If impFileId > 0 Then
                                            fallbackRow.ImpFileStageId = impFileId
                                            fallbackRow.ImpFileId = impFileId
                                        ElseIf adjFileId.HasValue AndAlso adjFileId.Value > 0 Then
                                            fallbackRow.ImpFileStageId = adjFileId.Value
                                            fallbackRow.ImpFileId = adjFileId.Value
                                        End If
                                        If impRunId > 0 Then
                                            fallbackRow.ImpRunId = impRunId
                                        End If
                                        rows.Add(fallbackRow)
                                    End If
                                End Using
                            End Using
                        End If
                    End While
                End Using
            End Using
        End Sub

        Private Function CreateOrdersStageRowFromOrdersReader(r As IDataRecord, userId As String, pgId As String) As OrdersStageRow
            Dim row As New OrdersStageRow()
            row.OrderId = GetSafeInt64(r, "order_id")
            row.CustomerSettingId = If(GetSafeInt64(r, "customer_setting_id"), 0L)
            row.ImpFileStageId = If(GetSafeInt64(r, "imp_file_id"), 0L)
            row.ImpFileId = row.ImpFileStageId
            row.ImpRunId = If(GetSafeInt64(r, "imp_run_id"), 0L)
            row.CustomerCode = GetSafeString(r, "customer_code")
            row.BillingTo = GetSafeString(r, "billing_to")
            row.CustomerOrderNo = GetSafeString(r, "customer_order_no")
            row.DemandStatus = GetSafeString(r, "demand_status")
            row.ShipTo = GetSafeString(r, "ship_to")
            row.CustomerItemNo = GetSafeString(r, "customer_item_no")
            row.ItemNo = GetSafeString(r, "item_no")
            row.DemandQty = GetSafeInt64(r, "demand_qty")
            row.DemandUnit = GetSafeString(r, "demand_unit")
            row.CurrencyCode = GetSafeString(r, "currency_code")
            row.ShipStockLocation = GetSafeString(r, "ship_stock_location")
            row.CompanyId = GetSafeString(r, "company_id")
            row.ProductCode = GetSafeString(r, "product_code")
            row.BillingStandard = GetSafeString(r, "billing_standard")
            row.ShipProcessType = GetSafeString(r, "ship_process_type")
            row.DeliveryInstrFlag = GetSafeString(r, "delivery_instr_flag")
            row.TransportMethod = GetSafeString(r, "transport_method")
            Dim oType = GetSafeInt64(r, "order_type")
            row.OrderType = If(oType.HasValue, CType(oType.Value, Short?), Nothing)
            Dim pType = GetSafeInt64(r, "prorated_type")
            row.ProratedType = If(pType.HasValue, CType(pType.Value, Short?), Nothing)
            row.CustomerInfoType = GetSafeString(r, "customer_info_type")
            row.InfoType = GetSafeString(r, "info_type")
            row.SelfFcstFlag = GetSafeString(r, "self_fcst_flag")
            row.SelfFcstDeleteFlag = GetSafeString(r, "self_fcst_delete_flag")
            row.OrderDate = GetSafeDate(r, "order_date")
            row.DueDate = GetSafeDate(r, "due_date")
            row.PreDailyDeliveryDate = GetSafeDate(r, "pre_daily_delivery_date")
            row.PreDailyOrderQty = GetSafeDecimal(r, "pre_daily_order_qty")
            row.DeliveryTime = GetSafeDecimal(r, "delivery_time")
            row.ContainerCapacity = GetSafeDecimal(r, "container_capacity")
            row.InitialFlag = GetSafeString(r, "initial_flag")
            row.TargetReferenceDateType = GetSafeString(r, "target_reference_date_type")
            row.TargetReferenceDate = GetSafeString(r, "target_reference_date")
            row.InfoTypeCode = GetSafeString(r, "info_type_code")
            row.DeliveryType = GetSafeString(r, "delivery_type")
            row.ProductionMonthType = GetSafeString(r, "production_month_type")
            row.Status = "PROCESSED"
            row.ActiveFlag = "Y"
            row.CreatedAt = DateTime.Now
            row.CreatedUserId = userId
            row.CreatedPgId = pgId
            row.UpdatedAt = DateTime.Now
            row.UpdatedUserId = userId
            row.UpdatedPgId = pgId
            Return row
        End Function

        Private Function GetSafeString(r As IDataRecord, col As String) As String
            Try
                Dim idx = r.GetOrdinal(col)
                If r.IsDBNull(idx) Then Return ""
                Return r.GetString(idx).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Function GetSafeInt64(r As IDataRecord, col As String) As Long?
            Try
                Dim idx = r.GetOrdinal(col)
                If r.IsDBNull(idx) Then Return Nothing
                Return Convert.ToInt64(r.GetValue(idx))
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetSafeDecimal(r As IDataRecord, col As String) As Decimal?
            Try
                Dim idx = r.GetOrdinal(col)
                If r.IsDBNull(idx) Then Return Nothing
                Return Convert.ToDecimal(r.GetValue(idx))
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetSafeDate(r As IDataRecord, col As String) As Date?
            Try
                Dim idx = r.GetOrdinal(col)
                If r.IsDBNull(idx) Then Return Nothing
                Return Convert.ToDateTime(r.GetValue(idx))
            Catch
                Return Nothing
            End Try
        End Function

#End Region

#Region "検証 & 出荷先変換"

        ''' <summary>
        ''' 【Stage 2 工程5】エラーチェック
        ''' </summary>
        Private Sub ValidateStagedRows(rows As List(Of OrdersStageRow), defaultSettingId As Long, errors As List(Of String))
            For Each r In rows
                ' 1. 取引先設定IDチェック
                If r.CustomerSettingId <= 0 Then
                    errors.Add($"品目No: {r.ItemNo} の取引先設定IDが特定できません。")
                End If

                ' 2. 必須項目チェック
                If String.IsNullOrWhiteSpace(r.ItemNo) Then
                    errors.Add("品目Noが未設定です。")
                End If

                If Not r.DueDate.HasValue Then
                    errors.Add($"希望納期が不正です (品目No: {r.ItemNo})")
                End If

                If Not r.DemandQty.HasValue OrElse r.DemandQty.Value < 0 Then
                    errors.Add($"需要数が負数または不正です (品目No: {r.ItemNo})")
                End If

                If String.IsNullOrWhiteSpace(r.ShipTo) Then
                    errors.Add($"出荷先が未設定です (品目No: {r.ItemNo})")
                End If
            Next
        End Sub

        ''' <summary>
        ''' メモリ内出荷先変換 (画面実行時用)
        ''' </summary>
        Private Sub ApplyShipToConversionInMemory(conn As OracleConnection, tran As OracleTransaction, rows As List(Of OrdersStageRow))
            Dim convMap As New Dictionary(Of String, String)()
            Using cmd As New OracleCommand("SELECT delivery_code_order, delivery_code_plan FROM suzuki_spirits_conversion_mst WHERE active_flag = 'Y'", conn)
                cmd.Transaction = tran
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim ord = r("delivery_code_order").ToString().Trim()
                        Dim plan = r("delivery_code_plan").ToString().Trim()
                        convMap(ord) = plan
                    End While
                End Using
            End Using

            For Each row In rows
                If convMap.ContainsKey(row.ShipTo) Then
                    row.ShipTo = convMap(row.ShipTo)
                End If
            Next
        End Sub

#End Region

#Region "STRA連携・完了処理・ユーティリティ"

        ''' <summary>
        ''' 【Stage 2 工程8】STRA納期設定実行 (バッチモードのみ)
        ''' </summary>
        Private Sub ExecuteStraDeliveryDateSetting(conn As OracleConnection, committedSettingIds As List(Of Long), userId As String, logger As Logger)
            Try
                ' 納期設定対象の取引先設定 (5455の各設定) を取得
                Dim pcMap As New Dictionary(Of Long, String)()
                Using cmd As New OracleCommand("SELECT customer_setting_id, profit_center FROM customer_setting_mst WHERE customer_code = '5455' AND active_flag = 'Y'", conn)
                    Using r = cmd.ExecuteReader()
                        While r.Read()
                            Dim csId = Convert.ToInt64(r("customer_setting_id"))
                            Dim pc = If(r.IsDBNull(r.GetOrdinal("profit_center")), "", r("profit_center").ToString().Trim())
                            pcMap(csId) = pc
                        End While
                    End Using
                End Using

                ' 本番登録された distinctSettings を優先し、未指定時は全5455設定を対象にする
                Dim targetSettingIds As List(Of Long)
                If committedSettingIds IsNot Nothing AndAlso committedSettingIds.Count > 0 Then
                    targetSettingIds = committedSettingIds
                Else
                    targetSettingIds = pcMap.Keys.ToList()
                End If

                For Each csId In targetSettingIds
                    Dim pc As String = If(pcMap.ContainsKey(csId), pcMap(csId), "")
                    Using tran As OracleTransaction = conn.BeginTransaction()
                        Try
                            Dim straErr = _orderStageRepo.DeliveryDateSetting(conn, tran, "5455", pc, csId, DateTime.Now, userId)
                            If String.IsNullOrEmpty(straErr) Then
                                tran.Commit()
                                logger.Write($"[SUZUKI_IMPORT] STRA納期設定処理が正常に完了しました (取引先設定ID: {csId}, PC: '{pc}')")
                            Else
                                tran.Rollback()
                                logger.Write($"[SUZUKI_IMPORT] STRA納期設定処理警告 (取引先設定ID: {csId}, PC: '{pc}'): {straErr}")
                            End If
                        Catch ex As Exception
                            tran.Rollback()
                            logger.Write($"[SUZUKI_IMPORT] STRA納期設定実行エラー (取引先設定ID: {csId}, PC: '{pc}'): {ex.Message}")
                        End Try
                    End Using
                Next
            Catch ex As Exception
                logger.Write($"[SUZUKI_IMPORT] STRA納期設定処理で例外が発生しました: {ex.Message}")
            End Try
        End Sub

        Private Sub MoveWorkFilesToCompleted(conn As OracleConnection, tran As OracleTransaction, impRunId As Long, folderPath As String, userId As String, Optional folderType As Integer = 4, Optional targetFileNames As List(Of String) = Nothing)
            Try
                If targetFileNames Is Nothing OrElse targetFileNames.Count = 0 Then Return

                Dim workRoot = ConfigurationManager.AppSettings("WorkFolderRoot")
                If String.IsNullOrWhiteSpace(workRoot) Then workRoot = "C:\ASTI\DATA\WORK"
                Dim completedRoot = ConfigurationManager.AppSettings("CompletedFolderRoot")
                If String.IsNullOrWhiteSpace(completedRoot) Then completedRoot = "C:\ASTI\DATA\COMPLETED"

                Dim userWorkDir = Path.Combine(workRoot, userId, "5455", folderType.ToString())
                Dim userCompletedDir = Path.Combine(completedRoot, userId, "5455", folderType.ToString())
                Utils.EnsureDirectory(userCompletedDir)

                If Directory.Exists(userWorkDir) Then
                    For Each fileName In targetFileNames
                        Dim srcPath = Path.Combine(userWorkDir, fileName)
                        If File.Exists(srcPath) Then
                            Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                            Dim ext = Path.GetExtension(fileName)
                            Dim destName = $"{nameNoExt}_{userId}_{DateTime.Now:yyyyMMddHHmmss}{ext}"
                            Dim destPath = Path.Combine(userCompletedDir, destName)

                            Try
                                File.Move(srcPath, destPath)
                            Catch
                            End Try
                        End If
                    Next
                End If
            Catch
            End Try
        End Sub

        Private Sub ExportErrorsToCsv(folderPath As String, errors As List(Of String))
            Try
                If String.IsNullOrEmpty(folderPath) OrElse Not Directory.Exists(folderPath) Then Return
                Dim errDir = Path.Combine(folderPath, "エラーリスト")
                Utils.EnsureDirectory(errDir)

                Dim errFile = Path.Combine(errDir, $"エラーリスト_{DateTime.Now:yyyyMMddHHmmss}.csv")
                Dim sb As New StringBuilder()
                sb.AppendLine("エラー内容")
                For Each errStr In errors
                    sb.AppendLine($"""{errStr.Replace("""", """""")}""")
                Next
                File.WriteAllText(errFile, sb.ToString(), Encoding.UTF8)
            Catch
            End Try
        End Sub

        Private Function GetWorkingDueDateCached(d As Date, cache As Dictionary(Of Date, Date)) As Date
            If cache.ContainsKey(d.Date) Then Return cache(d.Date)
            Dim res = _calender.AddWorkingDays2("00001", d, 0)
            cache(d.Date) = res
            Return res
        End Function

        Private Function GetUnitAndWhCached(itemNo As String, cache As Dictionary(Of String, (DemandUnit As String, ShipWh As String))) As (DemandUnit As String, ShipWh As String)
            If cache.ContainsKey(itemNo) Then Return cache(itemNo)
            Dim unit As String = ""
            Dim wh As String = ""
            Dim errMsg As String = ""
            _orderStageRepo.GetDemandUnitAndShipStockLocation(itemNo, unit, wh, errMsg)
            Dim res = (unit, wh)
            cache(itemNo) = res
            Return res
        End Function

        Private Function GetFolderPath(customerSettingId As Long, folderType As Integer) As String
            Try
                Using conn As New OracleConnection(_connectionString)
                    conn.Open()
                    Using cmd As New OracleCommand("SELECT folder_path FROM folder_mst WHERE customer_setting_id = :p_id AND folder_type = :p_type AND active_flag = 'Y'", conn)
                        cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = customerSettingId
                        cmd.Parameters.Add(":p_type", OracleDbType.Int32).Value = folderType
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                            Return res.ToString().Trim()
                        End If
                    End Using
                End Using
            Catch
            End Try
            Return ""
        End Function

        Private Function GetRuleFlag(conn As OracleConnection, tran As OracleTransaction, csId As Long, folderType As Integer, colName As String) As String
            Try
                Using cmd As New OracleCommand($"SELECT {colName} FROM imp_rule_mst WHERE customer_setting_id = :p_id AND folder_type = :p_type AND active_flag = 'Y'", conn)
                    cmd.Transaction = tran
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = csId
                    cmd.Parameters.Add(":p_type", OracleDbType.Int32).Value = folderType
                    Dim res = cmd.ExecuteScalar()
                    If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                        Return res.ToString().Trim()
                    End If
                End Using
            Catch
            End Try
            Return "N"
        End Function

        Private Function GetRuleReconcileType(conn As OracleConnection, tran As OracleTransaction, csId As Long, folderType As Integer) As Integer
            Try
                Using cmd As New OracleCommand("SELECT reconcile_type FROM imp_rule_mst WHERE customer_setting_id = :p_id AND folder_type = :p_type AND active_flag = 'Y'", conn)
                    cmd.Transaction = tran
                    cmd.Parameters.Add(":p_id", OracleDbType.Int64).Value = csId
                    cmd.Parameters.Add(":p_type", OracleDbType.Int32).Value = folderType
                    Dim res = cmd.ExecuteScalar()
                    If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                        Return Convert.ToInt32(res)
                    End If
                End Using
            Catch
            End Try
            Return 1
        End Function

#End Region

    End Class

End Namespace

