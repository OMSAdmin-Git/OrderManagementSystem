Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.SqlServer
Imports Microsoft.VisualBasic.ApplicationServices
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data
    Public Class OrderStageImport

        'Private _folderRepo As FolderRepository

        Public Property IsValid As Boolean = False
        Public Property ErrorMessage As String = ""
        Public Property InsertedCount As Integer = 0
        ' 検証成功時に呼び出し元で利用するパース済みの値
        Public Property CustomerSettingId As Long
        Public Property ImpFileStageId As Long
        Public Property FolderType As Integer
        Public Property CustomerCode As Integer
        Public Property ReconcileFlag As String = ""
        Public Property ReconcileType As String = ""
        Public Property FcstReconcileFlag As String = ""
        Public Property WorkFolderPath As String = ""
        Public Property StagedFileName As String = ""
        Public Property UserId As String = ""
        Public Property PgId As String = ""


        '/// ヤマハ取込データ保存配列
        Private Structure ImportDataType
            Dim siyosha As String           '使用者
            Dim status As String            '品目ステータス
            Dim customeritemNo As String    '旧体系部品番号(客先品目No)
            Dim nonyuplat As String         '納入プラットフォーム
            Dim yokisyuuyousuu As String    '容器収容数
            Dim yokibangou As String        '容器番号
            Dim henkoukubun As String        '変更区分
            Dim datakubun As String          'データ区分
            Dim ordersikibetuNo As String   'オーダー識別番号
            Dim nonyusijibi As String       '納入指示日
            Dim nonyusibijikan As String    '納入時間
            Dim nonyusijisu As String       '納入指示数
            Dim cardkubun As String         'カード区分
            Dim naijikubun As String        '内示区分
            Dim icdenpyoNo As String        'IC伝票No
            Dim nohinshoNo As String        '納品書番号
            Dim hinmokugyoNo As String      '品目情報行番号
            Dim ordergyoNo As String        'オーダー情報行番号
        End Structure
        Private m_ImpData() As ImportDataType


        Public Shared Function Orders_Saved(ByVal tran As OracleTransaction,
                                        ByVal idx As Integer,
                                        ByVal csidRaw As Object,
                                        ByVal ifsidRaw As Object,
                                        ByVal folderTypeRaw As Object,
                                        ByVal customerCodeRaw As Object,
                                        ByVal reconcileFlagRaw As Object,
                                        ByVal reconcileTypeRaw As Object,
                                        ByVal fcstReconcileFlagRaw As Object,
                                        ByVal stagedFolderPathRaw As Object,
                                        ByVal stagedFileNameRaw As Object,
                                        ByVal userIdRaw As Object,
                                        ByVal pgIdRaw As Object,
                                        ByVal blnHandFlag As Boolean,
                                        rowsForTemp2 As List(Of OrdersStageRow)
                                            ) As OrderStageImport

            Dim _oderStageRepo As New OrderStageRepository(Utils.GetConnectionString())
            Dim _impFileStageRepo As New ImpFilesStageRepository(Utils.GetConnectionString())

            Dim result As New OrderStageImport()

            Dim cnt As Integer = 0

            '-----------------------------------------------
            '取込ファイルの内容を受注ワーク登録
            cnt = _oderStageRepo.InsertRange(tran, rowsForTemp2)
            '-----------------------------------------------

            result.InsertedCount = cnt

            'If cnt > 0 Then
            If result.InsertedCount > 0 Then


                '-----------------------------------------------
                'IMP_FILE_STAGEを更新
                Dim strHandFlag As String = ""
                'If chkHandFlag.Checked = True Then
                If blnHandFlag = True Then
                    strHandFlag = "Y"
                Else
                    strHandFlag = "N"
                End If

                Dim rowsForTemp3 = New List(Of ImpFilesStageRow) From {
                                                        New ImpFilesStageRow With {
                                                            .ImpFileStageId = result.ImpFileStageId,
                                                            .HandFlag = strHandFlag,
                                                            .Status = "PARSED",
                                                            .UpdatedAt = Now,
                                                            .UpdatedUserId = result.UserId,
                                                            .UpdatedPgId = result.PgId
                                                            }
                                                        }

                _impFileStageRepo.UpdateRange(tran, rowsForTemp3)
                '-----------------------------------------------


                'デバック用
                'tran.Commit()



                '-----------------------------------------------
                rowsForTemp2.Clear()

                'ORDER_STAGEにORDERSのレコード追加
                rowsForTemp2 = New List(Of OrdersStageRow) From {
                                        New OrdersStageRow With {
                                            .CustomerSettingId = result.CustomerSettingId
                                        }
                                    }
                _oderStageRepo.InsertStageFromOrders(tran, result.CustomerSettingId)
                '-----------------------------------------------



                '-----------------------------------------------
                '--------
                '内示加工
                '--------
                '今回取込した内示データの抽出
                Dim dtNaiji As DataTable = _oderStageRepo.GetNaijiData(tran, result.ImpFileStageId)

                '今回取込した内示データの件数をチェック
                If dtNaiji.Rows.Count > 0 Then

                    '内示洗い替え
                    _oderStageRepo.ReplaceNaijiRelation(tran, result.ImpFileStageId, result.CustomerSettingId, Now, result.UserId, result.PgId)

                    'ステータス更新
                    _oderStageRepo.UpdateNaijiStatusProcessed(tran, result.ImpFileStageId)



                    '2026/05/26 酒井 フェーズ2 受注残対応
                    '受注残加工
                    '消込フラグをチェック
                    If result.ReconcileFlag = "Y" Then

                        '--------------------------------
                        '受注残消込
                        '--------------------------------
                        _oderStageRepo.BacklogForecast(tran,
                                                                            result.CustomerSettingId,
                                                                            result.ImpFileStageId,
                                                                            result.ReconcileType,
                                                                            Now,
                                                                            result.UserId,
                                                                            result.PgId)

                    End If
                    '--



                End If
                '-----------------------------------------------






                '-----------------------------------------------
                '--------
                '確定加工
                '--------

                '打切処理
                _oderStageRepo.UpdateClese(tran, result.ImpFileStageId, 2)

                '取消処理
                _oderStageRepo.UpdateCancel(tran, result.ImpFileStageId, 2)


                '確定データ無効化
                _oderStageRepo.ReplaceKakuteiRelation(tran,
                                                                            result.CustomerSettingId,
                                                                            result.ImpFileStageId,
                                                                            Now,
                                                                            result.UserId,
                                                                            result.PgId)

                'デバック用
                'tran.Commit()
                'Exit Sub

                '消込フラグをチェック
                If result.ReconcileFlag = "Y" Then

                    '確定の抽出
                    Dim dtKakutei As DataTable = _oderStageRepo.GetKakuteiData(tran, result.ImpFileStageId)

                    If dtKakutei.Rows.Count > 0 Then

                        '--------------------------------
                        '確定データで内示消込
                        '--------------------------------
                        _oderStageRepo.ReconcileForecast(tran,
                                                                                result.CustomerSettingId,
                                                                                result.ImpFileStageId,
                                                                                2,
                                                                                result.ReconcileType,
                                                                                Now,
                                                                                result.UserId,
                                                                                result.PgId)

                    End If

                End If

                '--------------------------------
                '確定 新規
                '--------------------------------
                _oderStageRepo.UpdateKakuteiNewOrders(tran, result.ImpFileStageId)

                '-----------------------------------------------





                '-----------------------------------------------
                '--------
                '納入指示加工
                '--------

                '打切処理
                _oderStageRepo.UpdateClese(tran, result.ImpFileStageId, 3)

                '取消処理
                _oderStageRepo.UpdateCancel(tran, result.ImpFileStageId, 3)

                '消込フラグをチェック
                If result.ReconcileFlag = "Y" Then


                    '--------------------------------
                    '確定消込 ※客先発注Noでの消込
                    '--------------------------------
                    '受注消込(客先発注No)
                    _oderStageRepo.ExecuteOrderReconciliationByOrderNo(tran,
                                                                                            result.CustomerSettingId,
                                                                                            result.ImpFileStageId,
                                                                                            Now,
                                                                                            result.UserId,
                                                                                            result.PgId)

                    '内示消込フラグをチェック
                    If result.FcstReconcileFlag = "Y" Then

                        '納入指示注文の抽出
                        Dim dtNonyuSiji As DataTable = _oderStageRepo.GetNonyuSijiData(tran, result.ImpFileStageId)

                        If dtNonyuSiji.Rows.Count > 0 Then

                            ''デバック用
                            'tran.Commit()
                            'Exit Sub

                            '--------------------------------
                            '納入指示データで内示消込
                            '--------------------------------
                            _oderStageRepo.ReconcileForecast(tran,
                                                                                    result.CustomerSettingId,
                                                                                    result.ImpFileStageId,
                                                                                    3,
                                                                                    result.ReconcileType,
                                                                                    Now,
                                                                                    result.UserId,
                                                                                    result.PgId)

                        End If

                    End If

                End If

                '--------------------------------
                '納入指示 新規
                '--------------------------------
                _oderStageRepo.UpdateNonyuSijiNewOrders(tran, result.ImpFileStageId)

                '-----------------------------------------------


                'If ErrFlg = False Then

                '    resultCnt += 1
                '    resultRowCnt += cnt

                '    successs.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　読込 {cnt} 件　異常 {errcnt} 件")
                '    'successs.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　読込 {cnt} 件")

                '    'errcnt = 0

                'End If


            Else

                '取込対象が一件も無い場合、
                'nodata.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　対象データが1件も無いため破棄してください。")

            End If

            result.IsValid = True
            Return result

        End Function

        ''' <summary>
        ''' 他のプログラムからも呼び出せる共通の検証メソッド
        ''' </summary>
        Public Shared Function ValidateFileStageRow(ByVal tran As OracleTransaction,
                                        ByVal idx As Integer,
                                        ByVal csidRaw As Object,
                                        ByVal ifsidRaw As Object,
                                        ByVal folderTypeRaw As Object,
                                        ByVal customerCodeRaw As Object,
                                        ByVal reconcileFlagRaw As Object,
                                        ByVal fcstReconcileFlagRaw As Object,
                                        ByVal stagedFolderPathRaw As Object,
                                        ByVal stagedFileNameRaw As Object
                                            ) As OrderImport


            Dim oderStageRepo As New OrderStageRepository(Utils.GetConnectionString())
            Dim mappingRepo As New MappingRepository(Utils.GetConnectionString())

            'LIST用
            Dim nKyakusakiHattyuNo As Integer = 0    '/// 客先発注No
            Dim nJutyuuBi As Integer = 0 '/// 受注日
            Dim nKibouNouki As Integer = 0    '/// 希望納期
            Dim nKyakusakiHinmokuNo As Integer = 0      '/// 客先品目No
            Dim njuyouSuu As Integer = 0   '/// 需要数
            Dim nTukaCode As Integer = 0  '/// 通貨コード
            Dim nSeihinCode As Integer = 0     '/// 製品コード
            Dim nNonyusakiCode As Integer = 0   '/// 納入先コード
            Dim nComment As Integer = 0   '/// コメント
            Dim nJutyuKubun As Integer = 0   '/// 受注区分
            Dim nBunkatuKubun As Integer = 0   '/// 分割区分
            Dim nTorihikisakiJohoKubun As Integer = 0   '/// 取引先情報区分
            Dim nJishaYosokuFlag As Integer = 0   '/// 自社予測フラグ
            Dim nJishaYosokuDelFlag As Integer = 0   '

            'MATRIX用
            Dim mKyakusakiHattyuNo As String = ""    '/// 客先発注No
            Dim mJutyuuBi As String = "" '/// 受注日
            Dim mKibouNouki As String = ""    '/// 希望納期
            Dim mKyakusakiHinmokuNo As String = ""      '/// 客先品目No
            Dim mjuyouSuu As String = ""   '/// 需要数
            Dim mTukaCode As String = ""  '/// 通貨コード
            Dim mSeihinCode As String = ""     '/// 製品コード
            Dim mNonyusakiCode As String = ""   '/// 納入先コード
            Dim mComment As String = ""   '/// コメント
            Dim mJutyuKubun As String = ""   '/// 受注区分
            Dim mBunkatuKubun As String = ""   '/// 分割区分
            Dim mTorihikisakiJohoKubun As String = ""  '/// 取引先情報区分
            Dim mJishaYosokuFlag As String = ""   '/// 自社予測フラグ
            Dim mJishaYosokuDelFlag As String = ""   '

            Dim strTempDate As String  '日付検証用
            Dim strQtyValue As String  '数値検証用

            '日付検証用
            Dim formats As String() = {
                                                "yyyy/MM/dd", "yyyy-MM-dd", "yyyyMMdd",
                                                "yy/MM/dd", "yy-MM-dd", "yyMMdd",
                                                "yyyy/M/d", "yyyy-M-d",
                                                "yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss",
                                                "yyyy/MM/dd H:mm:ss", "yyyy-MM-dd H:mm:ss"
                                            }

            Dim result As New OrderImport()
            Dim rowsForTemp2 As New List(Of OrdersStageRow)


            Dim strWorkFile As String
            Dim TorikomiFile As String  '取込ファイル保持用


            '取引先設定ID　GridViewから取得
            Dim custSettingId As Long = 0
            If csidRaw Is Nothing OrElse Not Long.TryParse(csidRaw.ToString(), custSettingId) Then
                result.ErrorMessage = $"Row {idx}：CustomerSettingIdが不正"
                Return result
            End If
            result.CustomerSettingId = custSettingId

            '一時取込ファイルID　GridViewから取得
            Dim impFileId As Long = 0
            If ifsidRaw Is Nothing OrElse Not Long.TryParse(ifsidRaw.ToString(), impFileId) Then
                result.ErrorMessage = $"Row {idx}：ImpFileStageIdが不正"
                Return result
            End If
            result.ImpFileStageId = impFileId

            'フォルダタイプ　GridViewから取得
            Dim fType As Integer = 0
            If folderTypeRaw Is Nothing OrElse Not Integer.TryParse(folderTypeRaw.ToString(), fType) Then
                result.ErrorMessage = $"Row {idx}：FolderTypeが不正"
                Return result
            End If
            result.FolderType = fType

            '取引先コード　GridViewから取得
            Dim cCode As Integer = 0
            If customerCodeRaw Is Nothing OrElse Not Integer.TryParse(customerCodeRaw.ToString(), cCode) Then
                result.ErrorMessage = $"Row {idx}：CustomerCodeが不正"
                Return result
            End If
            result.CustomerCode = cCode

            '消込フラグ
            Dim recFlag As String = ""
            If reconcileFlagRaw IsNot Nothing Then
                recFlag = reconcileFlagRaw.ToString().Trim().ToUpper()
            End If
            If recFlag <> "Y" AndAlso recFlag <> "N" Then
                result.ErrorMessage = $"Row {idx}：ReconcileFlagが不正"
                Return result
            End If
            result.ReconcileFlag = recFlag

            '内示消込フラグ
            Dim fcstRecFlag As String = ""
            If fcstReconcileFlagRaw IsNot Nothing Then
                fcstRecFlag = fcstReconcileFlagRaw.ToString().Trim().ToUpper()
            End If
            If fcstRecFlag <> "Y" AndAlso fcstRecFlag <> "N" Then
                result.ErrorMessage = $"Row {idx}：FcstReconcileFlagが不正"
                Return result
            End If
            result.FcstReconcileFlag = fcstRecFlag

            ' 7. WORKフォルダパスの検証
            If stagedFolderPathRaw IsNot Nothing Then
                Dim folderPath As String = stagedFolderPathRaw.ToString()
                If Not Directory.Exists(folderPath) Then
                    result.ErrorMessage = $"Row {idx}：WORKフォルダが存在しません"
                    Return result
                End If
                result.WorkFolderPath = folderPath
            Else
                result.ErrorMessage = $"Row {idx}：WORKフォルダパスが不正"
                Return result
            End If

            ' [WORKファイル名]を取得
            'Dim strWorkFile As String = keys("StagedFileName").ToString()
            strWorkFile = ""
            strWorkFile = stagedFileNameRaw.ToString()

            '取込ファイル表示用に退避
            TorikomiFile = ""
            TorikomiFile = strWorkFile

            If strWorkFile IsNot Nothing Then
                'Dim workFile As String = workFolder & "/" & strWorkFile
                strWorkFile = result.WorkFolderPath & "\" & strWorkFile
                ' ファイル存在確認
                If Not File.Exists(strWorkFile) Then
                    'errors.Add($"Row {idx}：WORKファイルが存在しません")
                    'Continue For
                    result.ErrorMessage = $"Row {idx}：WORKファイルが存在しません"
                    result.IsValid = False
                    Return result
                End If
            Else
                'errors.Add($"Row {idx}：WORKファイル名が不正")
                'Continue For
                result.ErrorMessage = $"Row {idx}：WORKファイル名が不正"
                result.IsValid = False
                Return result
            End If



            '-----------------
            '受注ワーク削除 ※customerSettingId、folderType単位で削除
            '-----------------
            oderStageRepo.DeleteRange(tran, result.CustomerSettingId, result.FolderType)




            '※マッピングプロファイルマスタ(MAPPINNG_PROFILE_MST)と
            '  マッピング明細マスタ(FIELD_MAPPINNG_MST)を参照してファイルデータを取得する
            Dim mappingInfos As List(Of MappingInfo) = mappingRepo.GetMappingInfo(result.CustomerSettingId, result.FolderType)
            If mappingInfos Is Nothing OrElse mappingInfos.Count = 0 Then
                'errors.Add($"{mappingInfos}：MAPPINNG_PROFILE_MSTに未登録")
                'Continue For
                result.ErrorMessage = $"{mappingInfos}：MAPPINNG_PROFILE_MSTに未登録"
                result.IsValid = False
                Return result
            End If

            'MAPPINNG_PROFILE_MST
            Dim HeaderRowIndex As Integer = 0
            Dim DataStartRowIndex As Integer = 0
            Dim DefaultSheetName As String = ""
            'FILE_MST
            Dim FomatType As String = ""
            Dim FileType As String = ""
            Dim Delimiter As Char = ","c
            Dim Enclosure As String = ""
            Dim HeaderFlag As String = ""
            'Dim LineEnding As String = ""
            Dim CharSet As String = ""
            'FIELD_MAPPING_MST
            Dim TargetField As String = ""
            Dim SourceColumnIndex As Integer = 0
            Dim SourceCellAddress As String = ""

            Dim FstFlg As Boolean = True

            For Each info In mappingInfos

                HeaderRowIndex = info.HeaderRowIndex
                DataStartRowIndex = info.DataStartRowIndex
                DefaultSheetName = info.default_sheet_name  'デフォルトシート名

                FomatType = info.format_type    'LIST/MATRIX
                FileType = info.file_type       'CSV/TSV/FIXED/EXCEL
                CharSet = info.charset          'UTF8/SJIS
                HeaderFlag = info.header_flag   'N/Y
                Enclosure = info.enclosure      'D_QUOTE/S_QUOTE
                'LineEnding = info.line_ending   'CRLF/LF

                '区切り文字
                Select Case info.delimiter
                    Case "COMMA"
                        Delimiter = ","c
                    Case "TAB"
                        Delimiter = vbTab
                    Case "SEMICOLON"
                        Delimiter = ";"c
                    Case "PIPE"
                        Delimiter = "|"c
                    Case "SPACE"
                        Delimiter = " "c
                    Case "COLON"
                        Delimiter = ":"c
                    Case Else
                        Delimiter = ","c
                End Select

                '囲い文字
                Select Case info.enclosure
                    Case "D_QUOTE"
                        Enclosure = """"c
                    Case "S_QUOTE"
                        Enclosure = "'"c
                    Case Else
                        Enclosure = ""
                End Select

                'マッピング先項目名
                TargetField = info.target_field

                'LIST用
                If FileType = "EXCEL" Then

                    'EXCELは開始列が1からのためそのまま
                    SourceColumnIndex = info.source_column_index    '列番号
                Else
                    'CSV、TSVは開始列が0からのため-1
                    SourceColumnIndex = info.source_column_index - 1   '列番号

                    If FstFlg = True Then

                        nKyakusakiHattyuNo = -1
                        nJutyuuBi = -1
                        nKibouNouki = -1
                        nKyakusakiHinmokuNo = -1
                        njuyouSuu = -1
                        nTukaCode = -1
                        nSeihinCode = -1
                        nNonyusakiCode = -1
                        nComment = -1
                        nJutyuKubun = -1
                        nBunkatuKubun = -1
                        nTorihikisakiJohoKubun = -1
                        nJishaYosokuFlag = -1
                        nJishaYosokuDelFlag = -1

                        FstFlg = False

                    End If

                End If

                'MATRIX用
                SourceCellAddress = info.source_cell_address    'EXCELセルアドレス

                '取得する項目の位置
                Select Case TargetField
                    Case "CUSTOMER_ORDER_NO"
                        nKyakusakiHattyuNo = SourceColumnIndex
                        mKyakusakiHattyuNo = SourceCellAddress
                    Case "ORDER_DATE"
                        nJutyuuBi = SourceColumnIndex
                        mJutyuuBi = SourceCellAddress
                    Case "DUE_DATE"
                        nKibouNouki = SourceColumnIndex
                        mKibouNouki = SourceCellAddress
                    Case "CUSTOMER_ITEM_NO"
                        nKyakusakiHinmokuNo = SourceColumnIndex
                        mKyakusakiHinmokuNo = SourceCellAddress
                    Case "DEMAND_QTY"
                        njuyouSuu = SourceColumnIndex
                        mjuyouSuu = SourceCellAddress
                    Case "CURRENCY_CODE"
                        nTukaCode = SourceColumnIndex
                        mTukaCode = SourceCellAddress
                    Case "PRODUCT_CODE"
                        nSeihinCode = SourceColumnIndex
                        mSeihinCode = SourceCellAddress
                    Case "REMARKS"
                        nComment = SourceColumnIndex
                        mComment = SourceCellAddress
                    Case "DELIVERY_CODE"
                        nNonyusakiCode = SourceColumnIndex
                        mNonyusakiCode = SourceCellAddress
                    Case "ORDER_TYPE"
                        nJutyuKubun = SourceColumnIndex
                        mJutyuKubun = SourceCellAddress
                    Case "PRORATED_TYPE"
                        nBunkatuKubun = SourceColumnIndex
                        mBunkatuKubun = SourceCellAddress
                    Case "CUSTOMER_INFO_TYPE"
                        nTorihikisakiJohoKubun = SourceColumnIndex
                        mTorihikisakiJohoKubun = SourceCellAddress
                    Case "SELF_FCST_FLAG"
                        nJishaYosokuFlag = SourceColumnIndex
                        mJishaYosokuFlag = SourceCellAddress
                    Case "SELF_FCST_DELETE_FLAG"
                        nJishaYosokuDelFlag = SourceColumnIndex
                        mJishaYosokuDelFlag = SourceCellAddress
                End Select

            Next

            'ファイル内の行インデックス
            Dim fileidx As Integer = 0


            'ファイル読み込み
            Select Case FileType
                Case "CSV", "TSV"

                Case "FIXED"
                    '3:FIXED(固定長)
                    'フェーズ2で実装

                    '許可する拡張子のリスト（小文字で定義）
                    Dim allowedExtensions As New List(Of String) From {".txt"}

                    'ファイルパスから拡張子を取得
                    Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                    '拡張子チェック
                    If Not allowedExtensions.Contains(fileExtension) Then
                        'errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                        'Continue For
                        result.ErrorMessage = $" 取引先コード：{result.CustomerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。"
                        result.IsValid = False
                        Return result
                    End If

                    'ヤマハ取込データ保存配列　初期化
                    Dim m_ImpData(0) As ImportDataType

                    Using StmRdr As New IO.StreamReader(strWorkFile, MapEncoding(CharSet))

                        Dim isFirstRow As Boolean = True ' 初回の要素追加判定用

                        ' 品目情報をループ内で保持するための退避用変数
                        Dim currentYokibangou As String = ""
                        Dim currentYokisyuuyousuu As String = ""
                        Dim currentCustomeritemNo As String = ""
                        Dim currentStatus As String = ""
                        Dim currentNonyuplat As String = ""
                        Dim currentSiyosha As String = ""
                        Dim currentHinmokugyoNo As String = ""

                        fileidx = 0

                        '文字開始位置、文字数を指定する場合
                        While Not StmRdr.EndOfStream

                            Dim StrLine As String = StmRdr.ReadLine()
                            fileidx += 1

                            If Left(StrLine, 4) = "HEAD" Then
                                'ヘッダ情報

                            ElseIf Left(StrLine, 4) = "TRAL" Then
                                'EOF情報

                            ElseIf Left(StrLine, 4) = "LE01" Then
                                '品目情報

                                '品目情報を取込データから取得
                                currentSiyosha = StrLine.Substring(20 - 1, 4).Trim()          '使用者
                                currentStatus = StrLine.Substring(24 - 1, 1).Trim()           '品目ステータス
                                currentCustomeritemNo = StrLine.Substring(55 - 1, 14).Trim()  '旧体系部品番号(客先品目No)
                                currentNonyuplat = StrLine.Substring(78 - 1, 4).Trim()        '納入プラットフォーム
                                currentYokisyuuyousuu = StrLine.Substring(87 - 1, 5).Trim()   '容器収容数
                                currentYokibangou = StrLine.Substring(92 - 1, 5).Trim()       '容器番号
                                currentHinmokugyoNo = fileidx

                                '品目情報の2行目は今のところ使用しない予定
                                If Not StmRdr.EndOfStream Then
                                    StmRdr.ReadLine()
                                    fileidx += 1 ' 行数カウントの整合性を保つためにインクリメント
                                End If

                            Else
                                'オーダ情報（1行の中にデータ1とデータ2が最大2つ格納されている）

                                ' --------------------------------------------------
                                ' 1つ目のオーダ情報データの処理
                                ' --------------------------------------------------
                                ' 変更区分1（データ1のキー項目など）が空でないか確認
                                Dim HenkouKubun1 As String = StrLine.Substring(1 - 1, 1).Trim()

                                If HenkouKubun1 <> "" Then

                                    ' 初回のみ ReDim m_ImpData(0) をそのまま使い、2回目以降は配列を拡張する
                                    If isFirstRow Then
                                        isFirstRow = False
                                    Else
                                        ReDim Preserve m_ImpData(UBound(m_ImpData) + 1)
                                    End If

                                    '品目情報をセット
                                    m_ImpData(UBound(m_ImpData)).siyosha = currentSiyosha                                   '使用者
                                    m_ImpData(UBound(m_ImpData)).status = currentStatus                                     '品目ステータス
                                    m_ImpData(UBound(m_ImpData)).customeritemNo = currentCustomeritemNo                     '旧体系部品番号(客先品目No)
                                    m_ImpData(UBound(m_ImpData)).nonyuplat = currentNonyuplat                               '納入プラットフォーム
                                    m_ImpData(UBound(m_ImpData)).yokisyuuyousuu = currentYokisyuuyousuu                     '容器収容数
                                    m_ImpData(UBound(m_ImpData)).yokibangou = currentYokibangou                             '容器番号

                                    'オーダ情報を取込データから取得
                                    m_ImpData(UBound(m_ImpData)).ordersikibetuNo = StrLine.Substring(3 - 1, 5).Trim()      'オーダー識別番号1
                                    m_ImpData(UBound(m_ImpData)).nonyusijibi = StrLine.Substring(10 - 1, 8).Trim()         '納入指示日1
                                    m_ImpData(UBound(m_ImpData)).nonyusibijikan = StrLine.Substring(18 - 1, 4).Trim()      '納入時間1
                                    m_ImpData(UBound(m_ImpData)).nonyusijisu = StrLine.Substring(22 - 1, 6).Trim()         '納入指示数1
                                    m_ImpData(UBound(m_ImpData)).cardkubun = StrLine.Substring(29 - 1, 1).Trim()           'カード区分1
                                    m_ImpData(UBound(m_ImpData)).naijikubun = StrLine.Substring(30 - 1, 1).Trim()          '内示区分1
                                    m_ImpData(UBound(m_ImpData)).icdenpyoNo = StrLine.Substring(33 - 1, 5).Trim()          'IC伝票No1
                                    m_ImpData(UBound(m_ImpData)).nohinshoNo = StrLine.Substring(60 - 1, 15).Trim()         '納品書番号1

                                    m_ImpData(UBound(m_ImpData)).hinmokugyoNo = currentHinmokugyoNo                        '品目情報行番号
                                    m_ImpData(UBound(m_ImpData)).ordergyoNo = fileidx                                      'オーダー情報行番号

                                End If

                                ' --------------------------------------------------
                                ' 2つ目のオーダデータの処理（データが存在する場合のみ格納）
                                ' --------------------------------------------------
                                ' 変更区分2が空文字でない、かつSubstringできる長さがあるか確認
                                Dim HenkouKubun2 As String = ""
                                If StrLine.Length >= 88 Then ' Substring(87 - 1, 1)がエラーにならない長さチェック
                                    HenkouKubun2 = StrLine.Substring(87 - 1, 1).Trim()
                                End If

                                If HenkouKubun2 <> "" Then

                                    ' データ2用に必ず配列を新しく1枠拡張する
                                    If isFirstRow Then
                                        isFirstRow = False
                                    Else
                                        ReDim Preserve m_ImpData(UBound(m_ImpData) + 1)
                                    End If

                                    '品目情報をセット
                                    m_ImpData(UBound(m_ImpData)).siyosha = currentSiyosha                                   '使用者
                                    m_ImpData(UBound(m_ImpData)).status = currentStatus                                     '品目ステータス
                                    m_ImpData(UBound(m_ImpData)).customeritemNo = currentCustomeritemNo                     '旧体系部品番号(客先品目No)
                                    m_ImpData(UBound(m_ImpData)).nonyuplat = currentNonyuplat                               '納入プラットフォーム
                                    m_ImpData(UBound(m_ImpData)).yokisyuuyousuu = currentYokisyuuyousuu                     '容器収容数
                                    m_ImpData(UBound(m_ImpData)).yokibangou = currentYokibangou                             '容器番号

                                    m_ImpData(UBound(m_ImpData)).ordersikibetuNo = StrLine.Substring(89 - 1, 5).Trim()     'オーダー識別番号2
                                    m_ImpData(UBound(m_ImpData)).nonyusijibi = StrLine.Substring(96 - 1, 8).Trim()         '納入指示日2
                                    m_ImpData(UBound(m_ImpData)).nonyusibijikan = StrLine.Substring(104 - 1, 4).Trim()     '納入時間2
                                    m_ImpData(UBound(m_ImpData)).nonyusijisu = StrLine.Substring(108 - 1, 6).Trim()        '納入指示数2
                                    m_ImpData(UBound(m_ImpData)).cardkubun = StrLine.Substring(115 - 1, 1).Trim()          'カード区分2
                                    m_ImpData(UBound(m_ImpData)).naijikubun = StrLine.Substring(116 - 1, 1).Trim()         '内示区分2
                                    m_ImpData(UBound(m_ImpData)).icdenpyoNo = StrLine.Substring(119 - 1, 5).Trim()         'IC伝票No2
                                    m_ImpData(UBound(m_ImpData)).nohinshoNo = StrLine.Substring(146 - 1, 15).Trim()        '納品書番号2

                                    m_ImpData(UBound(m_ImpData)).hinmokugyoNo = currentHinmokugyoNo                        '品目情報行番号
                                    m_ImpData(UBound(m_ImpData)).ordergyoNo = fileidx                                      'オーダー情報行番号

                                End If




                            End If

                        End While

                    End Using

                    'デバック用に取込したデータをワークテーブルに保存
                    SaveImportDataToWorkTable(tran, m_ImpData, result.ImpFileStageId)

                    '-----------------
                    '取込データを加工
                    '-----------------

                    fileidx = 1



                    'For IntDataCnt As Integer = 0 To UBound(m_ImpData)

                    '    ''テスト用に変数宣言
                    '    Dim seisankubun As String = ""      'JU0580(生産区分)
                    '    Dim syohinkubun As String = ""      'JU0920(初品区分)
                    '    Dim siyosaki As String = ""         'JU0240(使用先)

                    '    '初期化
                    '    strTempDate = ""  '日付検証用
                    '    strQtyValue = ""  '数値検証用
                    '    customerorderNo = ""
                    '    orderDate = Nothing
                    '    DueDate = Nothing
                    '    customeritemNo = ""
                    '    demandqty = 0
                    '    demandunit = ""
                    '    currencycode = ""
                    '    productcode = ""
                    '    remarks = ""
                    '    deliverycode = ""
                    '    predailyorderqty = 0
                    '    predailydeliveryDate = Nothing
                    '    ordertype = 0
                    '    proratedtype = 1
                    '    customerinfotype = ""
                    '    selffcstflag = ""
                    '    selffcstdeleteflag = ""
                    '    shipto = ""
                    '    billingto = ""
                    '    itemNo = ""
                    '    demandstatus = ""
                    '    shipprocesstype = ""
                    '    deliveryinstrflag = ""
                    '    totalshipqty = 0
                    '    shipstocklocation = ""
                    '    infotype = ""
                    '    reconciletype = 1
                    '    profitcenterCSM = ""
                    '    profitcenter = ""

                    '    errMsg = ""

                    '    ErrFlg = False



                    '    '取引先設定IDのPC   （必須）
                    '    'CUSTOMER_SETTING_MSTより取得
                    '    profitcenterCSM = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If

                    '    '客先品目No　品目No検索 (客先品目Noにハイフォンをつけて検索)
                    '    'STRAMMIC.PRDSLSODRMより取得
                    '    'Dim Patterns As New List(Of String) From {"3-5-2", "5-5"}
                    '    Dim Patterns As New List(Of String)()

                    '    Select Case CustomerCode
                    '        Case "5384"
                    '            ' YMPC(ヤマハモーターパワープロダクツ)
                    '            Patterns.AddRange({"3-5-2", "5-5"})
                    '        Case "5952"
                    '            ' YEJP(ヤマハ発動機 遠州森町工場)
                    '            Patterns.AddRange({"3-5-2", "5-5-2-2", "4-5-2", "3-5-5", "3-11", "5-5"})
                    '        Case "5977"
                    '            ' YMC(ヤマハ発動機)
                    '            Patterns.AddRange({"3-5-2", "5-5"})
                    '    End Select

                    '    customeritemNo = m_ImpData(IntDataCnt).customeritemNo
                    '    itemNo = ""
                    '    errMsg = ""
                    '    ''デバック用
                    '    'If m_ImpData(IntDataCnt).status = 2 Then
                    '    '    Dim test As String
                    '    '    test = "test"
                    '    'End If
                    '    If _oderStageRepo.GetProductCode2(CustomerCode, customeritemNo, Patterns, m_ImpData(IntDataCnt).status, itemNo, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '        'ErrFlg = True
                    '        outputerrors.Add($"取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '    End If

                    '    '品目NoのPC   （必須）
                    '    'STRAMMIC.USRDEFFLDFより取得
                    '    profitcenter = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If

                    '    '取引先設定IDのPCと同じPCのみ取込対象とする
                    '    If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                    '        'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                    '        fileidx += 1
                    '        Continue For
                    '    End If



                    '    'フォルダタイプで処理分岐
                    '    If FolderType = 4 Then

                    '        '受注区分   (混在フォルダの場合は必須)
                    '        'strQtyValue = If(nJutyuKubun > 0, xlRow.Cell(nJutyuKubun).GetValue(Of String)().Trim(), "")
                    '        strQtyValue = m_ImpData(IntDataCnt).naijikubun
                    '        If String.IsNullOrEmpty(strQtyValue) Then
                    '            '必須チェック
                    '            errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：受注区分が空です。")
                    '            ErrFlg = True
                    '        ElseIf Not Decimal.TryParse(strQtyValue, ordertype) Then
                    '            '数値チェック
                    '            errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：受注区分が不正な値です。")
                    '            ErrFlg = True

                    '        End If

                    '        '受注区分 及び 分割区分 
                    '        If ordertype = 4 Then
                    '            '内示
                    '            '4(内示日別)だった場合、天方システム側の1(内示)、1(日割)にする)
                    '            ordertype = 1       '1(内示)
                    '            proratedtype = 1    '1(日割)

                    '        ElseIf ordertype = 1 Then
                    '            '確定
                    '            '1(確定)だった場合、天方システム側の2(確定)、2(日割以外)にする)
                    '            ordertype = 2       '2(確定)
                    '            proratedtype = 2    '2(日割以外)

                    '        End If

                    '    Else

                    '        '受注区分   (任意)
                    '        ordertype = FolderType

                    '        '分割区分   (任意)
                    '        'IMP_RULE_MSTより取得
                    '        proratedtype = 1
                    '        errMsg = ""
                    '        If _oderStageRepo.GetProratedType(CustomerSettingId, FolderType, proratedtype, errMsg) = False Then
                    '            'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                    '            'ErrFlg = True
                    '        End If

                    '    End If

                    '    '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                    '    customerorderNo = m_ImpData(IntDataCnt).ordersikibetuNo
                    '    If ordertype = 2 OrElse ordertype = 3 Then
                    '        'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                    '        If String.IsNullOrEmpty(customerorderNo) Then
                    '            errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：客先発注番号が空です。")
                    '            ErrFlg = True
                    '        End If
                    '    End If

                    '    '受注日
                    '    'orderDate = DateTime.Now
                    '    orderDate = DateTime.Today  'Todayにすることで当日の日付のみ、時刻は省かれる

                    '    '希望納期
                    '    strTempDate = m_ImpData(IntDataCnt).nonyusijibi
                    '    Dim parsedDate As Date
                    '    If DateTime.TryParseExact(strTempDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, parsedDate) Then
                    '        ' 変換成功
                    '        DueDate = parsedDate
                    '    Else
                    '        ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                    '        DueDate = CDate("1900/01/01")
                    '    End If

                    '    '日割前納期 ※希望納期をセット （希望納期が必須）
                    '    predailydeliveryDate = DueDate

                    '    '2026/06/29 日割前納期に希望納期をセットした後に希望納期の稼働日チェック
                    '    Dim cal = New CalenderRepository(Utils.GetConnectionString())
                    '    Dim tdt = New Date
                    '    tdt = DueDate
                    '    DueDate = cal.AddWorkingDays2("00001", tdt, 0)
                    '    '--

                    '    '需要数   (必須)
                    '    strQtyValue = m_ImpData(IntDataCnt).nonyusijisu
                    '    If String.IsNullOrEmpty(strQtyValue) Then
                    '        '必須チェック
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：需要数が空です。")
                    '        ErrFlg = True
                    '    ElseIf Not Decimal.TryParse(strQtyValue, demandqty) Then
                    '        '数値チェック
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：需要数が不正な値です")
                    '        ErrFlg = True
                    '    End If

                    '    '日割前受注数 ※需要数をセット　（需要数が必須）
                    '    predailyorderqty = demandqty

                    '    '自社予測フラグ
                    '    selffcstflag = "N"

                    '    '自社予測削除フラグ
                    '    selffcstdeleteflag = "N"

                    '    '需要ステイタス    （固定値）
                    '    demandstatus = If(ordertype = 1, "F", "O")

                    '    '累計出荷数      （固定値）
                    '    If ordertype = 1 Then
                    '        totalshipqty = Nothing
                    '    Else
                    '        totalshipqty = 0
                    '    End If

                    '    '出荷プロセスタイプ  （固定値）
                    '    Select Case ordertype
                    '        Case 1
                    '            shipprocesstype = "O"
                    '        Case 2
                    '            shipprocesstype = "E"
                    '        Case 3
                    '            shipprocesstype = "K"
                    '    End Select

                    '    '納入指示フラグ    （固定値）
                    '    deliveryinstrflag = If(ordertype = 3, "Y", "N")

                    '    '通貨コード  （任意）
                    '    'STRAMMIC.SECTMより取得
                    '    currencycode = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetCurrencyCode(CustomerCode, currencycode, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If


                    '    '客先品目No　品目No検索 (客先品目Noにハイフォンをつけて検索)
                    '    'STRAMMIC.PRDSLSODRMより取得
                    '    'Dim Patterns As New List(Of String) From {"3-5-2", "5-5"}
                    '    customeritemNo = m_ImpData(IntDataCnt).customeritemNo
                    '    itemNo = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetProductCode2(CustomerCode, customeritemNo, Patterns, m_ImpData(IntDataCnt).status, itemNo, errMsg) = False Then
                    '        errors.Add($"取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '        ErrFlg = True
                    '    End If

                    '    '製品コード
                    '    productcode = itemNo

                    '    '需要単位   （任意）
                    '    'STRAMMIC.ITEMMより取得
                    '    demandunit = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetDemandUnit(productcode, demandunit, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If

                    '    'コメント   （任意）
                    '    remarks = ""

                    '    '納入先コード
                    '    If m_ImpData(IntDataCnt).nonyuplat <> "" Then
                    '        deliverycode = m_ImpData(IntDataCnt).nonyuplat
                    '    Else
                    '        deliverycode = m_ImpData(IntDataCnt).siyosha
                    '    End If

                    '    '出荷在庫場所 （任意）
                    '    'STRAMMIC.ITEMMより取得
                    '    shipstocklocation = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If



                    '    '取引先情報区分
                    '    customerinfotype = ""

                    '    '情報区分
                    '    'INFO_TYPE_MSTより取得
                    '    infotype = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetInfoType(CustomerSettingId, customerinfotype, infotype, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If

                    '    '消込条件区分 ※順次/同月まで/同月内のみ
                    '    'IMP_RULE_MSTより取得
                    '    reconciletype = 1
                    '    errMsg = ""
                    '    If _oderStageRepo.GetReconcileType(CustomerSettingId, FolderType, reconciletype, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If


                    '    '出荷先　   （必須）
                    '    'STRAMMIC.SECTMより取得
                    '    shipto = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetShipTo(CustomerCode, deliverycode, shipto, errMsg) = False Then
                    '        errors.Add($"取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：{errMsg}")
                    '        ErrFlg = True
                    '    End If

                    '    '請求先    (任意）
                    '    'STRAMMIC.SECTMより取得
                    '    billingto = ""
                    '    errMsg = ""
                    '    If _oderStageRepo.GetBillingTo(CustomerCode, billingto, errMsg) = False Then
                    '        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '        'ErrFlg = True
                    '    End If

                    '    '生産区分　初品区分
                    '    If m_ImpData(IntDataCnt).status = "2" Then
                    '        seisankubun = "1"
                    '        syohinkubun = "1"
                    '    ElseIf m_ImpData(IntDataCnt).status = "3" Then
                    '        seisankubun = "2"
                    '        syohinkubun = "0"
                    '    Else
                    '        seisankubun = Nothing
                    '        syohinkubun = "1"
                    '    End If

                    '    ''品目NoのPCを取得
                    '    ''STRAMMIC.USRDEFFLDF(FTABLEID='ITEMM')より取得
                    '    'Dim pc As String = ""
                    '    ''If _oderStageRepo.GetItemNoPc(itemNo, pc, errMsg) = False Then
                    '    'If _oderStageRepo.GetProfitCenter(itemNo, pc, errMsg) = False Then
                    '    '    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                    '    '    ErrFlg = True
                    '    'End If

                    '    'Select Case pc
                    '    '    Case "E1", "E5", "E6", "E7", "E8", "E9"
                    '    '        siyosaki = "F999"
                    '    '    Case Else
                    '    '        siyosaki = ""
                    '    'End Select








                    '    '-----------------
                    '    '桁チェック
                    '    '-----------------
                    '    '受注区分
                    '    ordertype = SafeVarcharLength(ordertype, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：受注区分が桁数超過のためトリミングされました。")
                    '    End If
                    '    '分割区分
                    '    proratedtype = SafeVarcharLength(proratedtype, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：分割区分が桁数超過のためトリミングされました。")
                    '    End If
                    '    '客先発注番号
                    '    customerorderNo = SafeVarcharLength(customerorderNo, 40, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：客先発注番号が桁数超過のためトリミングされました。")
                    '    End If
                    '    '需要数
                    '    demandqty = Convert.ToDecimal(SafeVarcharLength(demandqty.ToString(), 10, isTruncated))
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).ordergyoNo}：需要数が桁数超過のためトリミングされました。")
                    '    End If
                    '    ''日割前受注数
                    '    'predailyorderqty = Convert.ToDecimal(SafeVarcharLength(predailyorderqty.ToString(), 10, isTruncated))
                    '    'If isTruncated = True Then
                    '    '    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：日割前受注数が桁数超過のためトリミングされました。")
                    '    'End If
                    '    '自社予測フラグ
                    '    selffcstflag = SafeVarcharLength(selffcstflag, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：自社予測フラグが桁数超過のためトリミングされました。")
                    '    End If
                    '    '自社予測削除フラグ
                    '    selffcstdeleteflag = SafeVarcharLength(selffcstdeleteflag, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：自社予測削除フラグが桁数超過のためトリミングされました。")
                    '    End If
                    '    '通貨コード
                    '    currencycode = SafeVarcharLength(currencycode, 3, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：通貨コードが桁数超過のためトリミングされました。")
                    '    End If
                    '    '客先品目No
                    '    customeritemNo = SafeVarcharLength(customeritemNo, 45, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：客先品目Noが桁数超過のためトリミングされました。")
                    '    End If
                    '    '製品コード
                    '    productcode = SafeVarcharLength(productcode, 45, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：製品コードが桁数超過のためトリミングされました。")
                    '    End If
                    '    '品目No
                    '    itemNo = SafeVarcharLength(itemNo, 45, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：品目Noが桁数超過のためトリミングされました。")
                    '    End If
                    '    '需要単位
                    '    demandunit = SafeVarcharLength(demandunit, 4, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：需要単位が桁数超過のためトリミングされました。")
                    '    End If
                    '    'コメント
                    '    remarks = SafeVarcharLength(remarks, 45, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：コメントが桁数超過のためトリミングされました。")
                    '    End If
                    '    '納入先コード
                    '    deliverycode = SafeVarcharLength(deliverycode, 25, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：納入先コードが桁数超過のためトリミングされました。")
                    '    End If
                    '    '出荷在庫場所
                    '    shipstocklocation = SafeVarcharLength(shipstocklocation, 25, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：出荷在庫場所が桁数超過のためトリミングされました。")
                    '    End If
                    '    '取引先情報区分
                    '    customerinfotype = SafeVarcharLength(customerinfotype, 50, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：取引先情報区分が桁数超過のためトリミングされました。")
                    '    End If
                    '    '情報区分
                    '    infotype = SafeVarcharLength(infotype, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：情報区分が桁数超過のためトリミングされました。")
                    '    End If
                    '    '消込条件区分
                    '    reconciletype = SafeVarcharLength(reconciletype, 1, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：消込条件区分が桁数超過のためトリミングされました。")
                    '    End If
                    '    '出荷先
                    '    shipto = SafeVarcharLength(shipto, 25, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：出荷先が桁数超過のためトリミングされました。")
                    '    End If
                    '    '請求先
                    '    billingto = SafeVarcharLength(billingto, 25, isTruncated)
                    '    If isTruncated = True Then
                    '        errors.Add($" 取引先コード：{CustomerCode}　取込ファイル：[{TorikomiFile} ]　Row {m_ImpData(IntDataCnt).hinmokugyoNo}：請求先が桁数超過のためトリミングされました。")
                    '    End If
                    '    '-----------------


                    '    If ErrFlg = True Then

                    '        'ErrCustomerCode = customerCode
                    '        'ErrTorikomiFile = TorikomiFile

                    '        fileidx += 1
                    '        errcnt += 1
                    '        ErrFileFlg = True
                    '        Continue For
                    '    End If





                    '    '受注ワーク登録用リストへ格納
                    '    rowsForTemp2.Add(New OrdersStageRow With {
                    '        .CustomerSettingId = CustomerSettingId,
                    '        .CustomerCode = CustomerCode,
                    '        .BillingTo = billingto,
                    '        .CustomerOrderNo = customerorderNo,
                    '        .DemandStatus = demandstatus,
                    '        .ShipTo = shipto,
                    '        .OrderDate = orderDate,
                    '        .DueDate = FormatDate(DueDate),
                    '        .CustomerItemNo = customeritemNo,
                    '        .ItemNo = itemNo,
                    '        .DemandQty = demandqty,
                    '        .DemandUnit = demandunit,
                    '        .CurrencyCode = currencycode,
                    '        .ShipStockLocation = shipstocklocation,
                    '        .CompanyId = "1000",
                    '        .ProductCode = productcode,
                    '        .BillingStandard = "S",
                    '        .ShipProcessType = shipprocesstype,
                    '        .DeliveryInstrFlag = deliveryinstrflag,
                    '        .Remarks = remarks,
                    '        .DeliveryCode = deliverycode,
                    '        .TotalShipQty = totalshipqty,
                    '        .TransportMethod = "2",
                    '        .PreDailyOrderQty = predailyorderqty,
                    '        .PreDailyDeliveryDate = predailydeliveryDate,
                    '        .ImpFileStageId = ImpFileStageId,
                    '        .OrderType = ordertype,
                    '        .ProratedType = proratedtype,
                    '        .CustomerInfoType = customerinfotype,
                    '        .InfoType = infotype,
                    '        .SelfFcstFlag = selffcstflag,
                    '        .SelfFcstDeleteFlag = selffcstdeleteflag,
                    '        .ReconcileType = reconciletype,
                    '        .ImpRunId = newId,
                    '        .Status = "IMPORTED",
                    '        .ActiveFlag = "Y",
                    '        .CreatedAt = Now,
                    '        .CreatedUserId = UserId,
                    '        .CreatedPgId = pgId,
                    '        .UpdatedAt = Now,
                    '        .UpdatedUserId = UserId,
                    '        .UpdatedPgId = pgId
                    '    })

                    '    fileidx += 1

                    'Next

                Case "EXCEL"
                    '4:EXCEL LIST





            End Select



            ' すべての検証を通過
            result.IsValid = True
            Return result
        End Function

        ''' <summary>
        ''' 配列に格納された取込データをワークテーブルに一括保存します。
        ''' </summary>
        Private Shared Sub SaveImportDataToWorkTable(
            ByVal tran As OracleTransaction,
            ByVal impDataList() As ImportDataType,
            ByVal impFileStageId As Long)

            ' 配列が空、またはデータが1件もない場合は処理を抜ける
            If impDataList Is Nothing OrElse impDataList.Length = 0 Then Return
            ' 配列が初期化（0）のままで、かつ有効なデータが1件もセットされていなければスキップ
            If impDataList.Length = 1 AndAlso String.IsNullOrEmpty(impDataList(0).customeritemNo) Then Return

            '前回の取込データをクリアする
            Using delCmd As New OracleCommand("DELETE FROM yamaha_imp_orders_test", tran.Connection)
                'delCmd.Transaction = tran
                delCmd.ExecuteNonQuery()
            End Using

            Dim nowTime As DateTime = DateTime.Now

            Dim sql As String = "
                        INSERT INTO yamaha_imp_orders_test (
                            imp_file_stage_id, hinmoku_gyo_no, order_gyo_no, siyosha, status, customer_item_no, nonyuplat,
                            yokisyuuyousuu, yokibangou, ordersikibetu_no, nonyusijibi, nonyusibijikan,
                            nonyusijisu, cardkubun, naijikubun, icdenpyo_no, nohinsho_no, created_at
                        ) VALUES (
                            :p_file_id, :p_hinmoku_gyo_no, :p_order_gyo_no, :p_siyosha, :p_status, :p_item_no, :p_plat,
                            :p_suu, :p_ban, :p_shikibetu, :p_jibi, :p_jikan,
                            :p_sijisu, :p_card, :p_naiji, :p_ic, :p_nohin, :p_created
                        )
                    "

            Using cmd As New OracleCommand(sql, tran.Connection)
                'cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()

                ' 処理高速化のため、ループ外で一度だけ型を定義
                cmd.Parameters.Add("p_file_id", OracleDbType.Int64)
                cmd.Parameters.Add("p_hinmoku_gyo_no", OracleDbType.Int32)
                cmd.Parameters.Add("p_order_gyo_no", OracleDbType.Int32)
                cmd.Parameters.Add("p_siyosha", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_status", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_item_no", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_plat", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_suu", OracleDbType.Int32)
                cmd.Parameters.Add("p_ban", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_shikibetu", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_jibi", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_jikan", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_sijisu", OracleDbType.Int32)
                cmd.Parameters.Add("p_card", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_naiji", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_ic", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nohin", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_created", OracleDbType.Date).Value = nowTime

                ' 配列の全要素をループしてインサートを実行
                For i As Integer = 0 To UBound(impDataList)
                    cmd.Parameters("p_file_id").Value = impFileStageId
                    cmd.Parameters("p_hinmoku_gyo_no").Value = impDataList(i).hinmokugyoNo
                    cmd.Parameters("p_order_gyo_no").Value = impDataList(i).ordergyoNo
                    cmd.Parameters("p_siyosha").Value = impDataList(i).siyosha
                    cmd.Parameters("p_status").Value = impDataList(i).status
                    cmd.Parameters("p_item_no").Value = impDataList(i).customeritemNo
                    cmd.Parameters("p_plat").Value = impDataList(i).nonyuplat
                    cmd.Parameters("p_suu").Value = If(String.IsNullOrEmpty(impDataList(i).yokisyuuyousuu), DBNull.Value, Convert.ToInt32(impDataList(i).yokisyuuyousuu))
                    cmd.Parameters("p_ban").Value = impDataList(i).yokibangou
                    cmd.Parameters("p_shikibetu").Value = impDataList(i).ordersikibetuNo
                    cmd.Parameters("p_jibi").Value = impDataList(i).nonyusijibi
                    cmd.Parameters("p_jikan").Value = impDataList(i).nonyusibijikan
                    cmd.Parameters("p_sijisu").Value = If(String.IsNullOrEmpty(impDataList(i).nonyusijisu), DBNull.Value, Convert.ToInt32(impDataList(i).nonyusijisu))
                    cmd.Parameters("p_card").Value = impDataList(i).cardkubun
                    cmd.Parameters("p_naiji").Value = impDataList(i).naijikubun
                    cmd.Parameters("p_ic").Value = impDataList(i).icdenpyoNo
                    cmd.Parameters("p_nohin").Value = impDataList(i).nohinshoNo

                    cmd.ExecuteNonQuery()
                Next
            End Using
        End Sub

    End Class

End Namespace
