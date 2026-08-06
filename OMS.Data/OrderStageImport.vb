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
Imports CsvHelper
Imports CsvHelper.Configuration
Imports System.Globalization
Imports System.Runtime.InteropServices

Namespace OMS.Data
    Public Class OrderStageImport

        'Private _folderRepo As FolderRepository

        Public Property IsValid As Boolean = False
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

        'マッピングマスタ用
        Public Class MappingResult
            Public Property HeaderRowIndex As Integer = 0
            Public Property DataStartRowIndex As Integer = 0
            Public Property DefaultSheetName As String = ""
            Public Property FormatType As String = ""
            Public Property FileType As String = ""
            Public Property Delimiter As Char = ","c
            Public Property Enclosure As String = ""
            Public Property HeaderFlag As String = ""
            Public Property CharSet As String = ""

            ' 列インデックス(LIST用)
            Public Property nKyakusakiHattyuNo As Integer = -1
            Public Property nJutyuuBi As Integer = -1
            Public Property nKibouNouki As Integer = -1
            Public Property nKyakusakiHinmokuNo As Integer = -1
            Public Property nJuyouSuu As Integer = -1
            Public Property nTukaCode As Integer = -1
            Public Property nSeihinCode As Integer = -1
            Public Property nNonyusakiCode As Integer = -1
            Public Property nComment As Integer = -1
            Public Property nJutyuKubun As Integer = -1
            Public Property nBunkatuKubun As Integer = -1
            Public Property nTorihikisakiJohoKubun As Integer = -1
            Public Property nJishaYosokuFlag As Integer = -1
            Public Property nJishaYosokuDelFlag As Integer = -1

            ' セルアドレス(MATRIX用)
            Public Property mKyakusakiHattyuNo As String = ""
            Public Property mJutyuuBi As String = ""
            Public Property mKibouNouki As String = ""
            Public Property mKyakusakiHinmokuNo As String = ""
            Public Property mJuyouSuu As String = ""
            Public Property mTukaCode As String = ""
            Public Property mSeihinCode As String = ""
            Public Property mNonyusakiCode As String = ""
            Public Property mComment As String = ""
            Public Property mJutyuKubun As String = ""
            Public Property mBunkatuKubun As String = ""
            Public Property mTorihikisakiJohoKubun As String = ""
            Public Property mJishaYosokuFlag As String = ""
            Public Property mJishaYosokuDelFlag As String = ""

            'Public Property errors As New List(Of String)()
        End Class

        '/// ヤマハ取込データ保存配列
        Private Structure ImportDataType
            Dim hakkobi As String           '発行日
            Dim hakkojikan As String        '発光時間
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
            Dim nonyujikan As String    '納入時間
            Dim nonyusijisu As String       '納入指示数
            Dim cardkubun As String         'カード区分
            Dim naijikubun As String        '内示区分
            Dim icdenpyoNo As String        'IC伝票No
            Dim nohinshoNo As String        '納品書番号
            Dim hinmokugyoNo As String      '品目情報行番号
            Dim ordergyoNo As String        'オーダー情報行番号
            Dim customercode As String      '得意先コード
        End Structure
        Private Shared m_ImpData() As ImportDataType


        Public Shared Function ResolveMapping(ByVal mappingRepo As MappingRepository,
                                        ByVal customerSettingId As Long,
                                        ByVal folderType As Integer,
                                        ByVal errors As List(Of String)) As MappingResult

            Dim mappingInfos As List(Of MappingInfo) = mappingRepo.GetMappingInfo(customerSettingId, folderType)

            Dim result As New MappingResult()

            If mappingInfos Is Nothing OrElse mappingInfos.Count = 0 Then
                errors.Add($"{customerSettingId}:MAPPINNG_PROFILE_MSTに未登録")
                Return Nothing
            End If



            Dim SourceColumnIndex As Integer = 0
            Dim SourceCellAddress As String = ""
            Dim FstFlg As Boolean = True

            For Each info In mappingInfos
                result.HeaderRowIndex = info.HeaderRowIndex
                result.DataStartRowIndex = info.DataStartRowIndex
                result.DefaultSheetName = info.default_sheet_name  'デフォルトシート名
                result.FormatType = info.format_type    'LIST/MATRIX
                result.FileType = info.file_type       'CSV/TSV/FIXED/EXCEL
                result.CharSet = info.charset          'UTF8/SJIS
                result.HeaderFlag = info.header_flag   'N/Y

                '区切り文字
                Select Case info.delimiter
                    Case "COMMA"
                        result.Delimiter = ","c
                    Case "TAB"
                        result.Delimiter = vbTab
                    Case "SEMICOLON"
                        result.Delimiter = ";"c
                    Case "PIPE"
                        result.Delimiter = "|"c
                    Case "SPACE"
                        result.Delimiter = " "c
                    Case "COLON"
                        result.Delimiter = ":"c
                    Case Else
                        result.Delimiter = ","c
                End Select

                '囲い文字
                Select Case info.enclosure
                    Case "D_QUOTE"
                        result.Enclosure = """"c
                    Case "S_QUOTE"
                        result.Enclosure = "'"c
                    Case Else
                        result.Enclosure = ""
                End Select

                'Dim sourceColumnIndex As Integer = info.source_column_index
                'If result.FileType <> "EXCEL" Then
                '    sourceColumnIndex -= 1
                'End If

                If result.FileType = "EXCEL" Then
                    'EXCELは開始列が1からのためそのまま
                    SourceColumnIndex = info.source_column_index    '列番号
                Else
                    'CSV、TSVは開始列が0からのため-1
                    SourceColumnIndex = info.source_column_index - 1   '列番号

                    If FstFlg = True Then

                        result.nKyakusakiHattyuNo = -1
                        result.nJutyuuBi = -1
                        result.nKibouNouki = -1
                        result.nKyakusakiHinmokuNo = -1
                        result.nJuyouSuu = -1
                        result.nTukaCode = -1
                        result.nSeihinCode = -1
                        result.nNonyusakiCode = -1
                        result.nComment = -1
                        result.nJutyuKubun = -1
                        result.nBunkatuKubun = -1
                        result.nTorihikisakiJohoKubun = -1
                        result.nJishaYosokuFlag = -1
                        result.nJishaYosokuDelFlag = -1

                        FstFlg = False

                    End If
                End If

                'MATRIX用
                'SourceCellAddress = info.source_cell_address    'EXCELセルアドレス

                '取得する項目の位置
                Select Case info.target_field
                    Case "CUSTOMER_ORDER_NO"
                        result.nKyakusakiHattyuNo = SourceColumnIndex
                        result.mKyakusakiHattyuNo = info.source_cell_address
                    Case "ORDER_DATE"
                        result.nJutyuuBi = SourceColumnIndex
                        result.mJutyuuBi = info.source_cell_address
                    Case "DUE_DATE"
                        result.nKibouNouki = SourceColumnIndex
                        result.mKibouNouki = info.source_cell_address
                    Case "CUSTOMER_ITEM_NO"
                        result.nKyakusakiHinmokuNo = SourceColumnIndex
                        result.mKyakusakiHinmokuNo = info.source_cell_address
                    Case "DEMAND_QTY"
                        result.nJuyouSuu = SourceColumnIndex
                        result.mJuyouSuu = info.source_cell_address
                    Case "CURRENCY_CODE"
                        result.nTukaCode = SourceColumnIndex
                        result.mTukaCode = info.source_cell_address
                    Case "PRODUCT_CODE"
                        result.nSeihinCode = SourceColumnIndex
                        result.mSeihinCode = info.source_cell_address
                    Case "REMARKS"
                        result.nComment = SourceColumnIndex
                        result.mComment = info.source_cell_address
                    Case "DELIVERY_CODE"
                        result.nNonyusakiCode = SourceColumnIndex
                        result.mNonyusakiCode = info.source_cell_address
                    Case "ORDER_TYPE"
                        result.nJutyuKubun = SourceColumnIndex
                        result.mJutyuKubun = info.source_cell_address
                    Case "PRORATED_TYPE"
                        result.nBunkatuKubun = SourceColumnIndex
                        result.mBunkatuKubun = info.source_cell_address
                    Case "CUSTOMER_INFO_TYPE"
                        result.nTorihikisakiJohoKubun = SourceColumnIndex
                        result.mTorihikisakiJohoKubun = info.source_cell_address
                    Case "SELF_FCST_FLAG"
                        result.nJishaYosokuFlag = SourceColumnIndex
                        result.mJishaYosokuFlag = info.source_cell_address
                    Case "SELF_FCST_DELETE_FLAG"
                        result.nJishaYosokuDelFlag = SourceColumnIndex
                        result.mJishaYosokuDelFlag = info.source_cell_address
                End Select
            Next

            Return result

        End Function

        Public Shared Function ParseImportFileY(ByVal CustomerSettingId As Long,
                                               ByVal customerCode As String,
                                               ByRef ErrFlg As Boolean,
                                               ByRef ErrFileFlg As Boolean,
                                               ByRef errcnt As Integer,
                                               ByVal FolderType As Integer,
                                               ByVal UserId As String,
                                               ByVal pgId As String,
                                               ByVal errors As List(Of String),
                                               ByVal rowsForTemp2 As List(Of OrdersStageRow),
                                               ByVal mapResult As OrderStageImport.MappingResult) As Boolean



            '実行管理
            'IMP_RUNの新しいIDを取得
            Dim newId As Integer = 0
            'newId += 1
            'IMP_RUNに新規レコード追加
            Dim now As DateTime = DateTime.Now
            'Dim userId As String = (If(Context?.User?.Identity?.Name, "")).Trim()


            Dim rowsForTemp As New List(Of ImpRunRow) From {
                New ImpRunRow With {
                    .StartedAt = now,
                    .Status = "RUNNING",
                    .StartedUserId = UserId,
                    .StartedPgId = pgId,
                    .FileCount = 0,
                    .RowCount = 0,
                    .ErrorCount = 0
                }
            }



        End Function


        Public Shared Function ParseImportFile(ByVal tran As OracleTransaction,
                                               ByVal CustomerSettingId As Long,
                                               ByVal customerCode As String,
                                               ByVal impfilestageId As String,
                                               ByVal strWorkFile As String,
                                               ByVal TorikomiFile As String,
                                               ByRef ErrFlg As Boolean,
                                               ByRef ErrFileFlg As Boolean,
                                               ByRef errcnt As Integer,
                                               ByVal FolderType As Integer,
                                               ByVal newId As Integer,
                                               ByVal UserId As String,
                                               ByVal pgId As String,
                                               ByVal errors As List(Of String),
                                               ByVal rowsForTemp2 As List(Of OrdersStageRow),
                                               ByVal mapResult As OrderStageImport.MappingResult) As Boolean

            Dim _oderStageRepo As New OrderStageRepository(Utils.GetConnectionString())

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

            Dim customerorderNo As String
            Dim orderDate As Date
            Dim dueDate As Date
            Dim customeritemNo As String
            Dim demandqty As Decimal?
            Dim demandunit As String
            Dim currencycode As String
            Dim productcode As String
            Dim remarks As String
            Dim deliverycode As String
            Dim predailyorderqty As Decimal?
            Dim predailydeliveryDate As Date
            Dim ordertype As Integer
            Dim proratedtype As Integer
            Dim customerinfotype As String
            Dim selffcstflag As String
            Dim selffcstdeleteflag As String
            Dim shipto As String
            Dim billingto As String
            Dim itemNo As String
            Dim demandstatus As String
            Dim shipprocesstype As String
            Dim deliveryinstrflag As String
            Dim totalshipqty As Decimal?
            Dim shipstocklocation As String
            Dim infotype As String
            Dim reconciletype As Integer
            Dim profitcenter As String
            Dim profitcenterCSM As String
            Dim errMsg As String

            Dim isTruncated As Boolean = False

            ' 戻り値となる受注データリストの初期化
            Dim rows As New List(Of OrdersStageRow)()

            'ファイル内の行インデックス
            Dim fileidx As Integer = 0


            'ファイル読み込み
            'Select Case FileType
            Select Case mapResult.FileType
                Case "CSV", "TSV"
                    'CSV(カンマ区切り) TSV(タブ区切り) 

                    '許可する拡張子のリスト（小文字で定義）
                    Dim allowedExtensions As New List(Of String) From {".csv", ".txt"}

                    'ファイルパスから拡張子を取得
                    Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                    '拡張子チェック
                    If Not allowedExtensions.Contains(fileExtension) Then
                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                        'Continue For
                        Return False
                    End If

                    Dim config As New CsvConfiguration(CultureInfo.InvariantCulture)
                    config.Delimiter = mapResult.Delimiter            '区切り文字
                    config.HasHeaderRecord = If(mapResult.HeaderFlag = "Y", True, False)      ' 1行目がヘッダー（列名）の場合
                    config.TrimOptions = TrimOptions.Trim ' 前後の余計な空白を自動で消す

                    If mapResult.Enclosure <> "" Then
                        config.Quote = mapResult.Enclosure               '囲い文字
                    Else
                        config.Mode = CsvMode.NoEscape
                    End If

                    Using StmRdr As New IO.StreamReader(strWorkFile, MapEncoding(mapResult.CharSet))

                        Using csv As New CsvReader(StmRdr, config)

                            fileidx = 1

                            'ヘッダーの存在チェック
                            If mapResult.HeaderFlag = "Y" Then
                                '/// 1行目のヘッダー行を飛ばす
                                csv.Read()
                                csv.ReadHeader()    'ヘッダーとして登録
                                fileidx += 1
                            End If

                            'データ開始行の手前までポインタを空打ちして進める
                            While fileidx < mapResult.DataStartRowIndex - 1
                                csv.Read()
                                fileidx += 1
                            End While

                            While csv.Read()

                                '初期化
                                strTempDate = ""  '日付検証用
                                strQtyValue = ""  '数値検証用
                                customerorderNo = ""
                                orderDate = Nothing
                                dueDate = Nothing
                                customeritemNo = ""
                                demandqty = 0
                                demandunit = ""
                                currencycode = ""
                                productcode = ""
                                remarks = ""
                                deliverycode = ""
                                predailyorderqty = 0
                                predailydeliveryDate = Nothing
                                ordertype = 0
                                proratedtype = 1
                                customerinfotype = ""
                                selffcstflag = ""
                                selffcstdeleteflag = ""
                                shipto = ""
                                billingto = ""
                                itemNo = ""
                                demandstatus = ""
                                shipprocesstype = ""
                                deliveryinstrflag = ""
                                totalshipqty = 0
                                shipstocklocation = ""
                                infotype = ""
                                reconciletype = 1
                                profitcenterCSM = ""
                                profitcenter = ""

                                errMsg = ""

                                ErrFlg = False


                                '取引先設定IDのPC   （必須）
                                'CUSTOMER_SETTING_MSTより取得
                                profitcenterCSM = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '客先品目No   (任意)
                                If mapResult.nKyakusakiHinmokuNo > -1 Then
                                    customeritemNo = If(csv.ColumnCount > mapResult.nKyakusakiHinmokuNo AndAlso mapResult.nKyakusakiHinmokuNo > -1, csv.GetField(mapResult.nKyakusakiHinmokuNo).Trim(), "")
                                End If

                                '製品コード  （任意）
                                If mapResult.nSeihinCode > -1 Then
                                    productcode = If(csv.ColumnCount > mapResult.nSeihinCode AndAlso mapResult.nSeihinCode > -1, csv.GetField(mapResult.nSeihinCode).Trim(), "")
                                End If

                                '品目No   （必須）
                                'STRAMMIC.PRDSLSODRMより取得
                                itemNo = ""
                                errMsg = ""
                                If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '品目NoのPC   （必須）
                                'STRAMMIC.USRDEFFLDFより取得
                                profitcenter = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先設定IDのPCと同じPCのみ取込対象とする
                                If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                    'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                    fileidx += 1
                                    Continue While
                                End If


                                'フォルダタイプで処理分岐
                                If FolderType = 4 Then

                                    '受注区分   (混在フォルダの場合は必須)
                                    strQtyValue = If(csv.ColumnCount > mapResult.nJutyuKubun AndAlso mapResult.nJutyuKubun > -1, csv.GetField(mapResult.nJutyuKubun).Trim(), "")
                                    If String.IsNullOrEmpty(strQtyValue) Then
                                        '必須チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が空です。")
                                        ErrFlg = True
                                    ElseIf Not Decimal.TryParse(strQtyValue, ordertype) Then
                                        '数値チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が不正な値です。")
                                        ErrFlg = True
                                    End If

                                    '分割区分   (混在フォルダの場合は必須)
                                    strQtyValue = If(csv.ColumnCount > mapResult.nBunkatuKubun AndAlso mapResult.nBunkatuKubun > -1, csv.GetField(mapResult.nBunkatuKubun).Trim(), "")
                                    If String.IsNullOrEmpty(strQtyValue) Then
                                        '必須チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が空です。")
                                        ErrFlg = True
                                    ElseIf Not Decimal.TryParse(strQtyValue, proratedtype) Then
                                        '数値チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が不正な値です。")
                                        ErrFlg = True
                                    End If

                                Else

                                    '受注区分   (任意)
                                    ordertype = FolderType

                                    '分割区分   (任意)
                                    'IMP_RULE_MSTより取得
                                    proratedtype = 1
                                    errMsg = ""
                                    If _oderStageRepo.GetProratedType(CustomerSettingId, FolderType, proratedtype, errMsg) = False Then
                                        'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                        'ErrFlg = True
                                    End If

                                End If

                                '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                                customerorderNo = If(csv.ColumnCount > mapResult.nKyakusakiHattyuNo AndAlso mapResult.nKyakusakiHattyuNo > -1, csv.GetField(mapResult.nKyakusakiHattyuNo).Trim(), "")
                                If ordertype = 2 OrElse ordertype = 3 Then
                                    'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                    If String.IsNullOrEmpty(customerorderNo) Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '受注日   (任意)
                                strTempDate = If(csv.ColumnCount > mapResult.nJutyuuBi AndAlso mapResult.nJutyuuBi > -1, csv.GetField(mapResult.nJutyuuBi).Trim(), "")
                                If String.IsNullOrEmpty(strTempDate) Then
                                    orderDate = CDate("1900/01/01")
                                Else
                                    ' 日付変換を試みる（yyyy/MM/dd形式）
                                    If Not DateTime.TryParseExact(strTempDate, formats,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        orderDate) Then
                                        ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                        orderDate = CDate("1900/01/01")
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注日が不正な値です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '希望納期   (必須)
                                strTempDate = If(csv.ColumnCount > mapResult.nKibouNouki AndAlso mapResult.nKibouNouki > -1, csv.GetField(mapResult.nKibouNouki).Trim(), "")
                                If String.IsNullOrEmpty(strTempDate) Then
                                    '必須チェック
                                    dueDate = CDate("1900/01/01")
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：希望納期が空です。")
                                    ErrFlg = True
                                Else
                                    If Not DateTime.TryParseExact(strTempDate, formats,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        dueDate) Then
                                        ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                        dueDate = CDate("1900/01/01")
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：希望納期が不正な値です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '日割前納期 ※希望納期をセット （希望納期が必須）
                                predailydeliveryDate = dueDate

                                '2026/06/29 日割前納期に希望納期をセットした後に希望納期の稼働日チェック
                                Dim cal = New CalenderRepository(Utils.GetConnectionString())
                                Dim tdt = New Date?
                                tdt = dueDate
                                dueDate = cal.AddWorkingDays2("00001", tdt, 0)
                                '--

                                '需要数   (必須)
                                strQtyValue = If(csv.ColumnCount > mapResult.nJuyouSuu AndAlso mapResult.nJuyouSuu > -1, csv.GetField(mapResult.nJuyouSuu).Trim(), "")
                                If String.IsNullOrEmpty(strQtyValue) Then
                                    '必須チェック
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が空です。")
                                    ErrFlg = True
                                ElseIf Not Decimal.TryParse(strQtyValue, demandqty) Then
                                    '数値チェック
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が不正な値です。")
                                    ErrFlg = True
                                End If

                                '日割前受注数 ※需要数をセット　（需要数が必須）
                                predailyorderqty = demandqty

                                '自社予測フラグ   (任意)
                                selffcstflag = If(csv.ColumnCount > mapResult.nJishaYosokuFlag AndAlso mapResult.nJishaYosokuFlag > -1, csv.GetField(mapResult.nJishaYosokuFlag).Trim(), "")

                                '自社予測削除フラグ   (自社予測フラグ = Yの時は必須)
                                selffcstdeleteflag = If(csv.ColumnCount > mapResult.nJishaYosokuDelFlag AndAlso mapResult.nJishaYosokuDelFlag > -1, csv.GetField(mapResult.nJishaYosokuDelFlag).Trim(), "")
                                If selffcstflag = "Y" AndAlso String.IsNullOrEmpty(selffcstdeleteflag) Then
                                    'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測フラグが'Y'ですが、自社予測削除フラグが取得できないか空です。")
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測削除フラグが空です。")
                                    ErrFlg = True
                                End If

                                '需要ステイタス    （固定値）
                                demandstatus = If(ordertype = 1, "F", "O")

                                '累計出荷数    （固定値）
                                If ordertype = 1 Then
                                    totalshipqty = Nothing
                                Else
                                    totalshipqty = 0
                                End If

                                '出荷プロセスタイプ    （固定値）
                                Select Case ordertype
                                    Case 1
                                        shipprocesstype = "O"
                                    Case 2
                                        shipprocesstype = "E"
                                    Case 3
                                        shipprocesstype = "K"
                                End Select

                                '納入指示フラグ    （固定値）
                                deliveryinstrflag = If(ordertype = 3, "Y", "N")

                                '通貨コード  （任意）
                                If mapResult.nTukaCode > -1 Then
                                    '取得ファイルに存在
                                    currencycode = If(csv.ColumnCount > mapResult.nTukaCode AndAlso mapResult.nTukaCode > -1, csv.GetField(mapResult.nTukaCode).Trim(), "")
                                Else
                                    '取得できない場合はSTRAMMIC.SECTMより取得
                                    currencycode = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetCurrencyCode(customerCode, currencycode, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                        'ErrFlg = True
                                    End If
                                End If

                                '客先品目No   (任意)
                                If mapResult.nKyakusakiHinmokuNo > -1 Then
                                    customeritemNo = If(csv.ColumnCount > mapResult.nKyakusakiHinmokuNo AndAlso mapResult.nKyakusakiHinmokuNo > -1, csv.GetField(mapResult.nKyakusakiHinmokuNo).Trim(), "")
                                End If

                                '製品コード  （任意）
                                If mapResult.nSeihinCode > -1 Then
                                    productcode = If(csv.ColumnCount > mapResult.nSeihinCode AndAlso mapResult.nSeihinCode > -1, csv.GetField(mapResult.nSeihinCode).Trim(), "")
                                End If

                                '品目No   （必須）
                                'STRAMMIC.PRDSLSODRMより取得
                                itemNo = ""
                                errMsg = ""
                                If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    ErrFlg = True
                                End If

                                '需要単位   （任意）
                                'STRAMMIC.ITEMMより取得
                                demandunit = ""
                                errMsg = ""
                                If _oderStageRepo.GetDemandUnit(productcode, demandunit, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                'コメント   （任意）
                                remarks = If(csv.ColumnCount > mapResult.nComment AndAlso mapResult.nComment > -1, csv.GetField(mapResult.nComment).Trim(), "")

                                '納入先コード   （任意）
                                deliverycode = If(csv.ColumnCount > mapResult.nNonyusakiCode AndAlso mapResult.nNonyusakiCode > -1, csv.GetField(mapResult.nNonyusakiCode).Trim(), "")

                                '出荷在庫場所   （任意）
                                'STRAMMIC.ITEMMより取得
                                shipstocklocation = ""
                                errMsg = ""
                                'If _oderStageRepo.GetShipStockLocation(customerCode, deliverycode, shipstocklocation, errMsg) = False Then
                                If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先情報区分   （任意）
                                customerinfotype = If(csv.ColumnCount > mapResult.nTorihikisakiJohoKubun AndAlso mapResult.nTorihikisakiJohoKubun > -1, csv.GetField(mapResult.nTorihikisakiJohoKubun).Trim(), "")

                                '情報区分   （任意）
                                'INFO_TYPE_MSTより取得
                                infotype = ""
                                errMsg = ""
                                If _oderStageRepo.GetInfoType(CustomerSettingId, customerinfotype, infotype, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '消込条件区分 ※順次/同月まで/同月内のみ   （任意）
                                'IMP_RULE_MSTより取得
                                reconciletype = 1
                                errMsg = ""
                                If _oderStageRepo.GetReconcileType(CustomerSettingId, FolderType, reconciletype, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '出荷先　   （必須）　
                                'STRAMMIC.SECTMより取得
                                shipto = ""
                                errMsg = ""
                                If _oderStageRepo.GetShipTo(customerCode, deliverycode, shipto, errMsg) = False Then
                                    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    ErrFlg = True
                                End If

                                '請求先    (任意）
                                'STRAMMIC.SECTMより取得
                                billingto = ""
                                errMsg = ""
                                If _oderStageRepo.GetBillingTo(customerCode, billingto, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If



                                '取引先設定IDのPC   （必須）
                                'CUSTOMER_SETTING_MSTより取得
                                profitcenterCSM = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '品目NoのPC   （必須）
                                'STRAMMIC.USRDEFFLDFより取得
                                profitcenter = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先設定IDのPCと同じPCのみ取込対象とする
                                If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                    'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                    fileidx += 1
                                    Continue While
                                End If



                                '-----------------
                                '桁チェック
                                '-----------------
                                '受注区分
                                ordertype = SafeVarcharLength(ordertype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が桁数超過のためトリミングされました。")
                                End If
                                '分割区分
                                proratedtype = SafeVarcharLength(proratedtype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が桁数超過のためトリミングされました。")
                                End If
                                '客先発注番号
                                customerorderNo = SafeVarcharLength(customerorderNo, 40, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が桁数超過のためトリミングされました。")
                                End If
                                '需要数
                                demandqty = Convert.ToDecimal(SafeVarcharLength(demandqty.ToString(), 10, isTruncated))
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が桁数超過のためトリミングされました。")
                                End If
                                ''日割前受注数
                                'predailyorderqty = Convert.ToDecimal(SafeVarcharLength(predailyorderqty.ToString(), 10, isTruncated))
                                'If isTruncated = True Then
                                '    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：日割前受注数が桁数超過のためトリミングされました。")
                                'End If
                                '自社予測フラグ
                                selffcstflag = SafeVarcharLength(selffcstflag, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測フラグが桁数超過のためトリミングされました。")
                                End If
                                '自社予測削除フラグ
                                selffcstdeleteflag = SafeVarcharLength(selffcstdeleteflag, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測削除フラグが桁数超過のためトリミングされました。")
                                End If
                                '通貨コード
                                currencycode = SafeVarcharLength(currencycode, 3, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：通貨コードが桁数超過のためトリミングされました。")
                                End If
                                '客先品目No
                                customeritemNo = SafeVarcharLength(customeritemNo, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先品目Noが桁数超過のためトリミングされました。")
                                End If
                                '製品コード
                                productcode = SafeVarcharLength(productcode, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：製品コードが桁数超過のためトリミングされました。")
                                End If
                                '品目No
                                itemNo = SafeVarcharLength(itemNo, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：品目Noが桁数超過のためトリミングされました。")
                                End If
                                '需要単位
                                demandunit = SafeVarcharLength(demandunit, 4, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要単位が桁数超過のためトリミングされました。")
                                End If
                                'コメント
                                remarks = SafeVarcharLength(remarks, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：コメントが桁数超過のためトリミングされました。")
                                End If
                                '納入先コード
                                deliverycode = SafeVarcharLength(deliverycode, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：納入先コードが桁数超過のためトリミングされました。")
                                End If
                                '出荷在庫場所
                                shipstocklocation = SafeVarcharLength(shipstocklocation, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：出荷在庫場所が桁数超過のためトリミングされました。")
                                End If
                                '取引先情報区分
                                customerinfotype = SafeVarcharLength(customerinfotype, 50, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：取引先情報区分が桁数超過のためトリミングされました。")
                                End If
                                '情報区分
                                infotype = SafeVarcharLength(infotype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：情報区分が桁数超過のためトリミングされました。")
                                End If
                                '消込条件区分
                                reconciletype = SafeVarcharLength(reconciletype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：消込条件区分が桁数超過のためトリミングされました。")
                                End If
                                '出荷先
                                shipto = SafeVarcharLength(shipto, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：出荷先が桁数超過のためトリミングされました。")
                                End If
                                '請求先
                                billingto = SafeVarcharLength(billingto, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：請求先が桁数超過のためトリミングされました。")
                                End If
                                '-----------------

                                'ここまででエラーフラグがあれば登録しない
                                If ErrFlg = True Then

                                    'ErrCustomerCode = customerCode
                                    'ErrTorikomiFile = TorikomiFile

                                    fileidx += 1
                                    errcnt += 1
                                    ErrFileFlg = True
                                    Continue While
                                End If

                                '受注ワーク登録用リストへ格納
                                rowsForTemp2.Add(New OrdersStageRow With {
                                    .CustomerSettingId = CustomerSettingId,
                                    .CustomerCode = customerCode,
                                    .BillingTo = billingto,
                                    .CustomerOrderNo = customerorderNo,
                                    .DemandStatus = demandstatus,
                                    .ShipTo = shipto,
                                    .OrderDate = orderDate,
                                    .DueDate = FormatDate(dueDate),
                                    .CustomerItemNo = customeritemNo,
                                    .ItemNo = itemNo,
                                    .DemandQty = demandqty,
                                    .DemandUnit = demandunit,
                                    .CurrencyCode = currencycode,
                                    .ShipStockLocation = shipstocklocation,
                                    .CompanyId = "1000",
                                    .ProductCode = productcode,
                                    .BillingStandard = "S",
                                    .ShipProcessType = shipprocesstype,
                                    .DeliveryInstrFlag = deliveryinstrflag,
                                    .Remarks = remarks,
                                    .DeliveryCode = deliverycode,
                                    .TotalShipQty = totalshipqty,
                                    .TransportMethod = "2",
                                    .PreDailyOrderQty = predailyorderqty,
                                    .PreDailyDeliveryDate = predailydeliveryDate,
                                    .ImpFileStageId = impfilestageId,
                                    .OrderType = ordertype,
                                    .ProratedType = proratedtype,
                                    .CustomerInfoType = customerinfotype,
                                    .InfoType = infotype,
                                    .SelfFcstFlag = selffcstflag,
                                    .SelfFcstDeleteFlag = selffcstdeleteflag,
                                    .ReconcileType = reconciletype,
                                    .ImpRunId = newId,
                                    .Status = "IMPORTED",
                                    .ActiveFlag = "Y",
                                    .CreatedAt = Now,
                                    .CreatedUserId = UserId,
                                    .CreatedPgId = pgId,
                                    .UpdatedAt = Now,
                                    .UpdatedUserId = UserId,
                                    .UpdatedPgId = pgId
                                })

                                fileidx += 1

                            End While

                        End Using

                    End Using


                Case "FIXED"
                    '3:FIXED(固定長)
                    'フェーズ2で実装

                    '許可する拡張子のリスト（小文字で定義）
                    Dim allowedExtensions As New List(Of String) From {".txt"}

                    'ファイルパスから拡張子を取得
                    Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                    '拡張子チェック
                    If Not allowedExtensions.Contains(fileExtension) Then
                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                        Return False
                    End If

                    Using StmRdr As New IO.StreamReader(strWorkFile, MapEncoding(mapResult.CharSet))

                        'ヤマハ取込データ保存配列　初期化
                        'Dim m_ImpData(0) As ImportDataType
                        ReDim m_ImpData(0)

                        Dim isFirstRow As Boolean = True ' 初回の要素追加判定用

                        ' ヘッダ情報をループ内で保持するための退避用変数
                        Dim currentHakkobi As String = ""
                        Dim currentHakkojikan As String = ""
                        Dim currentCostomercode As String = ""

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

                                'ヘッダ情報を取込データから取得
                                currentHakkobi = StrLine.Substring(5 - 1, 8).Trim()           '発行日
                                currentHakkojikan = StrLine.Substring(13 - 1, 4).Trim()       '発行時刻
                                currentCostomercode = customerCode                            '得意先コード

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

                                    'ヘッダ情報をセット
                                    m_ImpData(UBound(m_ImpData)).hakkobi = currentHakkobi                                   '発行日
                                    m_ImpData(UBound(m_ImpData)).hakkojikan = currentHakkojikan                             '発行時間
                                    m_ImpData(UBound(m_ImpData)).customercode = currentCostomercode                         '得意先コード

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
                                    m_ImpData(UBound(m_ImpData)).nonyujikan = StrLine.Substring(18 - 1, 4).Trim()      '納入時間1
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

                                    'ヘッダ情報をセット
                                    m_ImpData(UBound(m_ImpData)).hakkobi = currentHakkobi                                   '発行日
                                    m_ImpData(UBound(m_ImpData)).hakkojikan = currentHakkojikan                             '発行時間
                                    m_ImpData(UBound(m_ImpData)).customercode = currentCostomercode                         '得意先コード

                                    '品目情報をセット
                                    m_ImpData(UBound(m_ImpData)).siyosha = currentSiyosha                                   '使用者
                                    m_ImpData(UBound(m_ImpData)).status = currentStatus                                     '品目ステータス
                                    m_ImpData(UBound(m_ImpData)).customeritemNo = currentCustomeritemNo                     '旧体系部品番号(客先品目No)
                                    m_ImpData(UBound(m_ImpData)).nonyuplat = currentNonyuplat                               '納入プラットフォーム
                                    m_ImpData(UBound(m_ImpData)).yokisyuuyousuu = currentYokisyuuyousuu                     '容器収容数
                                    m_ImpData(UBound(m_ImpData)).yokibangou = currentYokibangou                             '容器番号

                                    m_ImpData(UBound(m_ImpData)).ordersikibetuNo = StrLine.Substring(89 - 1, 5).Trim()     'オーダー識別番号2
                                    m_ImpData(UBound(m_ImpData)).nonyusijibi = StrLine.Substring(96 - 1, 8).Trim()         '納入指示日2
                                    m_ImpData(UBound(m_ImpData)).nonyujikan = StrLine.Substring(104 - 1, 4).Trim()     '納入時間2
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
                    SaveImportDataToWorkTable(tran, m_ImpData, impfilestageId, newId, UserId, pgId)


                Case "EXCEL"
                    '4:EXCEL LIST

                    '許可する拡張子のリスト（小文字で定義）
                    Dim allowedExtensions As New List(Of String) From {".xlsx"}

                    'ファイルパスから拡張子を取得
                    Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                    '拡張子チェック
                    If Not allowedExtensions.Contains(fileExtension) Then
                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                        'Continue For
                        Return False
                    End If

                    If mapResult.FormatType = "LIST" Then

                        'ワークブックを作成
                        Using objWorkBook As New ClosedXML.Excel.XLWorkbook(strWorkFile)

                            'ワークシートを作成
                            Dim objSheet As ClosedXML.Excel.IXLWorksheet

                            'ワークシート指定があれば指定
                            If mapResult.DefaultSheetName <> "" Then
                                objSheet = objWorkBook.Worksheet(mapResult.DefaultSheetName)
                            Else
                                objSheet = objWorkBook.Worksheet(1)
                            End If

                            'データの最終行を取得
                            Dim lastRow = objSheet.LastRowUsed().RowNumber()

                            fileidx = mapResult.DataStartRowIndex

                            'データ開始行からスタート
                            For rowNum As Integer = mapResult.DataStartRowIndex To lastRow

                                Dim xlRow = objSheet.Row(rowNum)

                                '初期化
                                strTempDate = ""  '日付検証用
                                strQtyValue = ""  '数値検証用
                                customerorderNo = ""
                                orderDate = Nothing
                                dueDate = Nothing
                                customeritemNo = ""
                                demandqty = 0
                                demandunit = ""
                                currencycode = ""
                                productcode = ""
                                remarks = ""
                                deliverycode = ""
                                predailyorderqty = 0
                                predailydeliveryDate = Nothing
                                ordertype = 0
                                proratedtype = 1
                                customerinfotype = ""
                                selffcstflag = ""
                                selffcstdeleteflag = ""
                                shipto = ""
                                billingto = ""
                                itemNo = ""
                                demandstatus = ""
                                shipprocesstype = ""
                                deliveryinstrflag = ""
                                totalshipqty = 0
                                shipstocklocation = ""
                                infotype = ""
                                reconciletype = 1
                                profitcenterCSM = ""
                                profitcenter = ""
                                errMsg = ""

                                ErrFlg = False



                                '取引先設定IDのPC   （必須）
                                'CUSTOMER_SETTING_MSTより取得
                                profitcenterCSM = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '客先品目No   (任意)
                                If mapResult.nKyakusakiHinmokuNo > 0 Then
                                    customeritemNo = If(mapResult.nKyakusakiHinmokuNo > 0, xlRow.Cell(mapResult.nKyakusakiHinmokuNo).GetValue(Of String)().Trim(), "")
                                End If

                                '製品コード  （任意）
                                If mapResult.nSeihinCode > 0 Then
                                    productcode = If(mapResult.nSeihinCode > 0, xlRow.Cell(mapResult.nSeihinCode).GetValue(Of String)().Trim(), "")
                                End If

                                '品目No   （必須）
                                'STRAMMIC.PRDSLSODRMより取得
                                itemNo = ""
                                errMsg = ""
                                If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '品目NoのPC   （必須）
                                'STRAMMIC.USRDEFFLDFより取得
                                profitcenter = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先設定IDのPCと同じPCのみ取込対象とする
                                If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                    'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                    fileidx += 1
                                    Continue For
                                End If



                                'フォルダタイプで処理分岐
                                If FolderType = 4 Then

                                    '受注区分   (混在フォルダの場合は必須)
                                    strQtyValue = If(mapResult.nJutyuKubun > 0, xlRow.Cell(mapResult.nJutyuKubun).GetValue(Of String)().Trim(), "")
                                    If String.IsNullOrEmpty(strQtyValue) Then
                                        '必須チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が空です。")
                                        ErrFlg = True
                                    ElseIf Not Decimal.TryParse(strQtyValue, ordertype) Then
                                        '数値チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が不正な値です。")
                                        ErrFlg = True

                                    End If

                                    '分割区分   (混在フォルダの場合は必須)
                                    strQtyValue = If(mapResult.nBunkatuKubun > 0, xlRow.Cell(mapResult.nBunkatuKubun).GetValue(Of String)().Trim(), "")
                                    If String.IsNullOrEmpty(strQtyValue) Then
                                        '必須チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が空です。")
                                        ErrFlg = True
                                    ElseIf Not Decimal.TryParse(strQtyValue, proratedtype) Then
                                        '数値チェック
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が不正な値です。")
                                        ErrFlg = True

                                    End If

                                Else

                                    '受注区分   (任意)
                                    ordertype = FolderType

                                    '分割区分   (任意)
                                    'IMP_RULE_MSTより取得
                                    proratedtype = 1
                                    errMsg = ""
                                    If _oderStageRepo.GetProratedType(CustomerSettingId, FolderType, proratedtype, errMsg) = False Then
                                        'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                        'ErrFlg = True
                                    End If

                                End If

                                '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                                customerorderNo = If((mapResult.nKyakusakiHattyuNo) > 0, xlRow.Cell(mapResult.nKyakusakiHattyuNo).GetValue(Of String)().Trim(), "")
                                If ordertype = 2 OrElse ordertype = 3 Then
                                    'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                    If String.IsNullOrEmpty(customerorderNo) Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '受注日   (任意)
                                strTempDate = If(mapResult.nJutyuuBi > 0, xlRow.Cell(mapResult.nJutyuuBi).GetValue(Of String)().Trim(), "")
                                If String.IsNullOrEmpty(strTempDate) Then
                                    orderDate = CDate("1900/01/01")
                                Else
                                    ' 日付変換を試みる（yyyy/MM/dd形式）
                                    If Not DateTime.TryParseExact(strTempDate, formats,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None,
                                            orderDate) Then
                                        ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                        orderDate = CDate("1900/01/01")
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注日が不正な値です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '希望納期   (必須)
                                strTempDate = If(mapResult.nKibouNouki > 0, xlRow.Cell(mapResult.nKibouNouki).GetValue(Of String)().Trim(), "")
                                If String.IsNullOrEmpty(strTempDate) Then
                                    '必須チェック
                                    dueDate = CDate("1900/01/01")
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：希望納期が空です。")
                                    ErrFlg = True
                                Else
                                    If Not DateTime.TryParseExact(strTempDate, formats,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        dueDate) Then
                                        ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                        dueDate = CDate("1900/01/01")
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：希望納期が不正な値です。")
                                        ErrFlg = True
                                    End If
                                End If

                                '日割前納期 ※希望納期をセット （希望納期が必須）
                                predailydeliveryDate = dueDate

                                '2026/06/29 日割前納期に希望納期をセットした後に希望納期の稼働日チェック
                                Dim cal = New CalenderRepository(Utils.GetConnectionString())
                                Dim tdt = New Date
                                tdt = dueDate
                                dueDate = cal.AddWorkingDays2("00001", tdt, 0)
                                '--


                                '需要数   (必須)
                                strQtyValue = If(mapResult.nJuyouSuu > 0, xlRow.Cell(mapResult.nJuyouSuu).GetValue(Of String)().Trim(), "")
                                If String.IsNullOrEmpty(strQtyValue) Then
                                    '必須チェック
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が空です。")
                                    ErrFlg = True
                                ElseIf Not Decimal.TryParse(strQtyValue, demandqty) Then
                                    '数値チェック
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が不正な値です")
                                    ErrFlg = True
                                End If

                                '日割前受注数 ※需要数をセット　（需要数が必須）
                                predailyorderqty = demandqty

                                '自社予測フラグ   (任意)
                                selffcstflag = If(mapResult.nJishaYosokuFlag > 0, xlRow.Cell(mapResult.nJishaYosokuFlag).GetValue(Of String)().Trim(), "")

                                '自社予測削除フラグ  ※'自社予測フラグ = Yの時は必須
                                selffcstdeleteflag = If(mapResult.nJishaYosokuDelFlag > 0, xlRow.Cell(mapResult.nJishaYosokuDelFlag).GetValue(Of String)().Trim(), "")
                                If selffcstflag = "Y" AndAlso String.IsNullOrEmpty(selffcstdeleteflag) Then
                                    'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測フラグが'Y'ですが、自社予測削除フラグが取得できないか空です。")
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測削除フラグが空です。")
                                    ErrFlg = True
                                End If

                                '需要ステイタス    （固定値）
                                demandstatus = If(ordertype = 1, "F", "O")

                                '累計出荷数      （固定値）
                                If ordertype = 1 Then
                                    totalshipqty = Nothing
                                Else
                                    totalshipqty = 0
                                End If

                                '出荷プロセスタイプ  （固定値）
                                Select Case ordertype
                                    Case 1
                                        shipprocesstype = "O"
                                    Case 2
                                        shipprocesstype = "E"
                                    Case 3
                                        shipprocesstype = "K"
                                End Select

                                '納入指示フラグ    （固定値）
                                deliveryinstrflag = If(ordertype = 3, "Y", "N")

                                '通貨コード  （任意）
                                If mapResult.nTukaCode > 0 Then
                                    '取得ファイルに存在
                                    currencycode = If(mapResult.nTukaCode > 0, xlRow.Cell(mapResult.nTukaCode).GetValue(Of String)().Trim(), "")
                                Else
                                    '取得できない場合はSTRAMMIC.SECTMより取得
                                    currencycode = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetCurrencyCode(customerCode, currencycode, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                        'ErrFlg = True
                                    End If
                                End If

                                '客先品目No   (任意)
                                If mapResult.nKyakusakiHinmokuNo > 0 Then
                                    customeritemNo = If(mapResult.nKyakusakiHinmokuNo > 0, xlRow.Cell(mapResult.nKyakusakiHinmokuNo).GetValue(Of String)().Trim(), "")
                                End If

                                '製品コード  （任意）
                                If mapResult.nSeihinCode > 0 Then
                                    productcode = If(mapResult.nSeihinCode > 0, xlRow.Cell(mapResult.nSeihinCode).GetValue(Of String)().Trim(), "")
                                End If

                                '品目No   （必須）
                                'STRAMMIC.PRDSLSODRMより取得
                                itemNo = ""
                                errMsg = ""
                                If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    ErrFlg = True
                                End If



                                '需要単位   （任意）
                                'STRAMMIC.ITEMMより取得
                                demandunit = ""
                                errMsg = ""
                                If _oderStageRepo.GetDemandUnit(productcode, demandunit, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                'コメント   （任意）
                                remarks = If(mapResult.nComment > 0, xlRow.Cell(mapResult.nComment).GetValue(Of String)().Trim(), "")

                                '納入先コード （任意）
                                deliverycode = If(mapResult.nNonyusakiCode > 0, xlRow.Cell(mapResult.nNonyusakiCode).GetValue(Of String)().Trim(), "")

                                '出荷在庫場所 （任意）
                                'STRAMMIC.ITEMMより取得
                                shipstocklocation = ""
                                errMsg = ""
                                'If _oderStageRepo.GetShipStockLocation(customerCode, deliverycode, shipstocklocation, errMsg) = False Then
                                If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先情報区分    （任意）
                                customerinfotype = If(mapResult.nTorihikisakiJohoKubun > 0, xlRow.Cell(mapResult.nTorihikisakiJohoKubun).GetValue(Of String)().Trim(), "")

                                '情報区分       （任意）
                                'INFO_TYPE_MSTより取得
                                infotype = ""
                                errMsg = ""
                                If _oderStageRepo.GetInfoType(CustomerSettingId, customerinfotype, infotype, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '消込条件区分 ※順次/同月まで/同月内のみ     （任意）
                                'IMP_RULE_MSTより取得
                                reconciletype = 1
                                errMsg = ""
                                If _oderStageRepo.GetReconcileType(CustomerSettingId, FolderType, reconciletype, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '出荷先　   （必須）
                                'STRAMMIC.SECTMより取得
                                shipto = ""
                                errMsg = ""
                                If _oderStageRepo.GetShipTo(customerCode, deliverycode, shipto, errMsg) = False Then
                                    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    ErrFlg = True
                                End If

                                '請求先    (任意）
                                'STRAMMIC.SECTMより取得
                                billingto = ""
                                errMsg = ""
                                If _oderStageRepo.GetBillingTo(customerCode, billingto, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If



                                '取引先設定IDのPC   （必須）
                                'CUSTOMER_SETTING_MSTより取得
                                profitcenterCSM = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '品目NoのPC   （必須）
                                'STRAMMIC.USRDEFFLDFより取得
                                profitcenter = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '取引先設定IDのPCと同じPCのみ取込対象とする
                                If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                    'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                    fileidx += 1
                                    Continue For
                                End If


                                '-----------------
                                '桁チェック
                                '-----------------
                                '受注区分
                                ordertype = SafeVarcharLength(ordertype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：受注区分が桁数超過のためトリミングされました。")
                                End If
                                '分割区分
                                proratedtype = SafeVarcharLength(proratedtype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：分割区分が桁数超過のためトリミングされました。")
                                End If
                                '客先発注番号
                                customerorderNo = SafeVarcharLength(customerorderNo, 40, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が桁数超過のためトリミングされました。")
                                End If
                                '需要数
                                demandqty = Convert.ToDecimal(SafeVarcharLength(demandqty.ToString(), 10, isTruncated))
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が桁数超過のためトリミングされました。")
                                End If
                                ''日割前受注数
                                'predailyorderqty = Convert.ToDecimal(SafeVarcharLength(predailyorderqty.ToString(), 10, isTruncated))
                                'If isTruncated = True Then
                                '    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：日割前受注数が桁数超過のためトリミングされました。")
                                'End If
                                '自社予測フラグ
                                selffcstflag = SafeVarcharLength(selffcstflag, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測フラグが桁数超過のためトリミングされました。")
                                End If
                                '自社予測削除フラグ
                                selffcstdeleteflag = SafeVarcharLength(selffcstdeleteflag, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：自社予測削除フラグが桁数超過のためトリミングされました。")
                                End If
                                '通貨コード
                                currencycode = SafeVarcharLength(currencycode, 3, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：通貨コードが桁数超過のためトリミングされました。")
                                End If
                                '客先品目No
                                customeritemNo = SafeVarcharLength(customeritemNo, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先品目Noが桁数超過のためトリミングされました。")
                                End If
                                '製品コード
                                productcode = SafeVarcharLength(productcode, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：製品コードが桁数超過のためトリミングされました。")
                                End If
                                '品目No
                                itemNo = SafeVarcharLength(itemNo, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：品目Noが桁数超過のためトリミングされました。")
                                End If
                                '需要単位
                                demandunit = SafeVarcharLength(demandunit, 4, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要単位が桁数超過のためトリミングされました。")
                                End If
                                'コメント
                                remarks = SafeVarcharLength(remarks, 45, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：コメントが桁数超過のためトリミングされました。")
                                End If
                                '納入先コード
                                deliverycode = SafeVarcharLength(deliverycode, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：納入先コードが桁数超過のためトリミングされました。")
                                End If
                                '出荷在庫場所
                                shipstocklocation = SafeVarcharLength(shipstocklocation, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：出荷在庫場所が桁数超過のためトリミングされました。")
                                End If
                                '取引先情報区分
                                customerinfotype = SafeVarcharLength(customerinfotype, 50, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：取引先情報区分が桁数超過のためトリミングされました。")
                                End If
                                '情報区分
                                infotype = SafeVarcharLength(infotype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：情報区分が桁数超過のためトリミングされました。")
                                End If
                                '消込条件区分
                                reconciletype = SafeVarcharLength(reconciletype, 1, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：消込条件区分が桁数超過のためトリミングされました。")
                                End If
                                '出荷先
                                shipto = SafeVarcharLength(shipto, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：出荷先が桁数超過のためトリミングされました。")
                                End If
                                '請求先
                                billingto = SafeVarcharLength(billingto, 25, isTruncated)
                                If isTruncated = True Then
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：請求先が桁数超過のためトリミングされました。")
                                End If
                                '-----------------


                                'ここまででエラーフラグがあれば登録しない
                                If ErrFlg = True Then

                                    'ErrCustomerCode = customerCode
                                    'ErrTorikomiFile = TorikomiFile

                                    fileidx += 1
                                    errcnt += 1
                                    ErrFileFlg = True
                                    Continue For

                                End If

                                '受注ワーク登録用リストへ格納
                                rowsForTemp2.Add(New OrdersStageRow With {
                                    .CustomerSettingId = CustomerSettingId,
                                    .CustomerCode = customerCode,
                                    .BillingTo = billingto,
                                    .CustomerOrderNo = customerorderNo,
                                    .DemandStatus = demandstatus,
                                    .ShipTo = shipto,
                                    .OrderDate = orderDate,
                                    .DueDate = FormatDate(dueDate),
                                    .CustomerItemNo = customeritemNo,
                                    .ItemNo = itemNo,
                                    .DemandQty = demandqty,
                                    .DemandUnit = demandunit,
                                    .CurrencyCode = currencycode,
                                    .ShipStockLocation = shipstocklocation,
                                    .CompanyId = "1000",
                                    .ProductCode = productcode,
                                    .BillingStandard = "S",
                                    .ShipProcessType = shipprocesstype,
                                    .DeliveryInstrFlag = deliveryinstrflag,
                                    .Remarks = remarks,
                                    .DeliveryCode = deliverycode,
                                    .TotalShipQty = totalshipqty,
                                    .TransportMethod = "2",
                                    .PreDailyOrderQty = predailyorderqty,
                                    .PreDailyDeliveryDate = predailydeliveryDate,
                                    .ImpFileStageId = impfilestageId,
                                    .OrderType = ordertype,
                                    .ProratedType = proratedtype,
                                    .CustomerInfoType = customerinfotype,
                                    .InfoType = infotype,
                                    .SelfFcstFlag = selffcstflag,
                                    .SelfFcstDeleteFlag = selffcstdeleteflag,
                                    .ReconcileType = reconciletype,
                                    .ImpRunId = newId,
                                    .Status = "IMPORTED",
                                    .ActiveFlag = "Y",
                                    .CreatedAt = Now,
                                    .CreatedUserId = UserId,
                                    .CreatedPgId = pgId,
                                    .UpdatedAt = Now,
                                    .UpdatedUserId = UserId,
                                    .UpdatedPgId = pgId
                                })

                                fileidx += 1

                            Next

                        End Using


                    ElseIf mapResult.FormatType = "MATRIX" Then

                        Dim orderDateErrFlg As Boolean = False
                        Dim HeaderErrFlg As Boolean = False

                        '希望納期がマッピングマスタにない場合は、この後の処理が成立しないため処理を中断する。
                        If mapResult.mKibouNouki = "" Then
                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　マッピングマスタに不備があります。(希望納期がマッピングマスタ未登録)")
                            'Continue For
                            Return False
                        End If

                        'ワークブックを作成
                        Using objWorkBook As New ClosedXML.Excel.XLWorkbook(strWorkFile)

                            'ワークシートを作成
                            Dim objSheet As ClosedXML.Excel.IXLWorksheet

                            'ワークシート指定があれば指定
                            If mapResult.DefaultSheetName <> "" Then
                                objSheet = objWorkBook.Worksheet(mapResult.DefaultSheetName)
                            Else
                                objSheet = objWorkBook.Worksheet(1)
                            End If

                            '--------------------
                            'ヘッダー部にある項目
                            '--------------------
                            '初期化
                            strTempDate = ""  '日付検証用
                            orderDate = Nothing

                            '受注日   (任意)
                            strTempDate = If(mapResult.mJutyuuBi <> "", objSheet.Cell(mapResult.mJutyuuBi).GetValue(Of String)().Trim(), "")
                            If String.IsNullOrEmpty(strTempDate) Then
                                orderDate = CDate("1900/01/01")
                            Else
                                ' 日付変換を試みる（yyyy/MM/dd形式）
                                If Not DateTime.TryParseExact(strTempDate, formats,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        orderDate) Then
                                    ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                    orderDate = CDate("1900/01/01") ' または特定の既定値
                                    'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKibouNouki).Address.ColumnLetter & (objSheet.Cell(mKibouNouki).Address.RowNumber - 1)} ：受注日が不正な値です。")
                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mJutyuuBi).Address.ColumnLetter & (objSheet.Cell(mapResult.mJutyuuBi).Address.RowNumber)} ：受注日が不正な値です。")
                                    ' ErrFlg = True
                                    orderDateErrFlg = True
                                End If
                            End If

                            'データの最終行を取得
                            Dim EdRowNum As Integer = objSheet.LastRowUsed().RowNumber()

                            '納期の行番号をセットする 例:B4=4
                            Dim StRowNum As Integer = objSheet.Cell(mapResult.mKibouNouki).Address.RowNumber
                            '納期の列番号をセットする 例:B4=2
                            Dim StColNum As Integer = objSheet.Cell(mapResult.mKibouNouki).Address.ColumnNumber

                            '行飛ばし数をセットする
                            Dim stepRow As Integer = 9

                            '納期の行をセット
                            Dim xlRow = objSheet.Row(StRowNum)
                            ' 納期の行の最後のセルを取得して最終列とする
                            Dim lastCell = xlRow.LastCellUsed()
                            Dim EdColNum As Integer = lastCell.Address.ColumnNumber


                            '客先品番行(需要数と兼用)へポインタを移動
                            StRowNum += 1

                            Dim ridx As Integer = 0
                            For intRowidx As Integer = StRowNum To EdRowNum Step stepRow

                                '初期化
                                strTempDate = ""  '日付検証用
                                strQtyValue = ""  '数値検証用
                                customerorderNo = ""
                                'orderDate = Nothing
                                dueDate = Nothing
                                customeritemNo = ""
                                demandqty = 0
                                demandunit = ""
                                currencycode = ""
                                productcode = ""
                                remarks = ""
                                deliverycode = ""
                                predailyorderqty = 0
                                predailydeliveryDate = Nothing
                                ordertype = 0
                                proratedtype = 1
                                customerinfotype = ""
                                selffcstflag = ""
                                selffcstdeleteflag = ""
                                shipto = ""
                                billingto = ""
                                itemNo = ""
                                demandstatus = ""
                                shipprocesstype = ""
                                deliveryinstrflag = ""
                                totalshipqty = 0
                                shipstocklocation = ""
                                infotype = ""
                                reconciletype = 1
                                profitcenterCSM = ""
                                profitcenter = ""

                                If orderDateErrFlg = True Then
                                    HeaderErrFlg = True
                                Else
                                    HeaderErrFlg = False
                                End If


                                '客先品目No   (任意だが、MATRIX形式は製品コードの項目がマトリックス共通表サンプルに存在しないので客先品目Noがないと品目Noが取得できない)
                                Dim cell1 As Integer = 0
                                'Dim cellname As String = ""
                                If Not String.IsNullOrEmpty(mapResult.mKyakusakiHinmokuNo) Then
                                    xlRow = objSheet.Row(objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.RowNumber + (stepRow * ridx))
                                    cell1 = objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnNumber
                                    'cellname = objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber
                                    customeritemNo = If(mapResult.mKyakusakiHinmokuNo <> "", xlRow.Cell(cell1).GetValue(Of String)().Trim(), "")
                                Else
                                    customeritemNo = ""
                                End If

                                '製品コード  （マトリックス共通表サンプルに存在しない）
                                productcode = ""

                                '品目No   （必須）
                                'STRAMMIC.PRDSLSODRMより取得
                                itemNo = ""
                                errMsg = ""
                                If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                    'If mapResult.mKyakusakiHinmokuNo = "" Then
                                    '    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　{errMsg} (マッピングマスタ未登録)")
                                    'Else
                                    '    errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                    'End If
                                    ''ErrFlg = True
                                    'HeaderErrFlg = True
                                End If


                                '取引先設定IDのPC   （必須）
                                'CUSTOMER_SETTING_MSTより取得
                                profitcenterCSM = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenterFromCSM(CustomerSettingId, profitcenterCSM, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If

                                '品目NoのPC   （必須）
                                'STRAMMIC.USRDEFFLDFより取得
                                profitcenter = ""
                                errMsg = ""
                                If _oderStageRepo.GetProfitCenter(itemNo, profitcenter, errMsg) = False Then
                                    'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                    'ErrFlg = True
                                End If


                                '取引先設定IDのPCと同じPCのみ取込対象とする
                                If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                    ''PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                    'fileidx += 1
                                    'Continue For
                                Else

                                    '品目No   （必須）
                                    'STRAMMIC.PRDSLSODRMより取得
                                    itemNo = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetProductCode(customerCode, customeritemNo, productcode, itemNo, errMsg) = False Then
                                        If mapResult.mKyakusakiHinmokuNo = "" Then
                                            errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　{errMsg} (マッピングマスタ未登録)")
                                        Else
                                            errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        End If
                                        'ErrFlg = True
                                        HeaderErrFlg = True
                                    End If

                                    '需要単位   （任意）
                                    'STRAMMIC.ITEMMより取得
                                    demandunit = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetDemandUnit(productcode, demandunit, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'ErrFlg = True
                                    End If

                                    '出荷在庫場所 （任意）
                                    'STRAMMIC.ITEMMより取得
                                    shipstocklocation = ""
                                    errMsg = ""
                                    'If _oderStageRepo.GetShipStockLocation(customerCode, deliverycode, shipstocklocation, errMsg) = False Then
                                    If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'ErrFlg = True
                                    End If


                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '客先品目No
                                    customeritemNo = SafeVarcharLength(customeritemNo, 45, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：客先品目Noが桁数超過のためトリミングされました。")
                                    End If
                                    '製品コード
                                    productcode = SafeVarcharLength(productcode, 45, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：製品コードが桁数超過のためトリミングされました。")
                                    End If
                                    '品目No
                                    itemNo = SafeVarcharLength(itemNo, 45, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：品目Noが桁数超過のためトリミングされました。")
                                    End If
                                    '需要単位
                                    demandunit = SafeVarcharLength(demandunit, 4, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：需要単位が桁数超過のためトリミングされました。")
                                    End If
                                    '出荷在庫場所
                                    shipstocklocation = SafeVarcharLength(shipstocklocation, 25, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mapResult.mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：出荷在庫場所が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                End If


                                '納期の列番号を始点として最終列までループ(横へループ)
                                For intColIdx As Integer = StColNum To EdColNum


                                    '取引先設定IDのPCと同じPCのみ取込対象とする
                                    If String.IsNullOrEmpty(profitcenterCSM) OrElse String.IsNullOrEmpty(profitcenter) OrElse profitcenterCSM <> profitcenter Then
                                        'PCが違う場合は取込しない、エラーメッセージなし、ファイル移動もなし
                                        fileidx += 1
                                        Continue For
                                    End If


                                    If HeaderErrFlg = True Then
                                        ErrFlg = True
                                    Else
                                        ErrFlg = False
                                    End If

                                    '受注日
                                    '※受注日はヘッダー部で取得済み

                                    '需要数   (必須)
                                    If Not String.IsNullOrEmpty(mapResult.mJuyouSuu) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mJuyouSuu).Address.RowNumber + (stepRow * ridx))
                                        strQtyValue = If(mapResult.mJuyouSuu <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        strQtyValue = ""
                                    End If

                                    If String.IsNullOrEmpty(strQtyValue) Then
                                        'MATRIXでは、需要数が空だった場合は登録対象外とする
                                        If mapResult.mJuyouSuu = "" Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　需要数が取得できません。 (マッピングマスタ未登録)")
                                        Else
                                            'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：需要数が不正な値です。")
                                        End If
                                        Continue For
                                    End If
                                    If Not Decimal.TryParse(strQtyValue, demandqty) Then
                                        If mapResult.mJuyouSuu = "" Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　需要数が取得できません。 (マッピングマスタ未登録)")
                                        Else
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：需要数が不正な値です。")
                                        End If
                                        ErrFlg = True
                                    End If

                                    'If String.IsNullOrEmpty(strQtyValue) Then
                                    '    '必須チェック
                                    '    If mjuyouSuu = "" Then
                                    '        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　需要数が取得できません。")
                                    '    Else
                                    '        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：需要数が空です。")
                                    '    End If
                                    '    ErrFlg = True
                                    'ElseIf Not Decimal.TryParse(strQtyValue, demandqty) Then
                                    '    '数値チェック
                                    '    If mjuyouSuu = "" Then
                                    '        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　需要数が取得できません。")
                                    '    Else
                                    '        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：需要数が不正な値です。")
                                    '    End If
                                    '    ErrFlg = True
                                    'End If

                                    '日割前受注数 ※需要数をセット
                                    predailyorderqty = demandqty

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '需要数
                                    demandqty = Convert.ToDecimal(SafeVarcharLength(demandqty.ToString(), 10, isTruncated))
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：需要数が桁数超過のためトリミングされました。")
                                    End If
                                    ''日割前受注数
                                    'predailyorderqty = Convert.ToDecimal(SafeVarcharLength(predailyorderqty.ToString(), 10, isTruncated))
                                    'If isTruncated = True Then
                                    '    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：日割前受注数が桁数超過のためトリミングされました。")
                                    'End If
                                    '-----------------

                                    'フォルダタイプで処理分岐
                                    If FolderType = 4 Then

                                        '受注区分   (混在フォルダの場合は必須)
                                        If Not String.IsNullOrEmpty(mapResult.mJutyuKubun) Then
                                            xlRow = objSheet.Row(objSheet.Cell(mapResult.mJutyuKubun).Address.RowNumber + (stepRow * ridx))
                                            strQtyValue = If(mapResult.mJutyuKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                        Else
                                            strQtyValue = ""
                                        End If
                                        If String.IsNullOrEmpty(strQtyValue) Then
                                            '必須チェック
                                            If mapResult.mJutyuKubun = "" Then
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　受注区分が取得できません。 (マッピングマスタ未登録)")
                                            Else
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：受注区分が空です。")
                                            End If
                                            ErrFlg = True
                                        ElseIf Not Decimal.TryParse(strQtyValue, ordertype) Then
                                            '数値チェック
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：受注区分が不正な値です。")
                                            ErrFlg = True
                                        End If

                                        '-----------------
                                        '桁チェック
                                        '-----------------
                                        '受注区分
                                        ordertype = SafeVarcharLength(ordertype, 1, isTruncated)
                                        If isTruncated = True Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：受注区分が桁数超過のためトリミングされました。")
                                        End If
                                        '-----------------

                                        '分割区分   (混在フォルダの場合は必須)
                                        If Not String.IsNullOrEmpty(mapResult.mBunkatuKubun) Then
                                            xlRow = objSheet.Row(objSheet.Cell(mapResult.mBunkatuKubun).Address.RowNumber + (stepRow * ridx))
                                            strQtyValue = If(mapResult.mBunkatuKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                        Else
                                            strQtyValue = ""
                                        End If
                                        If String.IsNullOrEmpty(strQtyValue) Then
                                            '必須チェック
                                            If mapResult.mBunkatuKubun = "" Then
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　分割区分が取得できません。 (マッピングマスタ未登録)")
                                            Else
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：分割区分が空です。")
                                            End If
                                            ErrFlg = True
                                        ElseIf Not Decimal.TryParse(strQtyValue, proratedtype) Then
                                            '数値チェック
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：分割区分が不正な値です。")
                                            ErrFlg = True
                                        End If

                                        '-----------------
                                        '桁チェック
                                        '-----------------
                                        '分割区分
                                        proratedtype = SafeVarcharLength(proratedtype, 1, isTruncated)
                                        If isTruncated = True Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：分割区分が桁数超過のためトリミングされました。")
                                        End If
                                        '-----------------

                                    Else

                                        '受注区分   (任意)
                                        ordertype = FolderType

                                        '分割区分   (任意)
                                        'IMP_RULE_MSTより取得
                                        proratedtype = 1
                                        errMsg = ""
                                        If _oderStageRepo.GetProratedType(CustomerSettingId, FolderType, proratedtype, errMsg) = False Then
                                            'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：分割区分が数値ではない、または空です。")
                                            'ErrFlg = True
                                        End If

                                        '-----------------
                                        '桁チェック
                                        '-----------------
                                        '受注区分
                                        ordertype = SafeVarcharLength(ordertype, 1, isTruncated)
                                        If isTruncated = True Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　受注区分が桁数超過のためトリミングされました。")
                                        End If
                                        '分割区分
                                        proratedtype = SafeVarcharLength(proratedtype, 1, isTruncated)
                                        If isTruncated = True Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　分割区分が桁数超過のためトリミングされました。")
                                        End If
                                        '-----------------

                                    End If

                                    '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                                    If Not String.IsNullOrEmpty(mapResult.mKyakusakiHattyuNo) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mKyakusakiHattyuNo).Address.RowNumber + (stepRow * ridx))
                                        customerorderNo = If((mapResult.mKyakusakiHattyuNo) <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        customerorderNo = ""
                                    End If
                                    'If folderType = 2 OrElse folderType = 3 Then
                                    If ordertype = 2 OrElse ordertype = 3 Then
                                        'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                        If String.IsNullOrEmpty(customerorderNo) Then
                                            If mapResult.mKyakusakiHattyuNo = "" Then
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　客先発注番号が取得できません。 (マッピングマスタ未登録)")
                                            Else
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：客先発注番号が空です。")
                                            End If
                                            ErrFlg = True
                                        End If
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '客先発注番号
                                    customerorderNo = SafeVarcharLength(customerorderNo, 40, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：客先発注番号が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '希望納期   (必須)
                                    If Not String.IsNullOrEmpty(mapResult.mKibouNouki) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mKibouNouki).Address.RowNumber)
                                        strTempDate = If(mapResult.mKibouNouki <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        strTempDate = ""
                                    End If
                                    If String.IsNullOrEmpty(strTempDate) Then
                                        dueDate = CDate("1900/01/01")
                                        If mapResult.mKibouNouki = "" Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　希望納期が取得できません。 (マッピングマスタ未登録)")
                                        Else
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：希望納期が空です。")
                                        End If
                                        ErrFlg = True
                                    Else
                                        If Not DateTime.TryParseExact(strTempDate, formats,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        dueDate) Then
                                            ' 変換に失敗した場合（空文字や不正な値など）のデフォルト値
                                            dueDate = CDate("1900/01/01")
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：希望納期が不正な値です。")
                                            ErrFlg = True
                                        End If
                                    End If

                                    '日割前納期 ※希望納期をセット
                                    predailydeliveryDate = dueDate

                                    '2026/06/29 日割前納期に希望納期をセットした後に希望納期の稼働日チェック
                                    Dim cal = New CalenderRepository(Utils.GetConnectionString())
                                    Dim tdt = New Date
                                    tdt = dueDate
                                    dueDate = cal.AddWorkingDays2("00001", tdt, 0)
                                    '--

                                    '自社予測フラグ   (任意)
                                    If Not String.IsNullOrEmpty(mapResult.mJishaYosokuFlag) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mJishaYosokuFlag).Address.RowNumber + (stepRow * ridx))
                                        selffcstflag = xlRow.Cell(intColIdx).GetValue(Of String)().Trim()
                                    Else
                                        selffcstflag = ""
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '自社予測フラグ
                                    selffcstflag = SafeVarcharLength(selffcstflag, 1, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：自社予測フラグが桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '自社予測削除フラグ   (自社予測フラグ = Yの時は必須)
                                    If Not String.IsNullOrEmpty(mapResult.mJishaYosokuDelFlag) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mJishaYosokuDelFlag).Address.RowNumber + (stepRow * ridx))
                                        selffcstdeleteflag = If(mapResult.mJishaYosokuDelFlag <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        selffcstdeleteflag = ""
                                    End If
                                    If selffcstflag = "Y" AndAlso String.IsNullOrEmpty(selffcstdeleteflag) Then
                                        If mapResult.mJishaYosokuDelFlag = "" Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　自社予測削除フラグが取得できません。 (マッピングマスタ未登録)")
                                        Else
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：自社予測削除フラグが空です。")
                                        End If
                                        ErrFlg = True
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '自社予測削除フラグ
                                    selffcstdeleteflag = SafeVarcharLength(selffcstdeleteflag, 1, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：自社予測削除フラグが桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '需要ステイタス    （固定値）
                                    demandstatus = If(ordertype = 1, "F", "O")

                                    '累計出荷数    （固定値）
                                    If ordertype = 1 Then
                                        totalshipqty = Nothing
                                    Else
                                        totalshipqty = 0
                                    End If

                                    '出荷プロセスタイプ    （固定値）
                                    Select Case ordertype
                                        Case 1
                                            shipprocesstype = "O"
                                        Case 2
                                            shipprocesstype = "E"
                                        Case 3
                                            shipprocesstype = "K"
                                    End Select

                                    '納入指示フラグ    （固定値）
                                    deliveryinstrflag = If(ordertype = 3, "Y", "N")

                                    '通貨コード  （任意）
                                    If mapResult.mTukaCode <> "" Then
                                        '取得ファイルに存在
                                        If Not String.IsNullOrEmpty(mapResult.mTukaCode) Then
                                            xlRow = objSheet.Row(objSheet.Cell(mapResult.mTukaCode).Address.RowNumber + (stepRow * ridx))
                                            currencycode = If(mapResult.mTukaCode <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                        Else
                                            currencycode = ""
                                        End If
                                    Else
                                        '取得できない場合はSTRAMMIC.SECTMより取得
                                        currencycode = ""
                                        errMsg = ""
                                        If _oderStageRepo.GetCurrencyCode(customerCode, currencycode, errMsg) = False Then
                                            'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                            'ErrFlg = True
                                        End If
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '通貨コード
                                    currencycode = SafeVarcharLength(currencycode, 3, isTruncated)
                                    If isTruncated = True Then
                                        If mapResult.mTukaCode <> "" Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：通貨コードが桁数超過のためトリミングされました。")
                                        Else
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　通貨コードが桁数超過のためトリミングされました。")
                                        End If
                                    End If
                                    '-----------------

                                    'コメント   （任意）
                                    If Not String.IsNullOrEmpty(mapResult.mComment) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mComment).Address.RowNumber + (stepRow * ridx))
                                        remarks = If(mapResult.mComment <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        remarks = ""
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    'コメント
                                    remarks = SafeVarcharLength(remarks, 45, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：コメントが桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '納入先コード   （任意）
                                    If Not String.IsNullOrEmpty(mapResult.mNonyusakiCode) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mNonyusakiCode).Address.RowNumber + (stepRow * ridx))
                                        deliverycode = If(mapResult.mNonyusakiCode <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        deliverycode = ""
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '納入先コード
                                    deliverycode = SafeVarcharLength(deliverycode, 25, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：納入先コードが桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '出荷先　　   （必須）
                                    'STRAMMIC.SECTMより取得
                                    shipto = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetShipTo(customerCode, deliverycode, shipto, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　{errMsg}")
                                        errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & intRowidx} ：{errMsg}")
                                        ErrFlg = True
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '出荷先
                                    shipto = SafeVarcharLength(shipto, 25, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & intRowidx} ：出荷先が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '取引先情報区分 （任意）
                                    If Not String.IsNullOrEmpty(mapResult.mTorihikisakiJohoKubun) Then
                                        xlRow = objSheet.Row(objSheet.Cell(mapResult.mTorihikisakiJohoKubun).Address.RowNumber + (stepRow * ridx))
                                        customerinfotype = If(mapResult.mTorihikisakiJohoKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                    Else
                                        customerinfotype = ""
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '取引先情報区分
                                    customerinfotype = SafeVarcharLength(customerinfotype, 50, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：取引先情報区分が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '情報区分 （任意）
                                    'INFO_TYPE_MSTより取得
                                    infotype = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetInfoType(CustomerSettingId, customerinfotype, infotype, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'ErrFlg = True
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '情報区分
                                    infotype = SafeVarcharLength(infotype, 1, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：情報区分が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '消込条件区分 ※順次/同月まで/同月内のみ （任意）
                                    'IMP_RULE_MSTより取得
                                    reconciletype = 1
                                    errMsg = ""
                                    If _oderStageRepo.GetReconcileType(CustomerSettingId, FolderType, reconciletype, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'ErrFlg = True
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '消込条件区分
                                    reconciletype = SafeVarcharLength(reconciletype, 1, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & intRowidx} ：消込条件区分が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    '請求先    (任意）
                                    'STRAMMIC.SECTMより取得
                                    billingto = ""
                                    errMsg = ""
                                    If _oderStageRepo.GetBillingTo(customerCode, billingto, errMsg) = False Then
                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
                                        'ErrFlg = True
                                    End If

                                    '-----------------
                                    '桁チェック
                                    '-----------------
                                    '請求先
                                    billingto = SafeVarcharLength(billingto, 25, isTruncated)
                                    If isTruncated = True Then
                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & intRowidx} ：請求先が桁数超過のためトリミングされました。")
                                    End If
                                    '-----------------

                                    'ここまででエラーフラグがあれば登録しない
                                    If ErrFlg = True Then

                                        'ErrCustomerCode = customerCode
                                        'ErrTorikomiFile = TorikomiFile
                                        fileidx += 1
                                        errcnt += 1
                                        ErrFileFlg = True
                                        Continue For

                                    End If


                                    '受注ワーク登録用リストへ格納
                                    rowsForTemp2.Add(New OrdersStageRow With {
                                        .CustomerSettingId = CustomerSettingId,
                                        .CustomerCode = customerCode,
                                        .BillingTo = billingto,
                                        .CustomerOrderNo = customerorderNo,
                                        .DemandStatus = demandstatus,
                                        .ShipTo = shipto,
                                        .OrderDate = orderDate,
                                        .DueDate = FormatDate(dueDate),
                                        .CustomerItemNo = customeritemNo,
                                        .ItemNo = itemNo,
                                        .DemandQty = demandqty,
                                        .DemandUnit = demandunit,
                                        .CurrencyCode = currencycode,
                                        .ShipStockLocation = shipstocklocation,
                                        .CompanyId = "1000",
                                        .ProductCode = productcode,
                                        .BillingStandard = "S",
                                        .ShipProcessType = shipprocesstype,
                                        .DeliveryInstrFlag = deliveryinstrflag,
                                        .Remarks = remarks,
                                        .DeliveryCode = deliverycode,
                                        .TotalShipQty = totalshipqty,
                                        .TransportMethod = "2",
                                        .PreDailyOrderQty = predailyorderqty,
                                        .PreDailyDeliveryDate = predailydeliveryDate,
                                        .ImpFileStageId = impfilestageId,
                                        .OrderType = ordertype,
                                        .ProratedType = proratedtype,
                                        .CustomerInfoType = customerinfotype,
                                        .InfoType = infotype,
                                        .SelfFcstFlag = selffcstflag,
                                        .SelfFcstDeleteFlag = selffcstdeleteflag,
                                        .ReconcileType = reconciletype,
                                        .ImpRunId = newId,
                                        .Status = "IMPORTED",
                                        .ActiveFlag = "Y",
                                        .CreatedAt = Now,
                                        .CreatedUserId = UserId,
                                        .CreatedPgId = pgId,
                                        .UpdatedAt = Now,
                                        .UpdatedUserId = UserId,
                                        .UpdatedPgId = pgId
                                    })

                                    fileidx += 1

                                Next

                                ridx += 1

                            Next

                        End Using

                    End If

            End Select

            Return True

        End Function

        Public Shared Function Orders_Saved(ByVal tran As OracleTransaction,
                                            ByVal CustomerSettingId As Long,
                                            ByVal FolderType As Integer,
                                            ByVal blnHandFlag As Boolean,
                                            ByVal rowsForTemp2 As List(Of OrdersStageRow)) As OrderStageImport

            Dim _oderStageRepo As New OrderStageRepository(Utils.GetConnectionString())
            Dim _impFileStageRepo As New ImpFilesStageRepository(Utils.GetConnectionString())

            Dim result As New OrderStageImport()

            Dim cnt As Integer = 0

            '-----------------
            '受注ワーク削除 ※customerSettingId、folderType単位で削除
            '-----------------
            _oderStageRepo.DeleteRange(tran, CustomerSettingId, FolderType)

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
        ''' 配列に格納された取込データをワークテーブルに一括保存します。
        ''' </summary>
        Private Shared Sub SaveImportDataToWorkTable(
            ByVal tran As OracleTransaction,
            ByVal impDataList() As ImportDataType,
            ByVal impFileStageId As Long,
            ByVal newId As Integer,
            ByVal UserId As String,
            ByVal pgId As String)

            ' 配列が空、またはデータが1件もない場合は処理を抜ける
            If impDataList Is Nothing OrElse impDataList.Length = 0 Then Return
            ' 配列が初期化（0）のままで、かつ有効なデータが1件もセットされていなければスキップ
            If impDataList.Length = 1 AndAlso String.IsNullOrEmpty(impDataList(0).customeritemNo) Then Return

            ''前回の取込データをクリアする
            'Using delCmd As New OracleCommand("DELETE FROM yamaha_imp_orders_test", tran.Connection)
            '    'delCmd.Transaction = tran
            '    delCmd.ExecuteNonQuery()
            'End Using

            Dim nowTime As DateTime = DateTime.Now

            Dim sql As String = "
                        INSERT INTO yamaha_imp_orders_test (
                            imp_file_stage_id, hinmoku_gyo_no, order_gyo_no, customer_code, siyosha, status, customer_item_no, nonyuplat,
                            yokisyuuyousuu, yokibangou, ordersikibetu_no, nonyusijibi, nonyujikan,
                            nonyusijisu, cardkubun, naijikubun, icdenpyo_no, nohinsho_no,
                            publication_date, publication_time, imp_run_id, active_flag,
                            created_at, created_user_id, created_pg_id
                        ) VALUES (
                            :p_imp_file_stage_id, :p_hinmoku_gyo_no, :p_order_gyo_no, :p_customer_code, :p_siyosha, :p_status, :p_customer_item_no, :p_nonyuplat,
                            :p_yokisyuuyousuu, :p_yokibangou, :p_ordersikibetu_no, :p_nonyusijibi, :p_nonyujikan,
                            :p_nonyusijisu, :p_cardkubun, :p_naijikubun, :p_icdenpyo_no, :p_nohinsho_no,
                            :p_publication_date, :p_publication_time, :p_imp_run_id, :p_active_flag,
                            :p_created_at, :p_created_user_id, :p_created_pg_id
                        )"

            Using cmd As New OracleCommand(sql, tran.Connection)
                'cmd.Transaction = tran
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Clear()

                ' 処理高速化のため、ループ外で一度だけ型を定義
                cmd.Parameters.Add("p_imp_file_stage_id", OracleDbType.Int64)
                cmd.Parameters.Add("p_hinmoku_gyo_no", OracleDbType.Int32)
                cmd.Parameters.Add("p_order_gyo_no", OracleDbType.Int32)
                cmd.Parameters.Add("p_customer_code", OracleDbType.Varchar2, 25)
                cmd.Parameters.Add("p_siyosha", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_status", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_customer_item_no", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nonyuplat", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_yokisyuuyousuu", OracleDbType.Int32)
                cmd.Parameters.Add("p_yokibangou", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_ordersikibetu_no", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nonyusijibi", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nonyujikan", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nonyusijisu", OracleDbType.Int32)
                cmd.Parameters.Add("p_cardkubun", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_naijikubun", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_icdenpyo_no", OracleDbType.Varchar2)
                cmd.Parameters.Add("p_nohinsho_no", OracleDbType.Varchar2)

                cmd.Parameters.Add("p_publication_date", OracleDbType.Date)
                cmd.Parameters.Add("p_publication_time", OracleDbType.Int32)
                cmd.Parameters.Add("p_imp_run_id", OracleDbType.Long)
                cmd.Parameters.Add("p_active_flag", OracleDbType.Char, 1)
                cmd.Parameters.Add("p_created_at", OracleDbType.Date)
                cmd.Parameters.Add("p_created_user_id", OracleDbType.Varchar2, 9)
                cmd.Parameters.Add("p_created_pg_id", OracleDbType.Varchar2, 150)

                ' 配列の全要素をループしてインサートを実行
                For i As Integer = 0 To UBound(impDataList)
                    cmd.Parameters("p_imp_file_stage_id").Value = impFileStageId
                    cmd.Parameters("p_hinmoku_gyo_no").Value = impDataList(i).hinmokugyoNo
                    cmd.Parameters("p_order_gyo_no").Value = impDataList(i).ordergyoNo
                    cmd.Parameters("p_customer_code").Value = SafeVarchar(impDataList(i).customercode, 25)
                    cmd.Parameters("p_siyosha").Value = impDataList(i).siyosha
                    cmd.Parameters("p_status").Value = impDataList(i).status
                    cmd.Parameters("p_customer_item_no").Value = impDataList(i).customeritemNo
                    cmd.Parameters("p_nonyuplat").Value = impDataList(i).nonyuplat
                    cmd.Parameters("p_yokisyuuyousuu").Value = If(String.IsNullOrEmpty(impDataList(i).yokisyuuyousuu), DBNull.Value, Convert.ToInt32(impDataList(i).yokisyuuyousuu))
                    cmd.Parameters("p_yokibangou").Value = impDataList(i).yokibangou
                    cmd.Parameters("p_ordersikibetu_no").Value = impDataList(i).ordersikibetuNo
                    cmd.Parameters("p_nonyusijibi").Value = impDataList(i).nonyusijibi
                    cmd.Parameters("p_nonyujikan").Value = impDataList(i).nonyujikan
                    cmd.Parameters("p_nonyusijisu").Value = If(String.IsNullOrEmpty(impDataList(i).nonyusijisu), DBNull.Value, Convert.ToInt32(impDataList(i).nonyusijisu))
                    cmd.Parameters("p_cardkubun").Value = impDataList(i).cardkubun
                    cmd.Parameters("p_naijikubun").Value = impDataList(i).naijikubun
                    cmd.Parameters("p_icdenpyo_no").Value = impDataList(i).icdenpyoNo
                    cmd.Parameters("p_nohinsho_no").Value = impDataList(i).nohinshoNo

                    cmd.Parameters("p_publication_date").Value = Date.ParseExact(impDataList(i).hakkobi, "yyyyMMdd", Nothing)
                    cmd.Parameters("p_publication_time").Value = impDataList(i).hakkojikan
                    cmd.Parameters("p_imp_run_id").Value = newId
                    cmd.Parameters("p_active_flag").Value = "Y"
                    cmd.Parameters("p_created_at").Value = nowTime
                    cmd.Parameters("p_created_user_id").Value = SafeVarchar(UserId, 9)
                    cmd.Parameters("p_created_pg_id").Value = SafeVarchar(pgId, 150)


                    cmd.ExecuteNonQuery()
                Next
            End Using
        End Sub

    End Class

End Namespace
