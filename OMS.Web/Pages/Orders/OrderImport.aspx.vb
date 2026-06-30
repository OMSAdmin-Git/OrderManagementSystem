Imports System
Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Web.Http
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess.Client

Imports System.Globalization
Imports CsvHelper
Imports CsvHelper.Configuration
Imports WebGrease
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Office2010.Excel
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.Ajax.Utilities

Namespace Pages.Orders
    Public Class OrderImport
        Inherits System.Web.UI.Page

        Private _impFileRepo As ImpFilesRepository
        Private _impRunRepo As ImpRunRepository
        Private _oderStageRepo As OrderStageRepository
        Private _mappingRepo As MappingRepository
        Private _impFileStageRepo As ImpFilesStageRepository
        Private _folderRepo As FolderRepository
        Private _compRootResolved As String
        Private _compUserRoot As String

        Private Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

            ' DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            ' リポジトリの初期化
            _impFileRepo = New ImpFilesRepository(connStr)
            _impRunRepo = New ImpRunRepository(connStr)
            _oderStageRepo = New OrderStageRepository(connStr)
            _mappingRepo = New MappingRepository(connStr)
            _impFileStageRepo = New ImpFilesStageRepository(connStr)
            _folderRepo = New FolderRepository(connStr)

            ' COMPLETEDフォルダの解決・作成
            Dim rawComp As String = ConfigurationManager.AppSettings("CompletedFolderRoot")
            If String.IsNullOrWhiteSpace(rawComp) Then
                Throw New ConfigurationErrorsException("appSettings['CompletedFolderRoot'] が未定義です。Web.config を確認してください。")
            End If

            ' 相対/UNC/絶対のいずれでも対応（Utils.ResolvePath は前回ご提案のヘルパ、または既存の ResolveFolderPath を流用）
            _compRootResolved = Utils.ResolvePath(Me.Server, rawComp)

            ' フォルダが存在しなければ作成
            Utils.EnsureDirectory(_compRootResolved)

            Dim UserId As String = PageHelpers.GetUserId(Me)
            If String.IsNullOrWhiteSpace(UserId) Then
                UserId = "AMAGATA"
            End If
            If UserId.Length > 9 Then
                UserId = UserId.Substring(0, 9)
            End If
            '' ユーザーID取得
            ''Dim userId As String = GetCurrentUserId()
            'Dim userId As String = "AMAGATA"

            ' フォルダ名として安全化（Windowsパスの禁止文字などを除去／長さ制限）
            Dim safeUserId As String = Utils.SafeFolderName(UserId, maxLength:=32)

            ' ユーザー単位のCOMPLETEDルートを作成
            _compUserRoot = System.IO.Path.Combine(_compRootResolved, safeUserId)
            Utils.EnsureDirectory(_compUserRoot)

        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

            If Not IsPostBack Then

                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索リスト
                Dim loginUserId As String = PageHelpers.GetUserId(Me)
                Dim repo As New CustomerRepository(Utils.GetConnectionString())
                Dim customerCodeList As List(Of String) = repo.GetCustomerCodes(loginUserId)
                Dim profitCenterList As List(Of String) = repo.GetProfitCenters(loginUserId)
                Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

                Dim sbCustomerCode As New StringBuilder()
                For Each code As String In customerCodeList
                    sbCustomerCode.AppendFormat("<option value='{0}' />", code)
                Next
                lstSearchCustomerCode.InnerHtml = sbCustomerCode.ToString()

                Dim sbProfitCenter As New StringBuilder()
                For Each pc As String In profitCenterList
                    sbProfitCenter.AppendFormat("<option value='{0}' />", pc)
                Next
                lstSearchProfitCenter.InnerHtml = sbProfitCenter.ToString()

                Dim sbCustomerUnitName As New StringBuilder()
                For Each customerUnit As String In customerUnitNameList
                    sbCustomerUnitName.AppendFormat("<option value='{0}' />", customerUnit)
                Next
                lstSearchCustomerUnitName.InnerHtml = sbCustomerUnitName.ToString()

                ' データバインド
                gvImpFilesStage_Init()
                gvImportOrder_Init()

            End If

        End Sub

        ' 受注メニューボタン
        Protected Sub btnOrderMenu_Click(sender As Object, e As EventArgs)
            Response.Redirect("OrderMenu.aspx")
        End Sub

        ' 検索ボタン
        Protected Sub btnSearchGv_Click(sender As Object, e As EventArgs)
            Dim customerCode As String = NullIfWhite(txtSearchCustomerCode.Value)
            Dim customerName As String = NullIfWhite(txtSearchCustomerName.Value)
            Dim profitCenter As String = NullIfWhite(txtSearchProfitCenter.Value)
            Dim customerUnitName As String = NullIfWhite(txtSearchCustomerUnitName.Value)
            gvImpFilesStage_Init(customerCode, customerName, profitCenter, customerUnitName)
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_Click(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""

            lblImportResult.Text = ""
            lblImportError.Text = ""

            lblSaveResult.Text = ""
            lblSaveError.Text = ""

            gvImpFilesStage_Init()
            gvImportOrder_Init()

        End Sub

        ' GridViewデータバインド
        Private Sub gvImpFilesStage_Init(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing
        )
            Dim repo As New ImpFilesStageRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetImpFilesStage(
                                        customerCode:=customerCode,
                                        customerName:=customerName,
                                        profitCenter:=profitCenter,
                                        customerUnitName:=customerUnitName,
                                        prodMgmtUserId:=PageHelpers.GetUserId(Me))
            gvImpFilesStage.DataSource = dt
            gvImpFilesStage.DataBind()
        End Sub

        ' GridViewヘッダーバインド
        Protected Sub gvImpFilesStage_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvImpFilesStage.RowDataBound
            If e.Row.RowType = DataControlRowType.DataRow Then

                ' ハンド列
                Dim chkHand As CheckBox = TryCast(e.Row.FindControl("chkHandFlag"), CheckBox)
                If chkHand IsNot Nothing Then
                    chkHand.InputAttributes("onclick") = $"OMS.Grid.updateHeader('{gvImpFilesStage.ClientID}', 'chkHandFlagAll', 'chkHandFlag');"
                End If

                ' 処理対象列
                Dim chkImport As CheckBox = TryCast(e.Row.FindControl("chkOrderImport"), CheckBox)
                If chkImport IsNot Nothing Then
                    chkImport.InputAttributes("onclick") = $"OMS.Grid.updateHeader('{gvImpFilesStage.ClientID}', 'chkOrderImportAll', 'chkOrderImport');"
                End If

            End If
        End Sub

        ' 取込実行ボタン
        Protected Sub btnImportFile_Click(sender As Object, e As EventArgs)

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
            Dim errMsg As String


            ' DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            lblImportResult.Text = ""
            lblImportError.Text = ""
            lblSaveResult.Text = ""
            lblSaveError.Text = ""

            Dim results As New List(Of ImpFilesStageResult)()
            Dim ok As Boolean = False
            Dim errors As New List(Of String)()

            Dim successs As New List(Of String)()

            Dim resultCnt As Integer = 0
            Dim resultRowCnt As Integer = 0
            Dim resultAllCnt As Integer = 0

            Dim ReDrawFlg As Boolean = False

            '---------------------------------------------------------------
            'グリッド内のチェック状態チェック
            '---------------------------------------------------------------
            ' 取引先ファイル選択行を走査
            Dim groups As New Dictionary(Of String, List(Of GridViewRow))
            Dim selectedRows As New List(Of GridViewRow)

            For Each row As GridViewRow In gvImpFilesStage.Rows
                If row.RowType <> DataControlRowType.DataRow Then Continue For

                Dim chk As CheckBox = TryCast(row.FindControl("chkOrderImport"), CheckBox)
                If chk IsNot Nothing AndAlso chk.Checked Then
                    selectedRows.Add(row)
                    Dim settingId As String = gvImpFilesStage.DataKeys(row.RowIndex)("CustomerSettingId").ToString()

                    If Not groups.ContainsKey(settingId) Then
                        groups(settingId) = New List(Of GridViewRow)()
                    End If
                    groups(settingId).Add(row)
                End If
            Next

            If selectedRows.Count = 0 Then
                errors.Add("処理対象を選択してください。")
            Else
                ' 取引先ごとのチェック
                For Each kvp In groups
                    Dim rowsInGroup = kvp.Value

                    ' 同じフォルダ区分が2件以上選択されているか（重複チェック）
                    Dim duplicateGroup = rowsInGroup.GroupBy(Function(r) gvImpFilesStage.DataKeys(r.RowIndex)("FolderType").ToString()) _
                                      .Where(Function(g) g.Count() > 1).FirstOrDefault()

                    If duplicateGroup IsNot Nothing Then
                        errors.Add("フォルダ区分が同じファイルは、2件以上同時に処理できません。")
                        ' 最初に見つかった重複レコードの「1件目」の詳細を改行付きで追加
                        AddErrorDetails(errors, duplicateGroup.First())
                        Exit For
                    End If

                    '「4:混在」と「それ以外」の共存チェック
                    Dim hasType4 = rowsInGroup.Any(Function(r) gvImpFilesStage.DataKeys(r.RowIndex)("FolderType").ToString() = "4")
                    Dim hasOther = rowsInGroup.Any(Function(r) gvImpFilesStage.DataKeys(r.RowIndex)("FolderType").ToString() <> "4")

                    If hasType4 AndAlso hasOther Then
                        errors.Add("フォルダ区分が混在のファイルと混在以外のファイルは、同時に処理できません。")
                        ' 最初に見つかったレコードの「1件目」の詳細を改行付きで追加
                        AddErrorDetails(errors, rowsInGroup.First())
                        Exit For
                    End If
                Next
            End If
            '---------------------------------------------------------------


            If errors.Count = 0 Then


                '実行管理
                'IMP_RUNの新しいIDを取得
                Dim newId As Integer = 0
                'newId += 1
                'IMP_RUNに新規レコード追加
                Dim now As DateTime = DateTime.Now
                'Dim userId As String = (If(Context?.User?.Identity?.Name, "")).Trim()
                Dim UserId As String = PageHelpers.GetUserId(Me)
                If String.IsNullOrWhiteSpace(UserId) Then
                    UserId = "AMAGATA"
                End If
                If UserId.Length > 9 Then
                    UserId = UserId.Substring(0, 9)
                End If
                Dim pgId As String = "OrderImport(Execute)"

                'Dim rowsForTemp As New List(Of ImpRunRow) From {
                '    New ImpRunRow With {
                '        .ImpRunId = newId,
                '        .StartedAt = now,
                '        .Status = "RUNNING",
                '        .StartedUserId = UserId,
                '        .StartedPgId = pgId,
                '        .FileCount = 0,
                '        .RowCount = 0,
                '        .ErrorCount = 0
                '    }
                '}
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


                _impRunRepo.InsertRange(rowsForTemp)

                'IMP_RUNテーブルの新しいIDを取得
                'newId = _impRunRepo.GetImpRunId()
                newId = rowsForTemp(0).ImpRunId

                Dim previousId As Long = -1 ' 前回の取引先設定ID保持用

                Dim preImpFileStageId As Long = -1  '前回の取込ファイルID保持用

                Dim isTruncated As Boolean = False



                Dim ErrFlg As Boolean = False
                Dim ErrFileFlg As Boolean = False

                Dim tran As OracleTransaction = Nothing

                Dim customerCode As Integer
                Dim strWorkFile As String

                '取込ファイル保持用
                Dim TorikomiFile As String
                'エラー情報表示用
                Dim ErrCustomerCode As Integer
                Dim ErrTorikomiFile As String


                Using conn As New OracleConnection(connStr)
                    conn.Open()
                    Try

                        Dim customerSettingId As Long
                        Dim impfilestageId As Long
                        Dim folderType As Integer
                        'Dim customerCode As Integer

                        Dim cnt As Integer = 0
                        Dim errcnt As Integer = 0

                        ' 取引先ファイル選択行を走査
                        For Each row As GridViewRow In gvImpFilesStage.Rows

                            errcnt = 0

                            ' データ行以外はスキップ
                            If row.RowType <> DataControlRowType.DataRow Then
                                Continue For
                            End If

                            '処理対象 未選択行はスキップ
                            Dim chk As CheckBox = TryCast(row.FindControl("chkOrderImport"), CheckBox)
                            If chk Is Nothing OrElse Not chk.Checked Then
                                Continue For
                            End If

                            Dim idx As Integer = row.RowIndex
                            Dim keys = gvImpFilesStage.DataKeys(idx)
                            If keys Is Nothing Then
                                errors.Add($"Row {idx}：DataKeys未設定")
                                Continue For
                            End If

                            Dim rowsForTemp2 As New List(Of OrdersStageRow)

                            '取引先設定ID　GridViewから取得
                            'Dim customerSettingId As Long
                            customerSettingId = 0
                            Dim csidObj = keys("CustomerSettingId")
                            If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), customerSettingId) Then
                                errors.Add($"Row {idx}：CustomerSettingIdが不正")
                                Continue For
                            End If

                            '一時取込ファイルID　GridViewから取得
                            'Dim impfilestageId As Long
                            impfilestageId = 0
                            csidObj = keys("ImpFileStageId")
                            If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), impfilestageId) Then
                                errors.Add($"Row {idx}：ImpFileStageIdが不正")
                                Continue For
                            End If

                            'フォルダタイプ　GridViewから取得
                            'Dim folderType As Integer
                            folderType = 0
                            csidObj = keys("FolderType")
                            If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), folderType) Then
                                errors.Add($"Row {idx}：FolderTypeが不正")
                                Continue For
                            End If

                            '取引先コード　GridViewから取得
                            'Dim customerCode As Integer
                            'ErrCustomerCode = 0
                            customerCode = 0
                            csidObj = keys("CustomerCode")
                            If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), customerCode) Then
                                errors.Add($"Row {idx}：CustomerCodeが不正")
                                Continue For
                            End If

                            '消込フラグ
                            Dim reconcileFlag As String = ""
                            csidObj = keys("ReconcileFlag")
                            If csidObj IsNot Nothing Then
                                reconcileFlag = csidObj.ToString().Trim().ToUpper() ' 大文字に統一して空白除去
                            End If
                            If reconcileFlag <> "Y" AndAlso reconcileFlag <> "N" Then
                                errors.Add($"Row {idx}：ReconcileFlagが不正")
                                Continue For
                            End If

                            '内示消込フラグ
                            Dim fcstreconcileFlag As String = ""
                            csidObj = keys("FcstReconcileFlag")
                            If csidObj IsNot Nothing Then
                                fcstreconcileFlag = csidObj.ToString().Trim().ToUpper() ' 大文字に統一して空白除去
                            End If
                            If fcstreconcileFlag <> "Y" AndAlso fcstreconcileFlag <> "N" Then
                                errors.Add($"Row {idx}：FcstReconcileFlagが不正")
                                Continue For
                            End If

                            'ハンドフラグ　GridViewから取得
                            Dim chkHandFlag = TryCast(row.FindControl("chkHandFlag"), CheckBox)

                            ' [WORKフォルダパス]を取得
                            Dim strWorkFolder As String = keys("StagedFolderPath").ToString()
                            If strWorkFolder IsNot Nothing Then
                                ' フォルダ存在確認
                                If Not Directory.Exists(strWorkFolder) Then
                                    errors.Add($"Row {idx}：WORKフォルダが存在しません")
                                    Continue For
                                End If
                            Else
                                errors.Add($"Row {idx}：WORKフォルダパスが不正")
                                Continue For
                            End If

                            ' [WORKファイル名]を取得
                            'Dim strWorkFile As String = keys("StagedFileName").ToString()
                            strWorkFile = ""
                            strWorkFile = keys("StagedFileName").ToString()

                            '取込ファイル表示用に退避
                            TorikomiFile = ""
                            TorikomiFile = strWorkFile

                            If strWorkFile IsNot Nothing Then
                                'Dim workFile As String = workFolder & "/" & strWorkFile
                                strWorkFile = strWorkFolder & "\" & strWorkFile
                                ' ファイル存在確認
                                If Not File.Exists(strWorkFile) Then
                                    errors.Add($"Row {idx}：WORKファイルが存在しません")
                                    Continue For
                                End If
                            Else
                                errors.Add($"Row {idx}：WORKファイル名が不正")
                                Continue For
                            End If


                            'トランザクション制御
                            'If customerSettingId <> previousId Then
                            If impfilestageId <> preImpFileStageId Then
                                ' IDが変わった場合、前のトランザクションがあればコミットして終了
                                If tran IsNot Nothing Then
                                    If ErrFileFlg = True Then

                                        'tran.Rollback()
                                        'errors.Add($" 取引先コード：{ErrCustomerCode}　取込ファイル：[{ErrTorikomiFile} ]　はデータ不備のため取込実行から除外されました。")

                                        ReDrawFlg = True

                                        '取込不可の際はimp_files_stageテーブルのステータスをFAILEDに更新
                                        '_impFileStageRepo.UpdateImpFileStageStatus(impfilestageId)

                                        Try

                                            Dim MoveFileErrFlg As Boolean = False

                                            ' [フォルダパス]＋[ワークフォルダパス]＋[ワークファイル名]を取得
                                            Dim folderInfos As List(Of FolderPathInfo) = _impFileStageRepo.GetFolderInfosByImpFileStageId(preImpFileStageId)
                                            If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                                                errors.Add($"{customerCode}：IMP_FILES_STAGEにフォルダ未登録")
                                                MoveFileErrFlg = True
                                            End If

                                            'Dim foundInThisCustomer As Boolean = False

                                            Dim info = folderInfos(0)

                                            ' WORKフォルダ存在確認
                                            Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.Staged_FolderPath)
                                            If Not Directory.Exists(sourceFolder) Then
                                                errors.Add($"{customerCode}：WORKフォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                                                MoveFileErrFlg = True
                                            End If

                                            '取込元フォルダ存在確認
                                            Dim destFolder As String = Utils.ResolvePath(Me.Server, info.FolderPath)
                                            If Not Directory.Exists(destFolder) Then
                                                errors.Add($"{customerCode}：フォルダが存在しません [{Server.HtmlEncode(destFolder)}]")
                                                MoveFileErrFlg = True
                                            End If

                                            If MoveFileErrFlg = False Then

                                                Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                                                .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly))

                                                Dim fileName = info.Staged_FileName
                                                Dim destPath = Path.Combine(destFolder, fileName)
                                                Dim srcPath = Path.Combine(sourceFolder, fileName)

                                                ' ファイル名にログインユーザーIDとタイムスタンプを付ける
                                                Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                                Dim ext = Path.GetExtension(fileName)
                                                destPath = Path.Combine(destFolder, $"{nameNoExt}_{UserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}")

                                                Try

                                                    ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                                    'File.Move(srcPath, destPath)
                                                    File.Copy(srcPath, destPath)

                                                    '取込ファイルワークテーブルを削除する
                                                    '_impFileStageRepo.DeleteImpFileStageRange(tran, preImpFileStageId)

                                                    ''受注ワーク(取込(加工)済みデータ)削除
                                                    'cnt = _oderStageRepo.DeleteProcessedOrdersByFileId(tran, UserId, preImpFileStageId)

                                                    'resultRowCnt += cnt
                                                    'foundInThisCustomer = True

                                                Catch ex As UnauthorizedAccessException
                                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（アクセス権限不足：{ex.Message}）")
                                                Catch ex As IOException
                                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（I/O：{ex.Message}）")
                                                Catch ex As Exception
                                                    errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（{ex.Message}）")
                                                End Try

                                            End If

                                        Catch ex As Exception
                                            errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")

                                        End Try

                                        'Next
                                        'resultCnt += 1
                                        tran.Commit()
                                        resultAllCnt += resultCnt

                                    Else

                                        tran.Commit()
                                        resultAllCnt += resultCnt

                                    End If

                                    tran.Dispose()
                                End If

                                ' 新しいID用にトランザクションを開始
                                tran = conn.BeginTransaction()
                                'previousId = customerSettingId
                                preImpFileStageId = impfilestageId
                                ErrFileFlg = False
                                resultCnt = 0

                                ErrCustomerCode = customerCode
                                ErrTorikomiFile = TorikomiFile

                                'デバック用
                                '-----------------
                                '受注ワーク削除 ※customerSettingId単位で削除
                                '-----------------
                                '_oderStageRepo.DeleteRange(tran, customerSettingId)

                            End If

                            '-----------------
                            '受注ワーク削除 ※customerSettingId、folderType単位で削除
                            '-----------------
                            _oderStageRepo.DeleteRange(tran, customerSettingId, folderType)


                            Try

                                '※マッピングプロファイルマスタ(MAPPINNG_PROFILE_MST)と
                                '  マッピング明細マスタ(FIELD_MAPPINNG_MST)を参照してファイルデータを取得する
                                Dim mappingInfos As List(Of MappingInfo) = _mappingRepo.GetMappingInfo(customerSettingId, folderType)
                                If mappingInfos Is Nothing OrElse mappingInfos.Count = 0 Then
                                    errors.Add($"{mappingInfos}：MAPPINNG_PROFILE_MSTに未登録")
                                    Continue For
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
                                        'CSV(カンマ区切り) TSV(タブ区切り) 

                                        '許可する拡張子のリスト（小文字で定義）
                                        Dim allowedExtensions As New List(Of String) From {".csv", ".txt"}

                                        'ファイルパスから拡張子を取得
                                        Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                                        '拡張子チェック
                                        If Not allowedExtensions.Contains(fileExtension) Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                                            Continue For
                                        End If

                                        Dim config As New CsvConfiguration(CultureInfo.InvariantCulture)
                                        config.Delimiter = Delimiter            '区切り文字
                                        config.HasHeaderRecord = If(HeaderFlag = "Y", True, False)      ' 1行目がヘッダー（列名）の場合
                                        config.TrimOptions = TrimOptions.Trim ' 前後の余計な空白を自動で消す

                                        If Enclosure <> "" Then
                                            config.Quote = Enclosure               '囲い文字
                                        Else
                                            config.Mode = CsvMode.NoEscape
                                        End If

                                        Using StmRdr As New IO.StreamReader(strWorkFile, MapEncoding(CharSet))

                                            Using csv As New CsvReader(StmRdr, config)

                                                fileidx = 1

                                                'ヘッダーの存在チェック
                                                If HeaderFlag = "Y" Then
                                                    '/// 1行目のヘッダー行を飛ばす
                                                    csv.Read()
                                                    csv.ReadHeader()    'ヘッダーとして登録
                                                    fileidx += 1
                                                End If

                                                'データ開始行の手前までポインタを空打ちして進める
                                                While fileidx < DataStartRowIndex - 1
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
                                                    errMsg = ""

                                                    ErrFlg = False

                                                    'フォルダタイプで処理分岐
                                                    If folderType = 4 Then

                                                        '受注区分   (混在フォルダの場合は必須)
                                                        strQtyValue = If(csv.ColumnCount > nJutyuKubun AndAlso nJutyuKubun > -1, csv.GetField(nJutyuKubun).Trim(), "")
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
                                                        strQtyValue = If(csv.ColumnCount > nBunkatuKubun AndAlso nBunkatuKubun > -1, csv.GetField(nBunkatuKubun).Trim(), "")
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
                                                        ordertype = folderType

                                                        '分割区分   (任意)
                                                        'INFO_TYPE_MSTより取得
                                                        proratedtype = 1
                                                        errMsg = ""
                                                        If _oderStageRepo.GetProratedType(customerSettingId, folderType, proratedtype, errMsg) = False Then
                                                            'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                                            'ErrFlg = True
                                                        End If

                                                    End If

                                                    '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                                                    customerorderNo = If(csv.ColumnCount > nKyakusakiHattyuNo AndAlso nKyakusakiHattyuNo > -1, csv.GetField(nKyakusakiHattyuNo).Trim(), "")
                                                    If ordertype = 2 OrElse ordertype = 3 Then
                                                        'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                                        If String.IsNullOrEmpty(customerorderNo) Then
                                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                                            ErrFlg = True
                                                        End If
                                                    End If

                                                    '受注日   (任意)
                                                    strTempDate = If(csv.ColumnCount > nJutyuuBi AndAlso nJutyuuBi > -1, csv.GetField(nJutyuuBi).Trim(), "")
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
                                                    strTempDate = If(csv.ColumnCount > nKibouNouki AndAlso nKibouNouki > -1, csv.GetField(nKibouNouki).Trim(), "")
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
                                                    strQtyValue = If(csv.ColumnCount > njuyouSuu AndAlso njuyouSuu > -1, csv.GetField(njuyouSuu).Trim(), "")
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
                                                    selffcstflag = If(csv.ColumnCount > nJishaYosokuFlag AndAlso nJishaYosokuFlag > -1, csv.GetField(nJishaYosokuFlag).Trim(), "")

                                                    '自社予測削除フラグ   (自社予測フラグ = Yの時は必須)
                                                    selffcstdeleteflag = If(csv.ColumnCount > nJishaYosokuDelFlag AndAlso nJishaYosokuDelFlag > -1, csv.GetField(nJishaYosokuDelFlag).Trim(), "")
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
                                                    If nTukaCode > -1 Then
                                                        '取得ファイルに存在
                                                        currencycode = If(csv.ColumnCount > nTukaCode AndAlso nTukaCode > -1, csv.GetField(nTukaCode).Trim(), "")
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
                                                    If nKyakusakiHinmokuNo > -1 Then
                                                        customeritemNo = If(csv.ColumnCount > nKyakusakiHinmokuNo AndAlso nKyakusakiHinmokuNo > -1, csv.GetField(nKyakusakiHinmokuNo).Trim(), "")
                                                    End If

                                                    '製品コード  （任意）
                                                    If nSeihinCode > -1 Then
                                                        productcode = If(csv.ColumnCount > nSeihinCode AndAlso nSeihinCode > -1, csv.GetField(nSeihinCode).Trim(), "")
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
                                                    remarks = If(csv.ColumnCount > nComment AndAlso nComment > -1, csv.GetField(nComment).Trim(), "")

                                                    '納入先コード   （任意）
                                                    deliverycode = If(csv.ColumnCount > nNonyusakiCode AndAlso nNonyusakiCode > -1, csv.GetField(nNonyusakiCode).Trim(), "")

                                                    '出荷在庫場所   （任意）
                                                    'STRAMMIC.SECTMより取得
                                                    shipstocklocation = ""
                                                    errMsg = ""
                                                    'If _oderStageRepo.GetShipStockLocation(customerCode, deliverycode, shipstocklocation, errMsg) = False Then
                                                    If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                                        'ErrFlg = True
                                                    End If

                                                    '取引先情報区分   （任意）
                                                    customerinfotype = If(csv.ColumnCount > nTorihikisakiJohoKubun AndAlso nTorihikisakiJohoKubun > -1, csv.GetField(nTorihikisakiJohoKubun).Trim(), "")

                                                    '情報区分   （任意）
                                                    'INFO_TYPE_MSTより取得
                                                    infotype = ""
                                                    errMsg = ""
                                                    If _oderStageRepo.GetInfoType(customerSettingId, customerinfotype, infotype, errMsg) = False Then
                                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                                        'ErrFlg = True
                                                    End If

                                                    '消込条件区分   （任意）
                                                    'IMP_RULE_MSTより取得
                                                    reconciletype = 1
                                                    errMsg = ""
                                                    If _oderStageRepo.GetReconcileType(customerSettingId, folderType, reconciletype, errMsg) = False Then
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
                                                        .CustomerSettingId = customerSettingId,
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
                                                        .CreatedAt = now,
                                                        .CreatedUserId = UserId,
                                                        .CreatedPgId = pgId,
                                                        .UpdatedAt = now,
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
                                            Continue For
                                        End If


                                        Using StmRdr As New IO.StreamReader(strWorkFile, MapEncoding(CharSet))

                                        End Using


                                    Case "EXCEL"
                                        '4:EXCEL LIST

                                        '許可する拡張子のリスト（小文字で定義）
                                        Dim allowedExtensions As New List(Of String) From {".xlsx"}

                                        'ファイルパスから拡張子を取得
                                        Dim fileExtension As String = Path.GetExtension(strWorkFile).ToLower()

                                        '拡張子チェック
                                        If Not allowedExtensions.Contains(fileExtension) Then
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　許可されていないファイル形式です。")
                                            Continue For
                                        End If

                                        If FomatType = "LIST" Then

                                            'ワークブックを作成
                                            Using objWorkBook As New ClosedXML.Excel.XLWorkbook(strWorkFile)

                                                'ワークシートを作成
                                                Dim objSheet As ClosedXML.Excel.IXLWorksheet

                                                'ワークシート指定があれば指定
                                                If DefaultSheetName <> "" Then
                                                    objSheet = objWorkBook.Worksheet(DefaultSheetName)
                                                Else
                                                    objSheet = objWorkBook.Worksheet(1)
                                                End If

                                                'データの最終行を取得
                                                Dim lastRow = objSheet.LastRowUsed().RowNumber()

                                                fileidx = DataStartRowIndex

                                                'データ開始行からスタート
                                                For rowNum As Integer = DataStartRowIndex To lastRow

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

                                                    ErrFlg = False

                                                    'フォルダタイプで処理分岐
                                                    If folderType = 4 Then

                                                        '受注区分   (混在フォルダの場合は必須)
                                                        strQtyValue = If(nJutyuKubun > 0, xlRow.Cell(nJutyuKubun).GetValue(Of String)().Trim(), "")
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
                                                        strQtyValue = If(nBunkatuKubun > 0, xlRow.Cell(nBunkatuKubun).GetValue(Of String)().Trim(), "")
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
                                                        ordertype = folderType

                                                        '分割区分   (任意)
                                                        'INFO_TYPE_MSTより取得
                                                        proratedtype = 1
                                                        errMsg = ""
                                                        If _oderStageRepo.GetProratedType(customerSettingId, folderType, proratedtype, errMsg) = False Then
                                                            'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                                            'ErrFlg = True
                                                        End If

                                                    End If

                                                    '客先発注番号   (ordertype = 1:内示は任意、2:確定と3：納入指示は必須)
                                                    customerorderNo = If((nKyakusakiHattyuNo) > 0, xlRow.Cell(nKyakusakiHattyuNo).GetValue(Of String)().Trim(), "")
                                                    If ordertype = 2 OrElse ordertype = 3 Then
                                                        'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                                        If String.IsNullOrEmpty(customerorderNo) Then
                                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：客先発注番号が空です。")
                                                            ErrFlg = True
                                                        End If
                                                    End If

                                                    '受注日   (任意)
                                                    strTempDate = If(nJutyuuBi > 0, xlRow.Cell(nJutyuuBi).GetValue(Of String)().Trim(), "")
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
                                                    strTempDate = If(nKibouNouki > 0, xlRow.Cell(nKibouNouki).GetValue(Of String)().Trim(), "")
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
                                                    dueDate = cal.AddWorkingDays(conn, tran, "00001", tdt, 0)
                                                    '--


                                                    '需要数   (必須)
                                                    strQtyValue = If(njuyouSuu > 0, xlRow.Cell(njuyouSuu).GetValue(Of String)().Trim(), "")
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
                                                    selffcstflag = If(nJishaYosokuFlag > 0, xlRow.Cell(nJishaYosokuFlag).GetValue(Of String)().Trim(), "")

                                                    '自社予測削除フラグ  ※'自社予測フラグ = Yの時は必須
                                                    selffcstdeleteflag = If(nJishaYosokuDelFlag > 0, xlRow.Cell(nJishaYosokuDelFlag).GetValue(Of String)().Trim(), "")
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
                                                    If nTukaCode > 0 Then
                                                        '取得ファイルに存在
                                                        currencycode = If(nTukaCode > 0, xlRow.Cell(nTukaCode).GetValue(Of String)().Trim(), "")
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
                                                    If nKyakusakiHinmokuNo > 0 Then
                                                        customeritemNo = If(nKyakusakiHinmokuNo > 0, xlRow.Cell(nKyakusakiHinmokuNo).GetValue(Of String)().Trim(), "")
                                                    End If

                                                    '製品コード  （任意）
                                                    If nSeihinCode > 0 Then
                                                        productcode = If(nSeihinCode > 0, xlRow.Cell(nSeihinCode).GetValue(Of String)().Trim(), "")
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
                                                    remarks = If(nComment > 0, xlRow.Cell(nComment).GetValue(Of String)().Trim(), "")

                                                    '納入先コード （任意）
                                                    deliverycode = If(nNonyusakiCode > 0, xlRow.Cell(nNonyusakiCode).GetValue(Of String)().Trim(), "")

                                                    '出荷在庫場所 （任意）
                                                    'STRAMMIC.SECTMより取得
                                                    shipstocklocation = ""
                                                    errMsg = ""
                                                    'If _oderStageRepo.GetShipStockLocation(customerCode, deliverycode, shipstocklocation, errMsg) = False Then
                                                    If _oderStageRepo.GetShipStockLocation(productcode, shipstocklocation, errMsg) = False Then
                                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                                        'ErrFlg = True
                                                    End If

                                                    '取引先情報区分    （任意）
                                                    customerinfotype = If(nTorihikisakiJohoKubun > 0, xlRow.Cell(nTorihikisakiJohoKubun).GetValue(Of String)().Trim(), "")

                                                    '情報区分       （任意）
                                                    'INFO_TYPE_MSTより取得
                                                    infotype = ""
                                                    errMsg = ""
                                                    If _oderStageRepo.GetInfoType(customerSettingId, customerinfotype, infotype, errMsg) = False Then
                                                        'errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　Row {fileidx}：{errMsg}")
                                                        'ErrFlg = True
                                                    End If

                                                    '消込条件区分     （任意）
                                                    'IMP_RULE_MSTより取得
                                                    reconciletype = 1
                                                    errMsg = ""
                                                    If _oderStageRepo.GetReconcileType(customerSettingId, folderType, reconciletype, errMsg) = False Then
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
                                                        .CustomerSettingId = customerSettingId,
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
                                                        .CreatedAt = now,
                                                        .CreatedUserId = UserId,
                                                        .CreatedPgId = pgId,
                                                        .UpdatedAt = now,
                                                        .UpdatedUserId = UserId,
                                                        .UpdatedPgId = pgId
                                                    })

                                                    fileidx += 1

                                                Next

                                            End Using


                                        ElseIf FomatType = "MATRIX" Then

                                            Dim orderDateErrFlg As Boolean = False
                                            Dim HeaderErrFlg As Boolean = False

                                            '希望納期がマッピングマスタにない場合は、この後の処理が成立しないため処理を中断する。
                                            If mKibouNouki = "" Then
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　マッピングマスタに不備があります。(希望納期がマッピングマスタ未登録)")
                                                Continue For
                                            End If

                                            'ワークブックを作成
                                            Using objWorkBook As New ClosedXML.Excel.XLWorkbook(strWorkFile)

                                                'ワークシートを作成
                                                Dim objSheet As ClosedXML.Excel.IXLWorksheet

                                                'ワークシート指定があれば指定
                                                If DefaultSheetName <> "" Then
                                                    objSheet = objWorkBook.Worksheet(DefaultSheetName)
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
                                                strTempDate = If(mJutyuuBi <> "", objSheet.Cell(mJutyuuBi).GetValue(Of String)().Trim(), "")
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
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mJutyuuBi).Address.ColumnLetter & (objSheet.Cell(mJutyuuBi).Address.RowNumber)} ：受注日が不正な値です。")
                                                        ' ErrFlg = True
                                                        orderDateErrFlg = True
                                                    End If
                                                End If

                                                'データの最終行を取得
                                                Dim EdRowNum As Integer = objSheet.LastRowUsed().RowNumber()

                                                '納期の行番号をセットする 例:B4=4
                                                Dim StRowNum As Integer = objSheet.Cell(mKibouNouki).Address.RowNumber
                                                '納期の列番号をセットする 例:B4=2
                                                Dim StColNum As Integer = objSheet.Cell(mKibouNouki).Address.ColumnNumber

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

                                                    If orderDateErrFlg = True Then
                                                        HeaderErrFlg = True
                                                    Else
                                                        HeaderErrFlg = False
                                                    End If


                                                    '客先品目No   (任意だが、MATRIX形式は製品コードの項目がマトリックス共通表サンプルに存在しないので客先品目Noがないと品目Noが取得できない)
                                                    Dim cell1 As Integer = 0
                                                    'Dim cellname As String = ""
                                                    If Not String.IsNullOrEmpty(mKyakusakiHinmokuNo) Then
                                                        xlRow = objSheet.Row(objSheet.Cell(mKyakusakiHinmokuNo).Address.RowNumber + (stepRow * ridx))
                                                        cell1 = objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnNumber
                                                        'cellname = objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber
                                                        customeritemNo = If(mKyakusakiHinmokuNo <> "", xlRow.Cell(cell1).GetValue(Of String)().Trim(), "")
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
                                                        If mKyakusakiHinmokuNo = "" Then
                                                            errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　{errMsg} (マッピングマスタ未登録)")
                                                        Else
                                                            errors.Add($"取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：{errMsg}")
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
                                                    'STRAMMIC.SECTMより取得
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
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：客先品目Noが桁数超過のためトリミングされました。")
                                                    End If
                                                    '製品コード
                                                    productcode = SafeVarcharLength(productcode, 45, isTruncated)
                                                    If isTruncated = True Then
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：製品コードが桁数超過のためトリミングされました。")
                                                    End If
                                                    '品目No
                                                    itemNo = SafeVarcharLength(itemNo, 45, isTruncated)
                                                    If isTruncated = True Then
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：品目Noが桁数超過のためトリミングされました。")
                                                    End If
                                                    '需要単位
                                                    demandunit = SafeVarcharLength(demandunit, 4, isTruncated)
                                                    If isTruncated = True Then
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：需要単位が桁数超過のためトリミングされました。")
                                                    End If
                                                    '出荷在庫場所
                                                    shipstocklocation = SafeVarcharLength(shipstocklocation, 25, isTruncated)
                                                    If isTruncated = True Then
                                                        errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {objSheet.Cell(mKyakusakiHinmokuNo).Address.ColumnLetter & xlRow.RowNumber} ：出荷在庫場所が桁数超過のためトリミングされました。")
                                                    End If
                                                    '-----------------

                                                    '納期の列番号を始点として最終列までループ
                                                    For intColIdx As Integer = StColNum To EdColNum

                                                        If HeaderErrFlg = True Then
                                                            ErrFlg = True
                                                        Else
                                                            ErrFlg = False
                                                        End If

                                                        '受注日
                                                        '※受注日はヘッダー部で取得済み

                                                        '需要数   (必須)
                                                        If Not String.IsNullOrEmpty(mjuyouSuu) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mjuyouSuu).Address.RowNumber + (stepRow * ridx))
                                                            strQtyValue = If(mjuyouSuu <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                        Else
                                                            strQtyValue = ""
                                                        End If

                                                        If String.IsNullOrEmpty(strQtyValue) Then
                                                            'MATRIXでは、需要数が空だった場合は登録対象外とする
                                                            If mjuyouSuu = "" Then
                                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　需要数が取得できません。 (マッピングマスタ未登録)")
                                                            Else
                                                                'errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：需要数が不正な値です。")
                                                            End If
                                                            Continue For
                                                        End If
                                                        If Not Decimal.TryParse(strQtyValue, demandqty) Then
                                                            If mjuyouSuu = "" Then
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
                                                        If folderType = 4 Then

                                                            '受注区分   (混在フォルダの場合は必須)
                                                            If Not String.IsNullOrEmpty(mJutyuKubun) Then
                                                                xlRow = objSheet.Row(objSheet.Cell(mJutyuKubun).Address.RowNumber + (stepRow * ridx))
                                                                strQtyValue = If(mJutyuKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                            Else
                                                                strQtyValue = ""
                                                            End If
                                                            If String.IsNullOrEmpty(strQtyValue) Then
                                                                '必須チェック
                                                                If mJutyuKubun = "" Then
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
                                                            If Not String.IsNullOrEmpty(mBunkatuKubun) Then
                                                                xlRow = objSheet.Row(objSheet.Cell(mBunkatuKubun).Address.RowNumber + (stepRow * ridx))
                                                                strQtyValue = If(mBunkatuKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                            Else
                                                                strQtyValue = ""
                                                            End If
                                                            If String.IsNullOrEmpty(strQtyValue) Then
                                                                '必須チェック
                                                                If mBunkatuKubun = "" Then
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
                                                            ordertype = folderType

                                                            '分割区分   (任意)
                                                            'INFO_TYPE_MSTより取得
                                                            proratedtype = 1
                                                            errMsg = ""
                                                            If _oderStageRepo.GetProratedType(customerSettingId, folderType, proratedtype, errMsg) = False Then
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
                                                        If Not String.IsNullOrEmpty(mKyakusakiHattyuNo) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mKyakusakiHattyuNo).Address.RowNumber + (stepRow * ridx))
                                                            customerorderNo = If((mKyakusakiHattyuNo) <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                        Else
                                                            customerorderNo = ""
                                                        End If
                                                        'If folderType = 2 OrElse folderType = 3 Then
                                                        If ordertype = 2 OrElse ordertype = 3 Then
                                                            'ordertype = 1:内示は任意、2:確定と3：納入指示は必須
                                                            If String.IsNullOrEmpty(customerorderNo) Then
                                                                If mKyakusakiHattyuNo = "" Then
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
                                                        If Not String.IsNullOrEmpty(mKibouNouki) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mKibouNouki).Address.RowNumber)
                                                            strTempDate = If(mKibouNouki <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                        Else
                                                            strTempDate = ""
                                                        End If
                                                        If String.IsNullOrEmpty(strTempDate) Then
                                                            dueDate = CDate("1900/01/01")
                                                            If mKibouNouki = "" Then
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
                                                        If Not String.IsNullOrEmpty(mJishaYosokuFlag) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mJishaYosokuFlag).Address.RowNumber + (stepRow * ridx))
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
                                                        If Not String.IsNullOrEmpty(mJishaYosokuDelFlag) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mJishaYosokuDelFlag).Address.RowNumber + (stepRow * ridx))
                                                            selffcstdeleteflag = If(mJishaYosokuDelFlag <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
                                                        Else
                                                            selffcstdeleteflag = ""
                                                        End If
                                                        If selffcstflag = "Y" AndAlso String.IsNullOrEmpty(selffcstdeleteflag) Then
                                                            If mJishaYosokuDelFlag = "" Then
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
                                                        If mTukaCode <> "" Then
                                                            '取得ファイルに存在
                                                            If Not String.IsNullOrEmpty(mTukaCode) Then
                                                                xlRow = objSheet.Row(objSheet.Cell(mTukaCode).Address.RowNumber + (stepRow * ridx))
                                                                currencycode = If(mTukaCode <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
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
                                                            If mTukaCode <> "" Then
                                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　セル名 {xlRow.Cell(intColIdx).Address.ColumnLetter & xlRow.RowNumber} ：通貨コードが桁数超過のためトリミングされました。")
                                                            Else
                                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　通貨コードが桁数超過のためトリミングされました。")
                                                            End If
                                                        End If
                                                        '-----------------

                                                        'コメント   （任意）
                                                        If Not String.IsNullOrEmpty(mComment) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mComment).Address.RowNumber + (stepRow * ridx))
                                                            remarks = If(mComment <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
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
                                                        If Not String.IsNullOrEmpty(mNonyusakiCode) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mNonyusakiCode).Address.RowNumber + (stepRow * ridx))
                                                            deliverycode = If(mNonyusakiCode <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
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
                                                        If Not String.IsNullOrEmpty(mTorihikisakiJohoKubun) Then
                                                            xlRow = objSheet.Row(objSheet.Cell(mTorihikisakiJohoKubun).Address.RowNumber + (stepRow * ridx))
                                                            customerinfotype = If(mTorihikisakiJohoKubun <> "", xlRow.Cell(intColIdx).GetValue(Of String)().Trim(), "")
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
                                                        If _oderStageRepo.GetInfoType(customerSettingId, customerinfotype, infotype, errMsg) = False Then
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

                                                        '消込条件区分 （任意）
                                                        'IMP_RULE_MSTより取得
                                                        reconciletype = 1
                                                        errMsg = ""
                                                        If _oderStageRepo.GetReconcileType(customerSettingId, folderType, reconciletype, errMsg) = False Then
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
                                                            .CustomerSettingId = customerSettingId,
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
                                                            .CreatedAt = now,
                                                            .CreatedUserId = UserId,
                                                            .CreatedPgId = pgId,
                                                            .UpdatedAt = now,
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





                                '-----------------------------------------------
                                '取込ファイルの内容を受注ワーク登録
                                cnt = _oderStageRepo.InsertRange(tran, rowsForTemp2)
                                '-----------------------------------------------




                                If cnt > 0 Then


                                    '-----------------------------------------------
                                    'IMP_FILE_STAGEを更新
                                    Dim strHandFlag As String = ""
                                    If chkHandFlag.Checked = True Then
                                        strHandFlag = "Y"
                                    Else
                                        strHandFlag = "N"
                                    End If

                                    Dim rowsForTemp3 = New List(Of ImpFilesStageRow) From {
                                                        New ImpFilesStageRow With {
                                                            .ImpFileStageId = impfilestageId,
                                                            .HandFlag = strHandFlag,
                                                            .Status = "PARSED",
                                                            .UpdatedAt = now,
                                                            .UpdatedUserId = UserId,
                                                            .UpdatedPgId = pgId
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
                                            .CustomerSettingId = customerSettingId
                                        }
                                    }
                                    _oderStageRepo.InsertStageFromOrders(tran, customerSettingId)
                                    '-----------------------------------------------



                                    '-----------------------------------------------
                                    '--------
                                    '内示加工
                                    '--------
                                    '今回取込した内示データの抽出
                                    Dim dtNaiji As DataTable = _oderStageRepo.GetNaijiData(tran, impfilestageId)

                                    '今回取込した内示データの件数をチェック
                                    If dtNaiji.Rows.Count > 0 Then

                                        '内示洗い替え
                                        _oderStageRepo.ReplaceNaijiRelation(tran, impfilestageId, customerSettingId, now, UserId, pgId)

                                        'ステータス更新
                                        _oderStageRepo.UpdateNaijiStatusProcessed(tran, impfilestageId)



                                        '2026/05/26 酒井 フェーズ2 受注残対応
                                        '受注残加工
                                        '消込フラグをチェック
                                        If reconcileFlag = "Y" Then

                                            '--------------------------------
                                            '受注残消込
                                            '--------------------------------
                                            _oderStageRepo.BacklogForecast(tran,
                                                                            customerSettingId,
                                                                            impfilestageId,
                                                                            reconciletype,
                                                                            now,
                                                                            UserId,
                                                                            pgId)

                                        End If
                                        '--



                                    End If
                                    '-----------------------------------------------






                                    '-----------------------------------------------
                                    '--------
                                    '確定加工
                                    '--------

                                    '打切処理
                                    _oderStageRepo.UpdateClese(tran, impfilestageId, 2)

                                    '取消処理
                                    _oderStageRepo.UpdateCancel(tran, impfilestageId, 2)


                                    '確定データ無効化
                                    _oderStageRepo.ReplaceKakuteiRelation(tran,
                                                                            customerSettingId,
                                                                            impfilestageId,
                                                                            now,
                                                                            UserId,
                                                                            pgId)

                                    'デバック用
                                    'tran.Commit()
                                    'Exit Sub

                                    '消込フラグをチェック
                                    If reconcileFlag = "Y" Then

                                        '確定の抽出
                                        Dim dtKakutei As DataTable = _oderStageRepo.GetKakuteiData(tran, impfilestageId)

                                        If dtKakutei.Rows.Count > 0 Then

                                            '--------------------------------
                                            '確定データで内示消込
                                            '--------------------------------
                                            _oderStageRepo.ReconcileForecast(tran,
                                                                                customerSettingId,
                                                                                impfilestageId,
                                                                                2,
                                                                                reconciletype,
                                                                                now,
                                                                                UserId,
                                                                                pgId)

                                        End If

                                    End If

                                    '--------------------------------
                                    '確定 新規
                                    '--------------------------------
                                    _oderStageRepo.UpdateKakuteiNewOrders(tran, impfilestageId)

                                    '-----------------------------------------------





                                    '-----------------------------------------------
                                    '--------
                                    '納入指示加工
                                    '--------

                                    '打切処理
                                    _oderStageRepo.UpdateClese(tran, impfilestageId, 3)

                                    '取消処理
                                    _oderStageRepo.UpdateCancel(tran, impfilestageId, 3)

                                    '消込フラグをチェック
                                    If reconcileFlag = "Y" Then


                                        '--------------------------------
                                        '確定消込 ※客先発注Noでの消込
                                        '--------------------------------
                                        '受注消込(客先発注No)
                                        _oderStageRepo.ExecuteOrderReconciliationByOrderNo(tran,
                                                                                            customerSettingId,
                                                                                            impfilestageId,
                                                                                            now,
                                                                                            UserId,
                                                                                            pgId)

                                        '内示消込フラグをチェック
                                        If fcstreconcileFlag = "Y" Then

                                            '納入指示注文の抽出
                                            Dim dtNonyuSiji As DataTable = _oderStageRepo.GetNonyuSijiData(tran, impfilestageId)

                                            If dtNonyuSiji.Rows.Count > 0 Then

                                                ''デバック用
                                                'tran.Commit()
                                                'Exit Sub

                                                '--------------------------------
                                                '納入指示データで内示消込
                                                '--------------------------------
                                                _oderStageRepo.ReconcileForecast(tran,
                                                                                    customerSettingId,
                                                                                    impfilestageId,
                                                                                    3,
                                                                                    reconciletype,
                                                                                    now,
                                                                                    UserId,
                                                                                    pgId)

                                            End If

                                        End If

                                    End If

                                    '--------------------------------
                                    '納入指示 新規
                                    '--------------------------------
                                    _oderStageRepo.UpdateNonyuSijiNewOrders(tran, impfilestageId)

                                    '-----------------------------------------------


                                    If ErrFlg = False Then

                                        resultCnt += 1
                                        resultRowCnt += cnt

                                        successs.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　読込 {cnt} 件　異常 {errcnt} 件")
                                        'successs.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　読込 {cnt} 件")

                                        'errcnt = 0

                                    End If


                                End If




                            Catch ex As Exception
                                Dim exmes As String = ex.Message
                                ' システムエラー発生時もフラグを立てる
                                'errors.Add($"Row {row.RowIndex + 1}：システムエラー({ex.Message})")
                                'errors.Add($"Row {idx}：エラーが発生したため取込実行されません。")
                                errors.Add($"システムエラー({ex.Message})")
                                'If errors.Count = 0 Then
                                '    lblImportResult.Text = "取込実行：予期せぬエラーが発生しました。"
                                'End If
                                ErrFlg = True
                            End Try

                        Next

                        ' ループ終了後、最後のグループをコミット
                        If tran IsNot Nothing Then
                            If ErrFileFlg = True Then

                                'tran.Rollback()
                                'errors.Add($"取引先コード：{ErrCustomerCode}　取込ファイル：[{ErrTorikomiFile} ]　はデータ不備のため取込実行から除外されました。")

                                ReDrawFlg = True

                                '取込不可の際はimp_files_stageテーブルのステータスをFAILEDに更新
                                '_impFileStageRepo.UpdateImpFileStageStatus(impfilestageId)


                                Try

                                    Dim MoveFileErrFlg As Boolean = False

                                    ' [フォルダパス]＋[ワークフォルダパス]＋[ワークファイル名]を取得
                                    Dim folderInfos As List(Of FolderPathInfo) = _impFileStageRepo.GetFolderInfosByImpFileStageId(impfilestageId)
                                    If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                                        errors.Add($"{customerCode}：IMP_FILES_STAGEにフォルダ未登録")
                                        MoveFileErrFlg = True
                                    End If

                                    'Dim foundInThisCustomer As Boolean = False

                                    Dim info = folderInfos(0)

                                    ' WORKフォルダ存在確認
                                    Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.Staged_FolderPath)
                                    If Not Directory.Exists(sourceFolder) Then
                                        errors.Add($"{customerCode}：WORKフォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                                        MoveFileErrFlg = True
                                    End If

                                    '取込元フォルダ存在確認
                                    Dim destFolder As String = Utils.ResolvePath(Me.Server, info.FolderPath)
                                    If Not Directory.Exists(destFolder) Then
                                        errors.Add($"{customerCode}：フォルダが存在しません [{Server.HtmlEncode(destFolder)}]")
                                        MoveFileErrFlg = True
                                    End If

                                    If MoveFileErrFlg = False Then

                                        Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                                        .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly))

                                        Dim fileName = info.Staged_FileName
                                        Dim destPath = Path.Combine(destFolder, fileName)
                                        Dim srcPath = Path.Combine(sourceFolder, fileName)

                                        ' ファイル名にログインユーザーIDとタイムスタンプを付ける
                                        Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                        Dim ext = Path.GetExtension(fileName)
                                        destPath = Path.Combine(destFolder, $"{nameNoExt}_{UserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}")

                                        Try

                                            ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                            'File.Move(srcPath, destPath)
                                            File.Copy(srcPath, destPath)

                                            '取込ファイルワークテーブルを削除する
                                            '_impFileStageRepo.DeleteImpFileStageRange(tran, impfilestageId)

                                            ''受注ワーク(取込(加工)済みデータ)削除
                                            'cnt = _oderStageRepo.DeleteProcessedOrdersByFileId(tran, UserId, impfilestageId)

                                            'resultRowCnt += cnt
                                            'foundInThisCustomer = True

                                        Catch ex As UnauthorizedAccessException
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（アクセス権限不足：{ex.Message}）")
                                        Catch ex As IOException
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（I/O：{ex.Message}）")
                                        Catch ex As Exception
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（{ex.Message}）")
                                        End Try

                                    End If

                                Catch ex As Exception
                                    errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")

                                End Try

                                'Next
                                'resultCnt += 1
                                tran.Commit()
                                resultAllCnt += resultCnt

                            Else

                                tran.Commit()
                                resultAllCnt += resultCnt
                                'resultRowCnt += cnt

                            End If
                        End If

                    Catch ex As Exception

                        Dim exmes As String = ex.Message
                        ' ロールバック処理
                        If tran IsNot Nothing Then
                            tran.Rollback()
                        End If

                        'errors.Add($"エラーが発生したため取込実行されません。")
                        errors.Add($"システムエラー({ex.Message})")

                    Finally

                        If tran IsNot Nothing Then
                            tran.Dispose()
                        End If

                    End Try

                End Using

                ' 結果表示
                'If ok AndAlso results.Count > 0 Then
                '    lblImportResult.Text = $"取込実行：対象ファイル {results.Count} 件"
                'Else
                '    If errors.Count = 0 Then
                '        lblImportResult.Text = "取込実行：予期せぬエラーが発生しました。"
                '    End If
                'End If
                'If resultCnt > 0 Then
                If resultAllCnt > 0 Then
                    'lblImportResult.Text &= $"（ORDERS_STAGE 読込 {resultAllCnt} 件）"


                    'lblImportResult.Text &= $"（ORDERS_STAGE 読込 {resultRowCnt} 件）"


                    Dim alreadyText = $"（ORDERS_STAGE 読込 {resultRowCnt} 件）"
                    Dim addText = String.Join("<br/>", successs.Select(Function(s) Server.HtmlEncode(s)))
                    If String.IsNullOrEmpty(alreadyText) Then
                        lblImportResult.Text = addText
                    Else
                        lblImportResult.Text = alreadyText & "<br/>" & addText
                    End If


                Else
                    If errors.Count = 0 Then
                        lblImportResult.Text = "取込実行：予期せぬエラーが発生しました。"
                    End If
                End If

            End If

            If errors.Count > 0 Then
                'Dim alreadyText = lblImportError.Text
                Dim alreadyText = "（ORDERS_STAGE 読込時にエラー）"
                Dim addText = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
                If String.IsNullOrEmpty(alreadyText) Then
                    lblImportError.Text = addText
                Else
                    lblImportError.Text = alreadyText & "<br/>" & addText
                End If
            End If

            If ReDrawFlg = True Then
                '取込ファイル一覧　再描画
                gvImpFilesStage_Init()
            End If
            '取込済み受注一覧　再描画
            gvImportOrder_Init()

        End Sub

        ' 破棄ボタン(取込前)　　(上部グリッドの処理対象チェックボックス単位で破棄を実行、取込ファイルを取込前に戻す処理(IMP_FILES_STAGEテーブル削除)と、該当するORDER_STAGEテーブル削除)
        Protected Sub btnImportCancel_Click(sender As Object, e As EventArgs)

            lblImportResult.Text = ""
            lblImportError.Text = ""
            lblSaveResult.Text = ""
            lblSaveError.Text = ""

            Dim ErrFlg As Boolean = False
            Dim resultCnt As Integer = 0
            'Dim resultRowCnt As Integer = 0
            Dim errors As New List(Of String)()

            ' DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            Dim UserId As String = PageHelpers.GetUserId(Me)
            If String.IsNullOrWhiteSpace(UserId) Then
                UserId = "AMAGATA"
            End If
            If UserId.Length > 9 Then
                UserId = UserId.Substring(0, 9)
            End If

            Using conn As New OracleConnection(connStr)
                conn.Open()

                ''加工済み受注　取得
                'Dim dtCustomerSettingIds As DataTable = _oderStageRepo.GetProcessedCustomerSettingIds(UserId)

                Dim previousId As Long = -1 ' 前回の取引先設定ID保持用
                Dim customerSettingId As Long = 0
                Dim impFileStageId As Long = 0
                Dim customerCode As Long = 0
                Dim cnt As Long = 0

                '---------------------------------------------------------------
                'グリッド内のチェック状態チェック
                '---------------------------------------------------------------
                Dim groups As New Dictionary(Of String, List(Of GridViewRow))
                Dim selectedRows As New List(Of GridViewRow)

                For Each row As GridViewRow In gvImpFilesStage.Rows
                    If row.RowType <> DataControlRowType.DataRow Then Continue For

                    Dim chk As CheckBox = TryCast(row.FindControl("chkOrderImport"), CheckBox)
                    If chk IsNot Nothing AndAlso chk.Checked Then
                        selectedRows.Add(row)
                        Dim settingId As String = gvImpFilesStage.DataKeys(row.RowIndex)("CustomerSettingId").ToString()

                        If Not groups.ContainsKey(settingId) Then
                            groups(settingId) = New List(Of GridViewRow)()
                        End If
                        groups(settingId).Add(row)
                    End If
                Next

                If selectedRows.Count = 0 Then
                    errors.Add("処理対象を選択してください。")
                End If

                'If dtCustomerSettingIds.Rows.Count > 0 Then
                If errors.Count = 0 Then

                    '取引先設定ID単位で処理をループ
                    'For intRowIndex As Integer = 0 To dtCustomerSettingIds.Rows.Count - 1
                    'Dim dr As DataRow = dtCustomerSettingIds.Rows(intRowIndex)
                    For Each row As GridViewRow In gvImpFilesStage.Rows

                        ' データ行以外はスキップ
                        If row.RowType <> DataControlRowType.DataRow Then
                            Continue For
                        End If

                        '処理対象 未選択行はスキップ
                        Dim chk As CheckBox = TryCast(row.FindControl("chkOrderImport"), CheckBox)
                        If chk Is Nothing OrElse Not chk.Checked Then
                            Continue For
                        End If

                        Using tran As OracleTransaction = conn.BeginTransaction()

                            Try

                                Dim idx As Integer = row.RowIndex
                                Dim keys = gvImpFilesStage.DataKeys(idx)
                                If keys Is Nothing Then
                                    errors.Add($"Row {idx}：DataKeys未設定")
                                    Continue For
                                End If

                                ErrFlg = False

                                'customerSettingId = If(dr.IsNull("customer_setting_id"), 0, dr.Item("customer_setting_id").ToString())
                                'customerCode = If(dr.IsNull("customer_code"), 0, dr.Item("customer_code").ToString())

                                '取引先設定ID　GridViewから取得
                                customerSettingId = 0
                                Dim csidObj = keys("CustomerSettingId")
                                If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), customerSettingId) Then
                                    errors.Add($"Row {idx}：CustomerSettingIdが不正")
                                    Continue For
                                End If

                                '取引先コード　GridViewから取得
                                customerCode = 0
                                csidObj = keys("CustomerCode")
                                If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), customerCode) Then
                                    errors.Add($"Row {idx}：CustomerCodeが不正")
                                    Continue For
                                End If

                                'cnt = If(dr.IsNull("cnt"), 0, dr.Item("cnt").ToString())

                                'Dim dtImpFileStageIds As DataTable = _oderStageRepo.GetProcessedImpFileStageIds(UserId, customerSettingId)

                                '取込ファイルの数だけループ
                                'For intFileIndex As Integer = 0 To dtImpFileStageIds.Rows.Count - 1
                                '    Dim dr2 As DataRow = dtImpFileStageIds.Rows(intFileIndex)

                                'impFileStageId = If(dr2.IsNull("imp_file_stage_id"), 0, dr2.Item("imp_file_stage_id").ToString())

                                '一時取込ファイルID　GridViewから取得
                                'Dim impfilestageId As Long
                                impFileStageId = 0
                                csidObj = keys("ImpFileStageId")
                                If csidObj Is Nothing OrElse Not Long.TryParse(csidObj.ToString(), impFileStageId) Then
                                    errors.Add($"Row {idx}：ImpFileStageIdが不正")
                                    Continue For
                                End If

                                Try
                                    ' [フォルダパス]＋[ワークフォルダパス]＋[ワークファイル名]を取得
                                    Dim folderInfos As List(Of FolderPathInfo) = _impFileStageRepo.GetFolderInfosByImpFileStageId(impFileStageId)
                                    If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                                        'errors.Add($"{customerCode}：IMP_FILES_STAGEにフォルダ未登録")
                                        'ErrFlg = True
                                        'Continue For
                                    Else

                                        'Dim foundInThisCustomer As Boolean = False

                                        Dim info = folderInfos(0)

                                        ' WORKフォルダ存在確認
                                        Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.Staged_FolderPath)
                                        If Not Directory.Exists(sourceFolder) Then
                                            errors.Add($"{customerCode}：WORKフォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                                            ErrFlg = True
                                            Continue For
                                        End If

                                        '取込元フォルダ存在確認
                                        Dim destFolder As String = Utils.ResolvePath(Me.Server, info.FolderPath)
                                        If Not Directory.Exists(destFolder) Then
                                            errors.Add($"{customerCode}：フォルダが存在しません [{Server.HtmlEncode(destFolder)}]")
                                            ErrFlg = True
                                            Continue For
                                        End If

                                        Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                                            .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly))

                                        Dim fileName = info.Staged_FileName
                                        Dim destPath = Path.Combine(destFolder, fileName)
                                        Dim srcPath = Path.Combine(sourceFolder, fileName)

                                        ' ファイル名にログインユーザーIDとタイムスタンプを付ける
                                        Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                        Dim ext = Path.GetExtension(fileName)
                                        destPath = Path.Combine(destFolder, $"{nameNoExt}_{UserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}")

                                        Try

                                            ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                            File.Move(srcPath, destPath)

                                            '取込ファイルワークテーブルを削除する
                                            _impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                            resultCnt += 1

                                            'foundInThisCustomer = True

                                        Catch ex As UnauthorizedAccessException
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（アクセス権限不足：{ex.Message}）")
                                            ErrFlg = True
                                        Catch ex As IOException
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（I/O：{ex.Message}）")
                                            ErrFlg = True
                                        Catch ex As Exception
                                            errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（{ex.Message}）")
                                            ErrFlg = True
                                        End Try

                                    End If

                                Catch ex As Exception
                                    errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")
                                    ErrFlg = True
                                End Try

                                ''取込ファイルワークテーブルを削除する
                                '_impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                'Next

                                'resultCnt += 1


                                If ErrFlg = False Then

                                    '受注ワーク(取込(加工)済みデータ)削除　※取込ファイルID単位で削除
                                    'cnt = _oderStageRepo.DeleteProcessedOrdersRange(tran, UserId, customerSettingId)
                                    cnt = _oderStageRepo.DeleteProcessedOrdersByFileId(tran, UserId, impFileStageId)

                                    'resultCnt += 1
                                    'resultRowCnt += cnt
                                End If


                                If tran IsNot Nothing Then
                                    If ErrFlg = True Then
                                        'ロールバック
                                        tran.Rollback()
                                    Else
                                        'コミット
                                        tran.Commit()
                                        'tran.Rollback()
                                    End If
                                End If

                            Catch ex As Exception

                                Dim exmes As String = ex.Message
                                ' ロールバック処理
                                If tran IsNot Nothing Then
                                    tran.Rollback()
                                End If

                                'errors.Add($"エラーが発生したため破棄実行されません。")
                                errors.Add($"システムエラー({ex.Message})")

                            Finally

                                If tran IsNot Nothing Then
                                    tran.Dispose()
                                End If

                            End Try

                        End Using

                    Next

                End If

            End Using

            If resultCnt > 0 Then
                'lblSaveResult.Text = "破棄：対象データを破棄しました。"
                lblImportResult.Text &= $"（IMP_FILES_STAGE 破棄 {resultCnt} 件）"

                'グリッド再描画
                gvImpFilesStage_Init()
                gvImportOrder_Init()

            Else
                If errors.Count = 0 Then
                    'lblSaveResult.Text = "破棄：対象データが見つかりませんでした。"
                    lblImportResult.Text = "破棄：対象データが見つかりませんでした。"
                End If
            End If

            If errors.Count > 0 Then
                'Dim alreadyText = lblImportError.Text
                Dim alreadyText = "（IMP_FILES_STAGE 破棄時にエラー）"
                Dim addText = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
                If String.IsNullOrEmpty(alreadyText) Then
                    'lblSaveError.Text = addText
                    lblImportError.Text = addText
                Else
                    'lblSaveError.Text = alreadyText & "<br/>" & addText
                    lblImportError.Text = alreadyText & "<br/>" & addText
                End If
            End If

            'グリッド再描画
            'gvImpFilesStage_Init()
            'gvImportOrder_Init()

            'lblImportResult.Text = ""
            'lblImportError.Text = ""

        End Sub

        ' 破棄ボタン
        Protected Sub btnCancelOrder_Click(sender As Object, e As EventArgs)

            lblSaveResult.Text = ""
            lblSaveError.Text = ""

            Dim ErrFlg As Boolean = False
            Dim resultCnt As Integer = 0
            Dim resultRowCnt As Integer = 0
            Dim errors As New List(Of String)()

            ' DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            Dim UserId As String = PageHelpers.GetUserId(Me)
            If String.IsNullOrWhiteSpace(UserId) Then
                UserId = "AMAGATA"
            End If
            If UserId.Length > 9 Then
                UserId = UserId.Substring(0, 9)
            End If

            Using conn As New OracleConnection(connStr)
                conn.Open()

                '加工済み受注　取得
                Dim dtCustomerSettingIds As DataTable = _oderStageRepo.GetProcessedCustomerSettingIds(UserId)

                Dim previousId As Long = -1 ' 前回の取引先設定ID保持用
                Dim customerSettingId As Long = 0
                Dim impFileStageId As Long = 0
                Dim customerCode As Long = 0
                Dim cnt As Long = 0

                If dtCustomerSettingIds.Rows.Count > 0 Then

                    '取引先設定ID単位で処理をループ
                    For intRowIndex As Integer = 0 To dtCustomerSettingIds.Rows.Count - 1
                        Dim dr As DataRow = dtCustomerSettingIds.Rows(intRowIndex)

                        Using tran As OracleTransaction = conn.BeginTransaction()

                            Try

                                ErrFlg = False

                                customerSettingId = If(dr.IsNull("customer_setting_id"), 0, dr.Item("customer_setting_id").ToString())
                                customerCode = If(dr.IsNull("customer_code"), 0, dr.Item("customer_code").ToString())

                                'cnt = If(dr.IsNull("cnt"), 0, dr.Item("cnt").ToString())

                                Dim dtImpFileStageIds As DataTable = _oderStageRepo.GetProcessedImpFileStageIds(UserId, customerSettingId)

                                '取込ファイルの数だけループ
                                For intFileIndex As Integer = 0 To dtImpFileStageIds.Rows.Count - 1
                                    Dim dr2 As DataRow = dtImpFileStageIds.Rows(intFileIndex)

                                    impFileStageId = If(dr2.IsNull("imp_file_stage_id"), 0, dr2.Item("imp_file_stage_id").ToString())

                                    Try
                                        ' [フォルダパス]＋[ワークフォルダパス]＋[ワークファイル名]を取得
                                        Dim folderInfos As List(Of FolderPathInfo) = _impFileStageRepo.GetFolderInfosByImpFileStageId(impFileStageId)
                                        If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                                            'errors.Add($"{customerCode}：IMP_FILES_STAGEにフォルダ未登録")
                                            'ErrFlg = True
                                            'Continue For
                                        Else

                                            'Dim foundInThisCustomer As Boolean = False

                                            Dim info = folderInfos(0)

                                            ' WORKフォルダ存在確認
                                            Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.Staged_FolderPath)
                                            If Not Directory.Exists(sourceFolder) Then
                                                errors.Add($"{customerCode}：WORKフォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                                                ErrFlg = True
                                                Continue For
                                            End If

                                            '取込元フォルダ存在確認
                                            Dim destFolder As String = Utils.ResolvePath(Me.Server, info.FolderPath)
                                            If Not Directory.Exists(destFolder) Then
                                                errors.Add($"{customerCode}：フォルダが存在しません [{Server.HtmlEncode(destFolder)}]")
                                                ErrFlg = True
                                                Continue For
                                            End If

                                            Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                                                .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly))

                                            Dim fileName = info.Staged_FileName
                                            Dim destPath = Path.Combine(destFolder, fileName)
                                            Dim srcPath = Path.Combine(sourceFolder, fileName)

                                            ' ファイル名にログインユーザーIDとタイムスタンプを付ける
                                            Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                            Dim ext = Path.GetExtension(fileName)
                                            destPath = Path.Combine(destFolder, $"{nameNoExt}_{UserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}")

                                            Try

                                                ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                                File.Move(srcPath, destPath)

                                                '取込ファイルワークテーブルを削除する
                                                _impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                                'foundInThisCustomer = True

                                            Catch ex As UnauthorizedAccessException
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（アクセス権限不足：{ex.Message}）")
                                                ErrFlg = True
                                            Catch ex As IOException
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（I/O：{ex.Message}）")
                                                ErrFlg = True
                                            Catch ex As Exception
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（{ex.Message}）")
                                                ErrFlg = True
                                            End Try

                                            ''取込ファイルワークテーブルを削除する
                                            '_impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                        End If

                                    Catch ex As Exception
                                        errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")
                                        ErrFlg = True
                                    End Try

                                Next

                                If ErrFlg = False Then

                                    '受注ワーク(取込(加工)済みデータ)削除
                                    cnt = _oderStageRepo.DeleteProcessedOrdersRange(tran, UserId, customerSettingId)

                                    resultCnt += 1
                                    resultRowCnt += cnt
                                End If

                                If tran IsNot Nothing Then
                                    If ErrFlg = True Then
                                        'ロールバック
                                        tran.Rollback()
                                    Else
                                        'コミット
                                        tran.Commit()
                                        'tran.Rollback()
                                    End If
                                End If

                            Catch ex As Exception

                                Dim exmes As String = ex.Message
                                ' ロールバック処理
                                If tran IsNot Nothing Then
                                    tran.Rollback()
                                End If

                                'errors.Add($"エラーが発生したため破棄実行されません。")
                                errors.Add($"システムエラー({ex.Message})")

                            Finally

                                If tran IsNot Nothing Then
                                    tran.Dispose()
                                End If

                            End Try

                        End Using

                    Next

                End If

            End Using

            If resultCnt > 0 Then
                'lblSaveResult.Text = "破棄：対象データを破棄しました。"
                lblSaveResult.Text &= $"（ORDERS_STAGE 破棄 {resultRowCnt} 件）"

            Else
                If errors.Count = 0 Then
                    lblSaveResult.Text = "破棄：対象データが見つかりませんでした。"
                End If
            End If

            If errors.Count > 0 Then
                'Dim alreadyText = lblImportError.Text
                Dim alreadyText = "（ORDERS_STAGE 破棄時にエラー）"
                Dim addText = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
                If String.IsNullOrEmpty(alreadyText) Then
                    lblSaveError.Text = addText
                Else
                    lblSaveError.Text = alreadyText & "<br/>" & addText
                End If
            End If

            'グリッド再描画
            gvImpFilesStage_Init()
            gvImportOrder_Init()

            lblImportResult.Text = ""
            lblImportError.Text = ""

        End Sub

        ' 保存ボタン
        Protected Sub btnSaveOrder_Click(sender As Object, e As EventArgs)

            lblSaveResult.Text = ""
            lblSaveError.Text = ""

            Dim ErrFlg As Boolean = False
            Dim resultCnt As Integer = 0
            Dim resultRowCnt As Integer = 0
            Dim errors As New List(Of String)()

            'DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            Dim UserId As String = PageHelpers.GetUserId(Me)
            If String.IsNullOrWhiteSpace(UserId) Then
                UserId = "AMAGATA"
            End If
            If UserId.Length > 9 Then
                UserId = UserId.Substring(0, 9)
            End If

            'IMP_RUN更新用変数
            Dim runid As Long = 0
            Dim now As DateTime = DateTime.Now
            Dim pgId As String = "OrderImport(Save)"

            Using conn As New OracleConnection(connStr)
                conn.Open()

                '加工済み受注　取得
                Dim dtCustomerSettingIds As DataTable = _oderStageRepo.GetProcessedCustomerSettingIds(UserId)

                Dim previousId As Long = -1 ' 前回の取引先設定ID保持用
                Dim customerSettingId As Long = 0
                Dim impFileStageId As Long = 0
                Dim customerCode As Long = 0
                Dim cnt As Long = 0

                If dtCustomerSettingIds.Rows.Count > 0 Then

                    '取引先設定ID単位で処理をループ
                    For intRowIndex As Integer = 0 To dtCustomerSettingIds.Rows.Count - 1
                        Dim dr As DataRow = dtCustomerSettingIds.Rows(intRowIndex)

                        Using tran As OracleTransaction = conn.BeginTransaction()

                            Try

                                ErrFlg = False

                                customerSettingId = If(dr.IsNull("customer_setting_id"), 0, dr.Item("customer_setting_id").ToString())
                                customerCode = If(dr.IsNull("customer_code"), 0, dr.Item("customer_code").ToString())

                                'cnt = If(dr.IsNull("cnt"), 0, dr.Item("cnt").ToString())

                                '正規データ更新
                                _oderStageRepo.UpdateOrdersFromStage(tran, customerSettingId, now, UserId, pgId)

                                '正規データ追加
                                cnt = _oderStageRepo.InsertOrdersFromStage(tran, customerSettingId, now, UserId, pgId)

                                '受注履歴データ追加
                                _oderStageRepo.InsertHistoryFromOrders(tran, customerSettingId, now, UserId, pgId)


                                Dim dtImpFileStageIds As DataTable = _oderStageRepo.GetProcessedImpFileStageIds(UserId, customerSettingId)

                                '取込ファイルの数だけループ
                                For intFileIndex As Integer = 0 To dtImpFileStageIds.Rows.Count - 1
                                    Dim dr2 As DataRow = dtImpFileStageIds.Rows(intFileIndex)

                                    impFileStageId = If(dr2.IsNull("imp_file_stage_id"), 0, dr2.Item("imp_file_stage_id").ToString())

                                    Try
                                        '[ワークフォルダフォルダパス]＋[ワークファイル名]＋[受注区分]を取得
                                        Dim folderInfos As List(Of FolderPathInfo) = _impFileStageRepo.GetStageFolderInfosByImpFileStageId(impFileStageId)
                                        If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                                            'errors.Add($"{customerCode}：IMP_FILES_STAGEにフォルダ未登録")
                                            'ErrFlg = True
                                            'Continue For

                                        Else

                                            'Dim foundInThisCustomer As Boolean = False

                                            Dim info = folderInfos(0)

                                            Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.Staged_FolderPath)

                                            ' フォルダ存在確認
                                            If Not Directory.Exists(sourceFolder) Then
                                                errors.Add($"{customerCode}：WORKフォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                                                ErrFlg = True
                                                Continue For
                                            End If

                                            ' COMPLETEDサブフォルダ作成
                                            Dim destFolder As String = Path.Combine(_compUserRoot, customerCode, info.FolderType.ToString())
                                            Utils.EnsureDirectory(destFolder)

                                            Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                                            .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly))

                                            Dim fileName = info.Staged_FileName
                                            Dim destPath = Path.Combine(destFolder, fileName)
                                            Dim srcPath = Path.Combine(sourceFolder, fileName)

                                            ' ファイル名にログインユーザーIDとタイムスタンプを付ける
                                            Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                            Dim ext = Path.GetExtension(fileName)
                                            destPath = Path.Combine(destFolder, $"{nameNoExt}_{UserId}_{DateTime.Now:yyyyMMddHHmmss}{ext}")

                                            Try

                                                ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                                File.Move(srcPath, destPath)

                                                '取込ファイルテーブル(imp_files)に取込ファイルワークテーブル(imp_files_stage)のレコードを追加する
                                                Dim newimpfileid As Long = _impFileStageRepo.InsertImpFileFromStage(tran, impFileStageId, now, UserId, pgId)

                                                '取込ファイルワークテーブルを削除する
                                                _impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                                '正規データ更新 (ORDERSテーブルのIMP_FILE_IDを更新する)
                                                _oderStageRepo.UpdateOrdersImpFileId(tran, newimpfileid, impFileStageId)

                                                'foundInThisCustomer = True

                                            Catch ex As UnauthorizedAccessException
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（アクセス権限不足：{ex.Message}）")
                                                ErrFlg = True
                                            Catch ex As IOException
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（I/O：{ex.Message}）")
                                                ErrFlg = True
                                            Catch ex As Exception
                                                errors.Add($" 取引先コード：{customerCode}　取込ファイル：[{fileName} ]　の移動に失敗（{ex.Message}）")
                                                ErrFlg = True
                                            End Try

                                            ''取込ファイルテーブル(imp_files)に取込ファイルワークテーブル(imp_files_stage)のレコードを追加する
                                            'Dim newimpfileid As Long = _impFileStageRepo.InsertImpFileFromStage(tran, impFileStageId, now, UserId, pgId)

                                            ''取込ファイルワークテーブルを削除する
                                            '_impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                            ''正規データ更新 (ORDERSテーブルのIMP_FILE_IDを更新する)
                                            '_oderStageRepo.UpdateOrdersImpFileId(tran, newimpfileid, impFileStageId)

                                        End If

                                    Catch ex As Exception
                                        errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")
                                        ErrFlg = True
                                    End Try

                                    ''取込ファイルテーブル(imp_files)に取込ファイルワークテーブル(imp_files_stage)のレコードを追加する
                                    'Dim newimpfileid As Long = _impFileStageRepo.InsertImpFileFromStage(tran, impFileStageId, now, UserId, pgId)

                                    ''取込ファイルワークテーブルを削除する
                                    '_impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                    ''正規データ更新 (ORDERSテーブルのIMP_FILE_IDを更新する)
                                    '_oderStageRepo.UpdateOrdersImpFileId(tran, newimpfileid, impFileStageId)

                                Next

                                If ErrFlg = False Then

                                    '受注ワーク(取込(加工)済みデータ)削除
                                    _oderStageRepo.DeleteProcessedOrdersRange(tran, UserId, customerSettingId)

                                    resultCnt += 1
                                    resultRowCnt += cnt
                                End If

                                If tran IsNot Nothing Then
                                    If ErrFlg = True Then
                                        'ロールバック
                                        tran.Rollback()
                                    Else
                                        'コミット
                                        tran.Commit()
                                        'tran.Rollback()
                                    End If
                                End If

                            Catch ex As Exception

                                Dim exmes As String = ex.Message
                                ' ロールバック処理
                                If tran IsNot Nothing Then
                                    tran.Rollback()
                                End If

                                'errors.Add($"エラーが発生したため登録実行されません。")
                                errors.Add($"システムエラー({ex.Message})")

                            Finally

                                If tran IsNot Nothing Then
                                    tran.Dispose()
                                End If

                            End Try

                        End Using

                    Next

                End If

            End Using

            If resultCnt > 0 Then

                lblSaveResult.Text &= $"（ORDERS 登録 {resultRowCnt} 件）"

                'IMP_RUN_IDを取得
                runid = _impRunRepo.GetImpRunId("RUNNING", UserId, "OrderImport(Execute)")

                'IMP_RUNを更新
                'Dim rowsForTemp As New List(Of ImpRunRow) From {
                '        New ImpRunRow With {
                '            .ImpRunId = runid,
                '            .EndedAt = now,
                '            .Status = "COMPLETED",
                '            .FileCount = resultCnt,
                '            .RowCount = resultRowCnt,
                '            .ErrorCount = 0
                '        }
                '    }

                '_impRunRepo.UpdateRange(rowsForTemp)
                _impRunRepo.UpdateRange(runid, "COMPLETED", EndedAt:=now, FileCount:=resultCnt, RowCount:=resultRowCnt)

            Else
                If errors.Count = 0 Then
                    lblSaveResult.Text = "保存：対象データが見つかりませんでした。"
                End If
            End If

            If errors.Count > 0 Then
                'Dim alreadyText = lblImportError.Text
                Dim alreadyText = "（ORDERS 登録時にエラー）"
                Dim addText = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
                If String.IsNullOrEmpty(alreadyText) Then
                    lblSaveError.Text = addText
                Else
                    lblSaveError.Text = alreadyText & "<br/>" & addText
                End If
            End If


            'lblSaveResult.Text = ""
            'lblImportError.Text = ""

            'グリッド再描画
            gvImpFilesStage_Init()
            gvImportOrder_Init()

        End Sub

        ' GridViewデータバインド
        Private Sub gvImportOrder_Init()

            'Dim repo As New OrderRepository(Utils.GetConnectionString())
            'Dim dt As DataTable = repo.GetOrders(
            '                            status:="PROCESSED",
            '                            prodMgmtUserId:=PageHelpers.GetUserId(Me))
            'gvImportOrder.DataSource = dt
            'gvImportOrder.DataBind()

            Dim repo As New OrderStageRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetOrdersStage(
                                        status:="PROCESSED",
                                        prodMgmtUserId:=PageHelpers.GetUserId(Me),
                                        activeFlag:="Y")
            gvImportOrder.DataSource = dt
            gvImportOrder.DataBind()

        End Sub
        ''' <summary>文字コード名をEncodingに変換（簡易マッピング＋同義語を許容）</summary>
        Private Function MapEncoding(code As String) As Encoding
            If String.IsNullOrWhiteSpace(code) Then
                Return New UTF8Encoding(False) ' 既定はUTF-8 (BOMなし)
            End If

            Dim key = code.Trim().ToUpperInvariant()

            Select Case key
                Case Constants.ENCODING_SJIS, "SHIFT_JIS", "SJIS"
                    Return Encoding.GetEncoding("Shift_JIS")

                Case Constants.ENCODING_UTF8, "UTF-8", "UTF8"
                    Return New UTF8Encoding(False) ' BOMなし

                Case Constants.ENCODING_UTF8_BOM, "UTF-8-BOM", "UTF8-BOM", "UTF8BOM"
                    Return New UTF8Encoding(True) ' BOMあり

                Case Else
                    ' 可能ならそのままGetEncodingに委譲（例: "ISO-8859-1"等）
                    Try
                        Return Encoding.GetEncoding(code)
                    Catch
                        Return New UTF8Encoding(False)
                    End Try
            End Select
        End Function

        Private Sub AddErrorDetails(ByRef errList As List(Of String), ByVal row As GridViewRow)
            Dim dk = gvImpFilesStage.DataKeys(row.RowIndex)

            ' 各項目を個別に Add する（これで各項目が独立した行になる）
            errList.Add($" - 取引先コード：{dk("CustomerCode")}")
            errList.Add($" - PC：{dk("ProfitCenter")}")
            errList.Add($" - 注文工場/担当者コード：{dk("CustomerUnitName")}")
            errList.Add($" - フォルダ区分：{OMS.Common.Utils.ToFolderTypeNameSafe(dk("FolderType"))}")
        End Sub

    End Class
End Namespace