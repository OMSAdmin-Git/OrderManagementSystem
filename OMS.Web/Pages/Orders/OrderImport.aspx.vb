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
            Dim nodata As New List(Of String)()

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
                        Dim spprocesstype As Integer


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

                            '20260819 yamaha Phase2
                            spprocesstype = 0
                            csidObj = keys("SpProcessType")
                            If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), spprocesstype) Then
                                errors.Add($"Row {idx}：SpProcessTypeが不正")
                                Continue For
                            End If
                            '--



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

                            ''-----------------
                            ''受注ワーク削除 ※customerSettingId、folderType単位で削除
                            ''-----------------
                            '_oderStageRepo.DeleteRange(tran, customerSettingId, folderType)


                            Try



                                'マッピングマスタ取得処理
                                Dim mapError As String = ""
                                Dim mapResult As OMS.Data.OrderStageImport.MappingResult = OMS.Data.OrderStageImport.ResolveMapping(_mappingRepo, customerSettingId, folderType, errors)

                                'Phase2対応 spprocesstype=1or2or3はマッピングマスタを使用しない
                                'If mapResult Is Nothing Then
                                'If mapResult Is Nothing And spprocesstype = 0 Then
                                '    '特殊加工なし
                                '    errors.Add($"顧客設定ID:{customerSettingId} - {mapError}")
                                '    Continue For
                                'ElseIf mapResult Is Nothing And spprocesstype = 1 And folderType <> 4 Then
                                '    '特殊加工:スズキ フォルダ区分:混合以外
                                '    errors.Add($"顧客設定ID:{customerSettingId} - {mapError}")
                                '    Continue For
                                'ElseIf mapResult Is Nothing And spprocesstype = 2 And folderType <> 4 Then
                                '    '特殊加工:ヤマハ フォルダ区分:混合以外
                                '    errors.Add($"顧客設定ID:{customerSettingId} - {mapError}")
                                '    Continue For
                                'End If

                                If mapResult IsNot Nothing AndAlso (spprocesstype = 0 OrElse
                                                                    (spprocesstype = 1 AndAlso chkHandFlag.Checked = True) OrElse
                                                                    (spprocesstype = 2 AndAlso chkHandFlag.Checked = True) OrElse
                                                                    (spprocesstype = 3 AndAlso chkHandFlag.Checked = True)) Then

                                    '特殊処理以外(通常の取込実行)　または　特殊処理でハンドフラグにチェックが入っている場合(特殊処理のASTI追加内示)

                                    '取込ファイルからデータを取得する処理
                                    OMS.Data.OrderStageImport.ParseImportFile(
                                                                tran,
                                                                customerSettingId,
                                                                customerCode,
                                                                impfilestageId,
                                                                spprocesstype,
                                                                strWorkFile,
                                                                TorikomiFile,
                                                                ErrFlg,
                                                                ErrFileFlg,
                                                                errcnt,
                                                                folderType,
                                                                newId,
                                                                UserId,
                                                                pgId,
                                                                errors,
                                                                rowsForTemp2,
                                                                mapResult)

                                ElseIf mapResult Is Nothing AndAlso spprocesstype = 1 AndAlso folderType = 4 Then
                                    '特殊加工:スズキ フォルダ区分:混合

                                ElseIf mapResult Is Nothing AndAlso spprocesstype = 2 AndAlso folderType = 4 Then
                                    '特殊加工:ヤマハ(IM以外) フォルダ区分:混合

                                    '取込ファイルからデータを取得する処理
                                    OMS.Data.OrderStageImport.ParseImportFileY(
                                                                tran,
                                                                customerSettingId,
                                                                customerCode,
                                                                impfilestageId,
                                                                spprocesstype,
                                                                strWorkFile,
                                                                TorikomiFile,
                                                                ErrFlg,
                                                                ErrFileFlg,
                                                                errcnt,
                                                                folderType,
                                                                newId,
                                                                UserId,
                                                                pgId,
                                                                errors,
                                                                rowsForTemp2)

                                ElseIf mapResult Is Nothing AndAlso spprocesstype = 3 AndAlso folderType <> 4 Then
                                    ' Yamaha robotex 内示/確定/ASTI内示取得
                                    OMS.Data.OrderStageImport.YamahaRobotexOrdersStageImport(tran,
                                                                                            customerSettingId,
                                                                                            customerCode,
                                                                                            impfilestageId,
                                                                                            spprocesstype,
                                                                                            strWorkFile,
                                                                                            TorikomiFile,
                                                                                            ErrFlg,
                                                                                            ErrFileFlg,
                                                                                            errcnt,
                                                                                            folderType,
                                                                                            newId,
                                                                                            UserId,
                                                                                            pgId,
                                                                                            errors,
                                                                                            rowsForTemp2)
                                ElseIf mapResult Is Nothing Then

                                    'errors.Add($"顧客設定ID:{customerSettingId} - {mapError}")
                                    errors.Add($"顧客設定ID:{customerSettingId}:MAPPINNG_PROFILE_MSTに未登録")
                                    Continue For

                                End If




                                'ハンドフラグの状態を取得
                                Dim blnHandFlag As Boolean = False
                                If chkHandFlag.Checked = True Then
                                    blnHandFlag = True
                                Else
                                    blnHandFlag = False
                                End If

                                '---------------------------------
                                'ORDERS_STAGEテーブルへ登録処理
                                '---------------------------------
                                Dim importResult As OMS.Data.OrderStageImport =
                                                        OMS.Data.OrderStageImport.OrdersStageSaved(
                                                            tran,
                                                            customerSettingId,
                                                            impfilestageId,
                                                            folderType,
                                                            reconcileFlag,
                                                            fcstreconcileFlag,
                                                            blnHandFlag,
                                                            UserId,
                                                            pgId,
                                                            rowsForTemp2,
                                                            spprocesstype
                                    )
                                Dim savedCount As Integer = importResult.InsertedCount

                                If savedCount > 0 Then

                                    'If ErrFlg = False Then

                                    resultCnt += 1
                                    resultRowCnt += savedCount
                                    successs.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　読込 {savedCount} 件　異常 {errcnt} 件")

                                    'End If

                                Else

                                    '取込対象が一件も無い場合、
                                    nodata.Add($" 取引先コード：{customerCode}　取込ファイル：[{TorikomiFile} ]　対象データが1件も無いため破棄してください。")

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
                    'If errors.Count = 0 Then
                    '    'lblImportResult.Text = "取込実行：予期せぬエラーが発生しました。"
                    '    lblImportResult.Text = "取込実行：取込対象データがありません。"
                    'End If
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

            If nodata.Count > 0 Then
                Dim alreadyText = lblImportError.Text & "<br/>" & "（取込実行：取込対象データがありません）"
                Dim addText = String.Join("<br/>", nodata.Select(Function(s) Server.HtmlEncode(s)))
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

                                            ''取込ファイルワークテーブルを削除する
                                            '_impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                            'resultCnt += 1

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

                                '取込ファイルワークテーブルを削除する
                                _impFileStageRepo.DeleteImpFileStageRange(tran, impFileStageId)

                                'resultCnt += 1

                                If ErrFlg = False Then

                                    '受注ワーク(取込(加工)済みデータ)削除　※取込ファイルID単位で削除
                                    'cnt = _oderStageRepo.DeleteProcessedOrdersRange(tran, UserId, customerSettingId)
                                    cnt = _oderStageRepo.DeleteProcessedOrdersByFileId(tran, UserId, impFileStageId)

                                    'resultCnt += 1
                                    'resultRowCnt += cnt

                                    resultCnt += cnt

                                End If

                                'If tran IsNot Nothing Then
                                '    If ErrFlg = True Then
                                '        'ロールバック
                                '        tran.Rollback()
                                '    Else
                                '        'コミット
                                '        tran.Commit()

                                '    End If
                                'End If

                                'コミット
                                tran.Commit()

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

                ''グリッド再描画
                'gvImpFilesStage_Init()
                'gvImportOrder_Init()

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
            gvImpFilesStage_Init()
            gvImportOrder_Init()

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