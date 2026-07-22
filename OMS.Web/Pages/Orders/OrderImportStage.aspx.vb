Imports System
Imports System.Data
Imports System.IO
Imports System.Configuration
Imports System.Linq
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports OMS.Common
Imports OMS.Data

Namespace Pages.Orders
    Public Class OrderImportStage
        Inherits System.Web.UI.Page

#Region "フィールド"
        Private _folderRepo As FolderRepository
        Private _impTempRepo As ImpFilesStageRepository
        Private _workRootResolved As String
        Private _workUserRoot As String
#End Region

#Region "ページ初期化（Page_Init）"
        Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
            ' DB接続の取得
            Dim csSetting = ConfigurationManager.ConnectionStrings("OMSConnection")
            If csSetting Is Nothing OrElse String.IsNullOrWhiteSpace(csSetting.ConnectionString) Then
                Throw New ConfigurationErrorsException("connectionStrings['OMSConnection'] が未定義です。Web.config を確認してください。")
            End If
            Dim connStr As String = csSetting.ConnectionString

            ' リポジトリの初期化
            _folderRepo = New FolderRepository(connStr)
            _impTempRepo = New ImpFilesStageRepository(connStr)

            ' WORKフォルダの解決・作成
            Dim rawWork As String = ConfigurationManager.AppSettings("WorkFolderRoot")
            If String.IsNullOrWhiteSpace(rawWork) Then
                Throw New ConfigurationErrorsException("appSettings['WorkFolderRoot'] が未定義です。Web.config を確認してください。")
            End If

            ' 相対/UNC/絶対のいずれでも対応（Utils.ResolvePath を使用）
            _workRootResolved = Utils.ResolvePath(Me.Server, rawWork)

            ' フォルダが存在しなければ作成
            Utils.EnsureDirectory(_workRootResolved)

            ' ユーザーID取得（実装が決まり次第置き換え）
            'Dim userId As String = GetCurrentUserId()
            Dim userId As String = PageHelpers.GetUserId(Me)

            ' フォルダ名として安全化（Windowsパスの禁止文字などを除去／長さ制限）
            Dim safeUserId As String = Utils.SafeFolderName(userId, maxLength:=32)

            ' ユーザー単位のWORKルートを作成
            _workUserRoot = System.IO.Path.Combine(_workRootResolved, safeUserId)
            Utils.EnsureDirectory(_workUserRoot)
        End Sub
#End Region

#Region "パス解決ヘルパー（参考：未使用だが残置）"
        ''' <summary>
        ''' FOLDER_MSTの値が仮想パス（~/、/）か物理パス（UNC含む）かに応じて実パスへ解決
        ''' </summary>
        Private Function ResolveFolderPath(raw As String) As String
            If String.IsNullOrWhiteSpace(raw) Then Return raw

            ' 絶対/UNC/ドライブなどはそのまま
            If System.IO.Path.IsPathRooted(raw) Then
                Return raw
            End If

            ' App相対（~/ や ~\）
            If raw.StartsWith("~/") OrElse raw.StartsWith("~\") Then
                Return Server.MapPath(raw)
            End If

            ' 先頭 / または \ はアプリルート相対
            If raw.StartsWith("/") OrElse raw.StartsWith("\") Then
                Return Server.MapPath(raw)
            End If

            ' 相対文字列はアプリルート基準で解決
            Return Server.MapPath("~/" & raw.TrimStart("/"c, "\"c))
        End Function
#End Region

#Region "ページ ロード（Page_Load）"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)

                ' 検索候補（顧客コード／PC／ユニット名）を初期化
                LoadSearchConditionLists()

                ' 初期表示（一覧）
                BindSelectCustomersGrid()
            End If
        End Sub
#End Region

#Region "検索候補：リスト初期化"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim repo As New CustomerRepository(Utils.GetConnectionString())

            Dim customerCodeList As List(Of String) = repo.GetCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = repo.GetProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

            lstSearchCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstSearchProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstSearchCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "ナビゲーション / 検索ボタン"
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

            BindSelectCustomersGrid(customerCode, customerName, profitCenter, customerUnitName)
        End Sub

        ' クリアボタン
        Protected Sub btnDefaultGv_Click(sender As Object, e As EventArgs)
            txtSearchCustomerCode.Value = ""
            txtSearchCustomerName.Value = ""
            txtSearchProfitCenter.Value = ""
            txtSearchCustomerUnitName.Value = ""

            BindSelectCustomersGrid()
        End Sub
#End Region

#Region "GridView バインド / イベント"
        ' 顧客候補の一覧をバインド
        Private Sub BindSelectCustomersGrid(
            Optional ByVal customerCode As String = Nothing,
            Optional ByVal customerName As String = Nothing,
            Optional ByVal profitCenter As String = Nothing,
            Optional ByVal customerUnitName As String = Nothing
        )
            Dim repo As New CustomerRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetCustomerImpRuleList(
                customerCode:=customerCode,
                customerName:=customerName,
                profitCenter:=profitCenter,
                customerUnitName:=customerUnitName,
                prodMgmtUserId:=PageHelpers.GetUserId(Me),
                activeFlag:="Y"
            )

            gvSelectCustomers.DataSource = dt
            gvSelectCustomers.DataBind()
        End Sub

        ' GridViewヘッダーバインド
        Protected Sub gvSelectCustomers_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvSelectCustomers.RowDataBound
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox = TryCast(e.Row.FindControl("chkDueDateSetting"), CheckBox)
                If chk IsNot Nothing Then
                    ' 個別のチェック操作時にヘッダー状態を更新
                    chk.InputAttributes("onclick") =
                        $"OMS.Grid.updateHeader('{gvSelectCustomers.ClientID}', 'chkDueDateSettingAll', 'chkDueDateSetting');"
                End If
            End If
        End Sub
#End Region

#Region "取込準備（ステージング）"
        ' 取込準備ボタン（ロジックは既存のまま）
        Protected Sub btnStageImport_Click(sender As Object, e As EventArgs)

            lblResult.Text = ""
            lblError.Text = ""

            Dim results As New List(Of ImpFilesStageResult)()
            Dim anyFound As Boolean = False
            Dim errors As New List(Of String)()

            ' 取引先選択行を走査
            For Each row As GridViewRow In gvSelectCustomers.Rows
                If row.RowType <> DataControlRowType.DataRow Then
                    Continue For ' データ行以外はスキップ
                End If

                Dim chk As CheckBox = TryCast(row.FindControl("chkStageImport"), CheckBox)
                If chk Is Nothing OrElse Not chk.Checked Then
                    Continue For ' 未選択行はスキップ
                End If

                Dim idx As Integer = row.RowIndex
                Dim keys = gvSelectCustomers.DataKeys(idx)
                If keys Is Nothing Then
                    errors.Add($"Row {idx}：DataKeys未設定")
                    Continue For
                End If

                ' 取得＆安全化
                Dim customerSettingId As Integer
                Dim csidObj = keys("CustomerSettingId")
                If csidObj Is Nothing OrElse Not Integer.TryParse(csidObj.ToString(), customerSettingId) Then
                    errors.Add($"Row {idx}：CustomerSettingIdが不正")
                    Continue For
                End If

                Dim customerCode As String = keys("CustomerCode")?.ToString()
                Dim customerName As String = keys("CustomerName")?.ToString()
                Dim profitCenter As String = keys("ProfitCenter")?.ToString()
                Dim customerUnitId As String = keys("CustomerUnitId")?.ToString()
                Dim customerUnitName As String = keys("CustomerUnitName")?.ToString()

                ' DropDownList を取得（IDは .aspx のテンプレート列に合わせる）
                Dim ddlReconcileFlag = TryCast(row.FindControl("ddlReconcileFlag"), DropDownList)
                Dim ddlFcstReconcileFlag = TryCast(row.FindControl("ddlFcstReconcileFlag"), DropDownList)
                Dim ddlSelfFcstDeleteFlag = TryCast(row.FindControl("ddlSelfFcstDeleteFlag"), DropDownList)

                Dim reconcileVal As String = Utils.NormalizeYN(ddlReconcileFlag?.SelectedValue)
                Dim fcstReconcileVal As String = Utils.NormalizeYN(ddlFcstReconcileFlag?.SelectedValue)

                Try
                    ' [フォルダパス]＋[受注区分]を取得
                    Dim folderInfos As List(Of FolderPathInfo) = _folderRepo.GetFolderInfosByCustomerSettingId(customerSettingId)
                    If folderInfos Is Nothing OrElse folderInfos.Count = 0 Then
                        errors.Add($"{customerCode}：FOLDER_MSTにフォルダ未登録")
                        Continue For
                    End If

                    Dim foundInThisCustomer As Boolean = False

                    For Each info In folderInfos
                        Dim sourceFolder As String = Utils.ResolvePath(Me.Server, info.FolderPath)

                        ' フォルダ存在確認
                        If Not Directory.Exists(sourceFolder) Then
                            errors.Add($"{customerCode}：フォルダが存在しません [{Server.HtmlEncode(sourceFolder)}]")
                            Continue For
                        End If

                        ' WORKサブフォルダ作成
                        Dim destFolder As String = Path.Combine(_workUserRoot, customerCode, info.FolderType.ToString())
                        Utils.EnsureDirectory(destFolder)

                        Dim files = Directory.EnumerateFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly) _
                            .Concat(Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly)) _
                            .Concat(Directory.EnumerateFiles(sourceFolder, "*.txt", SearchOption.TopDirectoryOnly))

                        For Each src In files
                            Dim fileName = Path.GetFileName(src)
                            Dim destPath = Path.Combine(destFolder, fileName)

                            ' 衝突回避：同名が既にある場合はタイムスタンプを付ける
                            If File.Exists(destPath) Then
                                Dim nameNoExt = Path.GetFileNameWithoutExtension(fileName)
                                Dim ext = Path.GetExtension(fileName)
                                destPath = Path.Combine(destFolder, $"{nameNoExt}_{DateTime.Now:yyyyMMddHHmmssfff}{ext}")
                            End If

                            Try
                                ' 実移動（同一ボリューム/別ボリュームどちらでもOK）
                                File.Move(src, destPath)

                                ' 移動成功→結果に追加（画面表示用）
                                Dim lw = File.GetLastWriteTime(destPath)
                                results.Add(New ImpFilesStageResult With {
                                    .CustomerSettingId = customerSettingId,
                                    .CustomerCode = customerCode,
                                    .CustomerName = customerName,
                                    .ProfitCenter = profitCenter,
                                    .CustomerUnitId = customerUnitId,
                                    .CustomerUnitName = customerUnitName,
                                    .FolderType = info.FolderType,
                                    .FolderPath = sourceFolder,
                                    .FileName = fileName,
                                    .StagedFolderPath = destFolder,
                                    .StagedFilePath = destPath,
                                    .StagedFileName = Path.GetFileName(destPath),
                                    .Status = "DISCOVERED",
                                    .LastWriteTime = lw,
                                    .ReconcileFlag = reconcileVal,
                                    .FcstReconcileFlag = fcstReconcileVal,
                                    .HandFlag = "N"
                                })

                                foundInThisCustomer = True

                            Catch ex As UnauthorizedAccessException
                                errors.Add($"{customerCode}：{fileName} の移動に失敗（アクセス権限不足：{ex.Message}）")
                            Catch ex As IOException
                                errors.Add($"{customerCode}：{fileName} の移動に失敗（I/O：{ex.Message}）")
                            Catch ex As Exception
                                errors.Add($"{customerCode}：{fileName} の移動に失敗（{ex.Message}）")
                            End Try
                        Next
                    Next

                    If foundInThisCustomer Then
                        anyFound = True
                    End If

                Catch ex As Exception
                    errors.Add($"{customerCode}：{Server.HtmlEncode(ex.Message)}")
                End Try
            Next

            ' 画面へバインド（DataKeyNames="FilePath" 必須）
            Session("OrderImportSearchResults") = results

            If errors.Count > 0 Then
                lblError.Text = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
            End If

            If anyFound AndAlso results.Count > 0 Then
                lblResult.Text = $"取込準備：対象ファイルが【 {results.Count} 件 】見つかりました。"
            Else
                If errors.Count = 0 Then
                    lblResult.Text = "取込準備：対象ファイルが見つかりませんでした。"
                End If
                ' 対象がなければ以降のDB登録処理はスキップ
                Exit Sub
            End If

            ' GridViewの選択状態（DropDownList）を読み取り、IMP_FILES_STAGE へ登録（Y/N管理）
            Dim now As DateTime = DateTime.Now
            'Dim userId As String = (If(Context?.User?.Identity?.Name, "")).Trim()
            'If String.IsNullOrWhiteSpace(userId) Then
            '    userId = "AMAGATA"
            'End If
            'If userId.Length > 9 Then
            '    userId = userId.Substring(0, 9)
            'End If
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim pgId As String = "OrderImport(Stage)"

            Dim rowsForTemp As New List(Of ImpFilesStageRow)()
            If results IsNot Nothing AndAlso results.Count > 0 Then
                rowsForTemp = results.Select(Function(r) New ImpFilesStageRow With {
                    .CustomerSettingId = r.CustomerSettingId,
                    .FolderType = r.FolderType,
                    .FolderPath = r.FolderPath,
                    .FileName = r.FileName,
                    .StagedFolderPath = r.StagedFolderPath,
                    .StagedFileName = r.StagedFileName,
                    .ReconcileFlag = r.ReconcileFlag,
                    .FcstReconcileFlag = r.FcstReconcileFlag,
                    .HandFlag = r.HandFlag,
                    .Status = r.Status,
                    .CreatedAt = now,
                    .CreatedUserId = loginUserId,
                    .CreatedPgId = pgId,
                    .UpdatedAt = now,
                    .UpdatedUserId = loginUserId,
                    .UpdatedPgId = pgId
                }).ToList()
            End If

            If rowsForTemp.Count > 0 Then
                Try
                    _impTempRepo.InsertRange(rowsForTemp)
                    lblResult.Text &= $"（IMP_FILES_STAGE 登録 {rowsForTemp.Count} 件）"
                Catch ex As Exception
                    errors.Add($"IMP_FILES_STAGE登録時にエラー：{ex.Message}")
                End Try
            End If

            If errors.Count > 0 Then
                Dim alreadyText = lblError.Text
                Dim addText = String.Join("<br/>", errors.Select(Function(s) Server.HtmlEncode(s)))
                If String.IsNullOrEmpty(alreadyText) Then
                    lblError.Text = addText
                Else
                    lblError.Text = alreadyText & "<br/>" & addText
                End If
            End If

        End Sub
#End Region

    End Class
End Namespace