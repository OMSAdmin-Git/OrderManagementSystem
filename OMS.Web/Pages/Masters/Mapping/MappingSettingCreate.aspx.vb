Imports System.Collections.Specialized
Imports System.Data
Imports System.Web.UI.WebControls
Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.Mapping
    Public Class MappingSettingCreate
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        Private Const SESSION_KEY As String = "FieldMappingDT"
        Private Const DUMMY_ROW_FLAG As String = "__IsDummy__"
#End Region

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ResetGrid()
                ' ユーザー名表示
                PageHelpers.SetUserName(Me, lblUser)
                ' 検索候補を初期化
                LoadSearchConditionLists()
                EnsureSessionTable()
                BindFieldMappingGrid()
            End If
        End Sub

        Protected Overrides Sub OnPreRender(e As EventArgs)
            MyBase.OnPreRender(e)
            ' lblError が DOM に出るタイミングで確実に ClientID が決まっている
            btnCreateMappingSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

#Region "検索候補：リスト初期化"
        Private Sub LoadSearchConditionLists()
            Dim loginUserId As String = PageHelpers.GetUserId(Me)
            Dim repo As New CustomerRepository(Utils.GetConnectionString())

            Dim customerCodeList As List(Of String) = repo.GetCustomerCodes(loginUserId)
            Dim profitCenterList As List(Of String) = repo.GetProfitCenters(loginUserId)
            Dim customerUnitNameList As List(Of String) = repo.GetCustomerUnitNames(loginUserId)

            lstCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "Session管理 / バインド（新規専用）"
        ' 新規専用：画面で使う空スキーマを作成（DBからは読み込まない）
        Private Function BuildEmptyFieldMappingTableSchema() As DataTable
            Dim dt As New DataTable("FieldMapping")

            ' 画面内ユニークキー
            dt.Columns.Add("TempId", GetType(String))

            ' DB主キー（新規は DBNull でOK）
            dt.Columns.Add("FieldMappingId", GetType(String))

            ' 保存時に確定するため、初期は DBNull
            dt.Columns.Add("ProfileId", GetType(String))

            ' 明細項目（Gridで扱う列）
            dt.Columns.Add("FormatType", GetType(String))
            dt.Columns.Add("TargetField", GetType(String))
            dt.Columns.Add("SourceColumnIndex", GetType(Integer))
            dt.Columns.Add("SourceHeaderName", GetType(String))
            dt.Columns.Add("SourceSheetName", GetType(String))
            dt.Columns.Add("SourceCellAddress", GetType(String))
            dt.Columns.Add("RowSelectorType", GetType(String))
            dt.Columns.Add("RowSelectorValue", GetType(String))
            dt.Columns.Add("DataType", GetType(String))
            dt.Columns.Add("FormatPattern", GetType(String))
            dt.Columns.Add("ActiveFlag", GetType(String))

            ' 0件時ダミー行判定
            dt.Columns.Add(DUMMY_ROW_FLAG, GetType(Boolean))

            ' 既定値／NULL許容（値型は NULL 許容にしておく）
            dt.Columns("ActiveFlag").DefaultValue = "Y"
            dt.Columns("SourceColumnIndex").AllowDBNull = True

            Return dt
        End Function

        Private Function GetSessionTable() As DataTable
            Return TryCast(Session(SESSION_KEY), DataTable)
        End Function

        Private Sub SetSessionTable(dt As DataTable)
            Session(SESSION_KEY) = dt
        End Sub

        ' 初回は空スキーマを準備（DBロードはしない）
        Private Sub EnsureSessionTable()
            Dim dt = GetSessionTable()
            If dt IsNot Nothing Then Exit Sub

            dt = BuildEmptyFieldMappingTableSchema()
            dt.AcceptChanges() ' 起点
            SetSessionTable(dt)
        End Sub

        ' この画面の明細グリッドだけをバインド（0件時はダミー行を注入してFooterを出す）
        Private Sub BindFieldMappingGrid()
            EnsureSessionTable()

            Dim dt As DataTable = GetSessionTable()
            If dt Is Nothing Then
                dt = BuildEmptyFieldMappingTableSchema()
                SetSessionTable(dt)
            End If

            ' 0件（CurrentRows）であれば、Footerを出すためのダミー行を1件追加
            Dim currentCount As Integer = dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length
            If currentCount = 0 Then
                Dim dummy As DataRow = dt.NewRow()
                dummy("TempId") = "DUMMY"
                dummy("FieldMappingId") = DBNull.Value
                dummy("ProfileId") = DBNull.Value
                dummy(DUMMY_ROW_FLAG) = True
                ' 他列は AllowDBNull=True のため未設定でOK
                dt.Rows.Add(dummy)
            End If

            Dim dv As New DataView(dt, Nothing, Nothing, DataViewRowState.CurrentRows)
            gvFieldMappingList.DataSource = dv
            gvFieldMappingList.DataBind()
        End Sub

        'グリッドを初期化する（フォームを開く時に使用）
        Private Sub ResetGrid()
            Dim dt = GetSessionTable()
            If dt Is Nothing Then Exit Sub
            Session.Remove(SESSION_KEY)
        End Sub
#End Region

#Region "GridView イベント（新規＋編集対応）"
        ' フッターの DDL 初期化 + 0件時ダミー行の非表示 + 編集行DDLの初期化
        Protected Sub gvFieldMappingList_RowDataBound(sender As Object, e As GridViewRowEventArgs) _
            Handles gvFieldMappingList.RowDataBound

            If e.Row.RowType = DataControlRowType.DataRow Then
                ' ダミー行なら非表示（0件時にFooterだけを見せる）
                Dim drv As DataRowView = TryCast(e.Row.DataItem, DataRowView)
                If drv IsNot Nothing AndAlso drv.Row.Table.Columns.Contains(DUMMY_ROW_FLAG) Then
                    Dim isDummy As Boolean = False
                    If Not Convert.IsDBNull(drv(DUMMY_ROW_FLAG)) Then
                        Boolean.TryParse(drv(DUMMY_ROW_FLAG).ToString(), isDummy)
                    End If
                    If isDummy Then
                        e.Row.Visible = False
                        Return
                    End If
                End If

                ' 編集行のみ DDL をバインド（Createでも編集は可能）
                If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                    BindAndSelect(e.Row, "ddlFormatType", OMS.Common.Constants.FormatTypeMap, DataBinder.Eval(e.Row.DataItem, "FormatType"))
                    BindAndSelect(e.Row, "ddlTargetField", OMS.Common.Constants.TargetFieldMap, DataBinder.Eval(e.Row.DataItem, "TargetField"))
                    BindAndSelect(e.Row, "ddlRowSelectorType", OMS.Common.Constants.RowSelectorTypeMap, DataBinder.Eval(e.Row.DataItem, "RowSelectorType"))
                    BindAndSelect(e.Row, "ddlDataType", OMS.Common.Constants.DataTypeMap, DataBinder.Eval(e.Row.DataItem, "DataType"))
                    BindAndSelect(e.Row, "ddlActiveFlag", OMS.Common.Constants.ActiveFlagMap, DataBinder.Eval(e.Row.DataItem, "ActiveFlag"))
                End If

            ElseIf e.Row.RowType = DataControlRowType.Footer Then
                ' フッター行（新規追加用）
                OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlFormatType_F"), DropDownList), OMS.Common.Constants.FormatTypeMap)
                OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlTargetField_F"), DropDownList), OMS.Common.Constants.TargetFieldMap)
                OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlRowSelectorType_F"), DropDownList), OMS.Common.Constants.RowSelectorTypeMap)
                OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlDataType_F"), DropDownList), OMS.Common.Constants.DataTypeMap)
                'OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlTrimFlag_F"), DropDownList), OMS.Common.Constants.TrimFlagMap)
                OMS.Common.Constants.BindDropDown(TryCast(e.Row.FindControl("ddlActiveFlag_F"), DropDownList), OMS.Common.Constants.ActiveFlagMap, sortByLabel:=False)
            End If
        End Sub

        ' 編集開始
        Protected Sub gvFieldMappingList_RowEditing(sender As Object, e As GridViewEditEventArgs) _
            Handles gvFieldMappingList.RowEditing
            gvFieldMappingList.EditIndex = e.NewEditIndex
            BindFieldMappingGrid()
        End Sub

        ' 編集キャンセル
        Protected Sub gvFieldMappingList_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) _
            Handles gvFieldMappingList.RowCancelingEdit
            e.Cancel = True
            gvFieldMappingList.EditIndex = -1
            BindFieldMappingGrid()
        End Sub

        ' 編集確定（画面内 DataTable を更新）
        Protected Sub gvFieldMappingList_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) _
            Handles gvFieldMappingList.RowUpdating
            Dim dt As DataTable = GetSessionTable()
            Dim row As GridViewRow = gvFieldMappingList.Rows(e.RowIndex)

            ' TempId で行特定（Createは NEW: のみ）
            Dim tempId As String = gvFieldMappingList.DataKeys(e.RowIndex).Values("TempId").ToString()
            Dim dr As DataRow = dt.AsEnumerable().
                First(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            ' DDL は FindControl、テキストは e.NewValues から
            dr("FormatType") = GetDDL(row, "ddlFormatType")
            dr("TargetField") = GetDDL(row, "ddlTargetField")
            dr("RowSelectorType") = GetDDL(row, "ddlRowSelectorType")
            dr("DataType") = GetDDL(row, "ddlDataType")
            dr("ActiveFlag") = GetDDL(row, "ddlActiveFlag")

            dr("SourceColumnIndex") = ToIntOrDBNull(SafeGet(e.NewValues, "SourceColumnIndex"))
            dr("SourceHeaderName") = SafeGet(e.NewValues, "SourceHeaderName")
            dr("SourceSheetName") = SafeGet(e.NewValues, "SourceSheetName")
            dr("SourceCellAddress") = SafeGet(e.NewValues, "SourceCellAddress")
            dr("RowSelectorValue") = SafeGet(e.NewValues, "RowSelectorValue")
            dr("FormatPattern") = SafeGet(e.NewValues, "FormatPattern")

            gvFieldMappingList.EditIndex = -1
            BindFieldMappingGrid()
            e.Cancel = True
        End Sub

        ' 追加（フッター行から）
        Protected Sub gvFieldMappingList_RowCommand(sender As Object, e As GridViewCommandEventArgs) _
            Handles gvFieldMappingList.RowCommand

            If Not String.Equals(e.CommandName, "Insert", StringComparison.OrdinalIgnoreCase) Then Return

            Dim dt As DataTable = GetSessionTable()
            Dim f As GridViewRow = gvFieldMappingList.FooterRow

            Dim r As DataRow = dt.NewRow()
            r("TempId") = "NEW:" & Guid.NewGuid().ToString("N")
            r("FieldMappingId") = DBNull.Value
            r("ProfileId") = DBNull.Value ' 新規登録時に確定（ヘッダーINSERT後のIDを用いる）

            r("FormatType") = TryCast(f.FindControl("ddlFormatType_F"), DropDownList)?.SelectedValue
            r("TargetField") = TryCast(f.FindControl("ddlTargetField_F"), DropDownList)?.SelectedValue
            r("SourceColumnIndex") = ToIntOrDBNull(TryCast(f.FindControl("txtSourceColumnIndex_F"), TextBox)?.Text)
            r("SourceHeaderName") = TryCast(f.FindControl("txtSourceHeaderName_F"), TextBox)?.Text
            r("SourceSheetName") = TryCast(f.FindControl("txtSourceSheetName_F"), TextBox)?.Text
            r("SourceCellAddress") = TryCast(f.FindControl("txtSourceCellAddress_F"), TextBox)?.Text
            r("RowSelectorType") = TryCast(f.FindControl("ddlRowSelectorType_F"), DropDownList)?.SelectedValue
            r("RowSelectorValue") = TryCast(f.FindControl("txtRowSelectorValue_F"), TextBox)?.Text
            r("DataType") = TryCast(f.FindControl("ddlDataType_F"), DropDownList)?.SelectedValue
            r("FormatPattern") = TryCast(f.FindControl("txtFormatPattern_F"), TextBox)?.Text
            r("ActiveFlag") = TryCast(f.FindControl("ddlActiveFlag_F"), DropDownList)?.SelectedValue

            If dt.Columns.Contains(DUMMY_ROW_FLAG) Then r(DUMMY_ROW_FLAG) = False

            dt.Rows.Add(r) ' RowState=Added
            BindFieldMappingGrid()
        End Sub

        ' 削除（新規：Remove のみ）
        Protected Sub gvFieldMappingList_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) _
            Handles gvFieldMappingList.RowDeleting
            Dim dt As DataTable = GetSessionTable()
            Dim tempId As String = gvFieldMappingList.DataKeys(e.RowIndex).Values("TempId").ToString()

            Dim dr As DataRow = dt.AsEnumerable().
                FirstOrDefault(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            If dr IsNot Nothing Then
                dt.Rows.Remove(dr)
            End If

            BindFieldMappingGrid()
            e.Cancel = True
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnMappingList_Click(sender As Object, e As EventArgs)
            Response.Redirect("MappingList.aspx")
        End Sub

        ' 登録ボタン（実装方針：ヘッダー→明細の順でINSERT）
        Protected Sub btnCreateMappingSetting_Click(sender As Object, e As EventArgs)
            Try
                ' 入力取得
                Dim customerCode As String = (If(txtCustomerCode.Value, "")).Trim()
                Dim profitCenterRaw As String = (If(txtProfitCenter.Value, "")).Trim()
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim folderType As String = If(ddlFolderType.SelectedValue, "").Trim()
                Dim version As String = (If(txtVersion.Text, "")).Trim()
                Dim headerRowIndex As String = (If(txtHeaderRowIndex.Text, "")).Trim()
                Dim dataStartRowIndex As String = (If(txtDataStartRowIndex.Text, "")).Trim()
                Dim defaultSheetName As String = (If(txtDefaultSheetName.Text, "")).Trim()

                Dim repoCust As New CustomerRepository(Utils.GetConnectionString())
                Dim repoFile As New FileRepository(Utils.GetConnectionString())
                Dim repoMappingProfile As New MappingRepository(Utils.GetConnectionString())

                Dim userId As String = PageHelpers.GetUserId(Me.Page)
                If String.IsNullOrEmpty(userId) Then
                    lblError.Text = "ログイン情報が見つかりません。"
                    lblResult.Text = ""
                    Exit Sub
                End If
                If userId.Length > 9 Then userId = userId.Substring(0, 9)

                ' 変換/解決
                Dim profitCenter As String = If(String.IsNullOrWhiteSpace(profitCenterRaw), Nothing, profitCenterRaw)
                Dim customerUnitIdNullable As Long? = Nothing
                If Not String.IsNullOrWhiteSpace(customerUnitName) Then
                    Dim customerUnitId As Long = repoCust.GetCustomerUnitIdByName(customerUnitName)
                    If customerUnitId <= 0 Then
                        lblError.Text = "注文工場／担当者名が正しく選択されていません。候補から選択してください。"
                        Return
                    End If
                    customerUnitIdNullable = customerUnitId
                End If

                ' 取引先設定IDの取得
                Dim customerSettingIdNullable As Long? = repoCust.FindCustomerSettingIdSimple(
                    customerCode:=customerCode,
                    profitCenter:=profitCenter,
                    customerUnitId:=If(customerUnitIdNullable.HasValue, customerUnitIdNullable.Value, CType(Nothing, Long?))
                )

                If Not customerSettingIdNullable.HasValue Then
                    lblError.Text = "該当する取引先設定が見つかりません。取引先コード／PC／注文工場／担当者名を確認してください。"
                    Return
                End If

                Dim customerSettingId As Long = customerSettingIdNullable.Value

                ' 重複チェック（NULL セーフ一致）
                Dim existsOther As Boolean = repoMappingProfile.ExistsMappingProfile(
                    customerSettingId:=customerSettingId,
                    folderType:=folderType,
                    excludeProfileId:=0 ' 新規のため除外なし
                )
                If existsOther Then
                    lblError.Text = "同一（取引先コード、PC、注文工場／担当者名、フォルダ区分）の登録が見つかりました。"
                    Return
                End If

                ' ファイルIDの取得
                Dim fileIdNullable As Long? = repoFile.FindFileIdSimple(
                    CustomerSettingId:=customerSettingId,
                    FolderType:=folderType
                    )
                Dim FileId As Long = fileIdNullable.Value


                ' ▼ 1) ヘッダーをINSERTして新しい ProfileId を採番（ヘッダー入力欄のID確定後に実装）
                Dim repo As New MappingRepository(Utils.GetConnectionString())
                Dim newProfileId As Long = repo.InsertMappingProfile(
                    fileId:=FileId,
                    customerSettingId:=customerSettingId,
                    folderType:=folderType,
                    version:=version,
                    headerRowIndex:=headerRowIndex,
                    dataStartRowindex:=dataStartRowIndex,
                    defaultSheetName:=defaultSheetName,
                    userId:=userId,
                    pgId:="CreateMappingSetting(Header)"
                )

                ' ▼ 2) 明細を一括INSERT（RowState=Added）
                Dim dt As DataTable = GetSessionTable()
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Added)
                    ' ダミーデータ判定（TempId = "DUMMY" の場合はスキップ）
                    If Not IsDBNull(r("TempId")) AndAlso r("TempId").ToString().Trim().ToUpper() = "DUMMY" Then
                        Continue For
                    End If

                    repo.InsertFieldMapping(
                        profileId:=newProfileId.ToString(),
                        formatType:=r("FormatType").ToString(),
                        targetField:=r("TargetField").ToString(),
                        sourceColumnIndex:=r("SourceColumnIndex").ToString(),
                        sourceHeaderName:=r("SourceHeaderName").ToString(),
                        sourceSheetName:=r("SourceSheetName").ToString(),
                        sourceCellAddress:=r("SourceCellAddress").ToString(),
                        rowSelectorType:=r("RowSelectorType").ToString(),
                        rowSelectorValue:=r("RowSelectorValue").ToString(),
                        dataType:=r("DataType").ToString(),
                        formatPattern:=r("FormatPattern").ToString(),
                        userId:=userId,
                        pgId:="CreateMappingSetting(Detail)"
                    )
                Next

                If newProfileId <= 0 Then
                    Throw New ApplicationException("入力情報の登録に失敗しました。")
                End If

                ' 完了メッセージ
                lblResult.Text = "入力情報を登録しました。"

            Catch ex As ApplicationException
                lblError.Text = Server.HtmlEncode(ex.Message)

            Catch ex As Oracle.ManagedDataAccess.Client.OracleException
                lblError.Text = "DBエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)

            Catch ex As Exception
                lblError.Text = "予期しないエラーが発生しました。詳細：" & Server.HtmlEncode(ex.Message)
            End Try
        End Sub
#End Region

#Region "汎用ヘルパ"
        ' 編集行の DDL に項目を流し込み、現在値があれば選択
        Private Sub BindAndSelect(Of TKey)(row As GridViewRow,
                                   ddlId As String,
                                   map As IReadOnlyDictionary(Of TKey, String),
                                   currentObj As Object,
                                   Optional orderByLabel As Boolean = True,
                                   Optional orderByKey As Boolean = False)
            Dim ddl = TryCast(row.FindControl(ddlId), DropDownList)
            If ddl Is Nothing Then Return

            ddl.Items.Clear()
            Dim items = map.ToList()

            If orderByKey Then
                items.Sort(Function(a, b) Comparer(Of TKey).Default.Compare(a.Key, b.Key))
            ElseIf orderByLabel Then
                items.Sort(Function(a, b) StringComparer.CurrentCulture.Compare(a.Value, b.Value))
            End If

            For Each kv In items
                ddl.Items.Add(New ListItem(kv.Value, kv.Key?.ToString()))
            Next

            Dim cur = If(currentObj, "").ToString().Trim()
            Dim found = ddl.Items.FindByValue(cur)
            If found Is Nothing AndAlso GetType(TKey) Is GetType(Integer) Then
                ' DataBinder.Eval が "2" などの文字列を返した場合にも対応
                Dim curInt As Integer
                If Integer.TryParse(cur, curInt) Then
                    found = ddl.Items.FindByValue(curInt.ToString())
                End If
            End If

            ddl.ClearSelection()
            If found IsNot Nothing Then
                found.Selected = True
            ElseIf ddl.Items.Count > 0 Then
                ddl.SelectedIndex = 0
            End If
        End Sub

        ' 文字列→Int32。空/不正は DBNull
        Private Function ToIntOrDBNull(text As String) As Object
            If String.IsNullOrWhiteSpace(text) Then Return DBNull.Value
            Dim i As Integer
            If Integer.TryParse(text.Trim(), i) Then
                Return i
            End If
            Return DBNull.Value
        End Function

        ' DDL の SelectedValue を取得（NULL 安全）
        Private Function GetDDL(row As GridViewRow, id As String) As String
            Dim ddl = TryCast(row.FindControl(id), DropDownList)
            Return If(ddl Is Nothing, String.Empty, ddl.SelectedValue)
        End Function

        ' e.NewValues の NULL 安全取り出し
        Private Function SafeGet(dic As IOrderedDictionary, key As String) As String
            Dim v = dic(key)
            Return If(v Is Nothing, String.Empty, v.ToString())
        End Function
#End Region

    End Class
End Namespace