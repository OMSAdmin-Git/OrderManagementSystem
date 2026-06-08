Imports System.Collections.Specialized
Imports System.Data
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic.ApplicationServices
Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.Mapping
    Public Class MappingSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        Private Const SESSION_KEY As String = "FieldMappingDT"
        Private Const DUMMY_ROW_FLAG As String = "__IsDummy__"

        ' id1 = ProfileId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property ProfileId As Long
            Get
                Dim s As String = If(Request.QueryString("id1"), "").Trim()
                Dim v As Long = 0
                Long.TryParse(s, v)
                Return v
            End Get
        End Property
#End Region

#Region "ページ ライフサイクル"
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                ResetGrid()
                PageHelpers.SetUserName(Me, lblUser)
                BindDropDown(ddlActiveFlag, OMS.Common.Constants.ActiveFlagMap, sortByLabel:=False)
                LoadProfileHeader(ProfileId)
                EnsureSessionTable()
                BindFieldMappingGrid()
            End If
        End Sub

        Protected Overrides Sub OnPreRender(e As EventArgs)
            MyBase.OnPreRender(e)
            ' lblError が DOM に出るタイミングで確実に ClientID が決まっている
            btnSaveMappingSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

#Region "マッピングマスタデータ"
        ' ProfileId でヘッダー情報を読み、画面の各コントロールに流し込む
        Private Sub LoadProfileHeader(profileId As Long)

            If profileId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New MappingRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetMappingProfile(profileId)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                SetHeaderControls(Nothing)
                lblError.Text = "指定されたプロファイルが見つかりません。"
                Exit Sub
            End If

            SetHeaderControls(dt.Rows(0))
        End Sub

        ' ヘッダー用コントロールに値をセット（Nothing なら空クリア）
        Private Sub SetHeaderControls(r As DataRow)

            txtCustomerCode.Text = If(GetStr(r, "CustomerCode"), "")
            'txtCustomerName.Text = If(GetStr(r, "CustomerName"), "")
            txtProfitCenter.Text = If(GetStr(r, "ProfitCenter"), "")
            txtCustomerUnitName.Text = If(GetStr(r, "CustomerUnitName"), "")
            txtFolderType.Text = Lookup(FolderTypeMap, GetStr(r, "FolderType"))
            txtVersion.Text = If(GetStr(r, "Version"), "")
            txtHeaderRowIndex.Text = If(GetStr(r, "HeaderRowIndex"), "")
            txtDataStartRowIndex.Text = If(GetStr(r, "DataStartRowIndex"), "")
            txtDefaultSheetName.Text = If(GetStr(r, "DefaultSheetName"), "")
            Dim flag As String = GetStr(r, "ActiveFlag")
            ddlActiveFlag.ClearSelection()
            If Not String.IsNullOrEmpty(flag) Then
                Dim item = ddlActiveFlag.Items.FindByValue(flag)
                If item IsNot Nothing Then item.Selected = True
            End If
            txtUpdatedAt.Text = If(GetStr(r, "UpdatedAt"), "")
            txtUpdatedUserName.Text = If(GetStr(r, "UpdatedUserId"), "")
        End Sub

        ' DataRowの項目を文字列で安全取得
        Private Function GetStr(r As DataRow, columnName As String) As String
            If r Is Nothing Then Return Nothing
            If Not r.Table.Columns.Contains(columnName) Then Return Nothing
            Dim v = r(columnName)
            If v Is DBNull.Value OrElse v Is Nothing Then Return Nothing
            Return v.ToString()
        End Function
#End Region

#Region "Session管理 / 読み込み / 緩和 / バインド（明細）"
        Private Function GetSessionTable() As DataTable
            Return TryCast(Session(SESSION_KEY), DataTable)
        End Function

        Private Sub SetSessionTable(dt As DataTable)
            Session(SESSION_KEY) = dt
        End Sub

        ' 初回のみ DB→Session へ。以降は Session 上のテーブルを使う
        Private Sub EnsureSessionTable()
            Dim dt = GetSessionTable()
            If dt IsNot Nothing Then Exit Sub

            dt = LoadFieldMappingTableFromDb()

            ' 画面内ユニークキー（TempId）を付与
            If Not dt.Columns.Contains("TempId") Then
                dt.Columns.Add("TempId", GetType(String))
            End If
            For Each r As DataRow In dt.Rows
                r("TempId") = "DB:" & r("FieldMappingId").ToString()
            Next

            ' 監査系を画面用に緩和（NULL許容など）
            RelaxDataTableForScreen(dt)

            ' 以降の差分検出の起点
            dt.AcceptChanges()
            SetSessionTable(dt)
        End Sub

        ' DB から一覧を読み込む（緩和や AcceptChanges は呼び出し元で実施）
        Private Function LoadFieldMappingTableFromDb() As DataTable
            Dim repo As New MappingRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetFieldMappingList(profileId:=ProfileId, activeFlag:="Y")
            If Not dt.Columns.Contains("TempId") Then dt.Columns.Add("TempId", GetType(String))
            For Each r As DataRow In dt.Rows
                r("TempId") = "DB:" & r("FieldMappingId").ToString()
            Next
            Return dt
        End Function

        ' 画面内の一時管理のために列制約を緩める
        Private Sub RelaxDataTableForScreen(dt As DataTable)
            Dim nullableCols = New String() {
                "FieldMappingId",
                "CreatedAt", "CreatedUserId", "CreatedPgId",
                "UpdatedAt", "UpdatedUserId", "UpdatedPgId"
            }
            For Each name In nullableCols
                If dt.Columns.Contains(name) Then
                    Dim c = dt.Columns(name)
                    c.AllowDBNull = True
                    c.DefaultValue = DBNull.Value
                End If
            Next

            If dt.Columns.Contains("ActiveFlag") Then
                dt.Columns("ActiveFlag").DefaultValue = "Y"
            End If
        End Sub

        ' この画面の明細グリッドだけをバインド（0件時はダミー行を注入してFooterを出す）
        Private Sub BindFieldMappingGrid()
            EnsureSessionTable()

            Dim dt As DataTable = GetSessionTable()
            If dt Is Nothing Then
                dt = New DataTable()
            End If

            ' ダミー判定用の内部列（なければ追加）
            If Not dt.Columns.Contains(DUMMY_ROW_FLAG) Then
                dt.Columns.Add(DUMMY_ROW_FLAG, GetType(Boolean))
            End If

            ' 0件（CurrentRows）であれば、Footerを出すためのダミー行を1件追加
            Dim currentCount As Integer = dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length
            If currentCount = 0 Then
                Dim dummy As DataRow = dt.NewRow()

                ' 必須列（DataKeyNames 等）の最低限の整合性を保つ
                For Each col As DataColumn In dt.Columns
                    Select Case col.ColumnName
                        Case "TempId"
                            dummy(col) = "DUMMY"
                        Case "FieldMappingId"
                            dummy(col) = DBNull.Value
                        Case DUMMY_ROW_FLAG
                            dummy(col) = True
                        Case Else
                            If col.AllowDBNull Then
                                dummy(col) = DBNull.Value
                            Else
                                dummy(col) = GetDefault(col.DataType)
                            End If
                    End Select
                Next

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

#Region "GridView イベント（明細）"
        ' 編集行/フッターの DDL 初期化 + 0件時ダミー行の非表示
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

                ' 編集行のみ（ItemTemplate はそのまま表示）
                If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                    BindAndSelect(e.Row, "ddlFormatType", FormatTypeMap, DataBinder.Eval(e.Row.DataItem, "FormatType"))
                    BindAndSelect(e.Row, "ddlTargetField", TargetFieldMap, DataBinder.Eval(e.Row.DataItem, "TargetField"))
                    BindAndSelect(e.Row, "ddlRowSelectorType", RowSelectorTypeMap, DataBinder.Eval(e.Row.DataItem, "RowSelectorType"))
                    BindAndSelect(e.Row, "ddlDataType", DataTypeMap, DataBinder.Eval(e.Row.DataItem, "DataType"))
                    BindAndSelect(e.Row, "ddlActiveFlag", ActiveFlagMap, DataBinder.Eval(e.Row.DataItem, "ActiveFlag"))
                End If

            ElseIf e.Row.RowType = DataControlRowType.Footer Then
                ' フッター行（新規追加用）
                BindDropDown(TryCast(e.Row.FindControl("ddlFormatType_F"), DropDownList), FormatTypeMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlTargetField_F"), DropDownList), TargetFieldMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlRowSelectorType_F"), DropDownList), RowSelectorTypeMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlDataType_F"), DropDownList), DataTypeMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlActiveFlag_F"), DropDownList), ActiveFlagMap, sortByLabel:=False)
            End If
        End Sub

        ' 編集開始
        Protected Sub gvFieldMappingList_RowEditing(sender As Object, e As GridViewEditEventArgs)
            gvFieldMappingList.EditIndex = e.NewEditIndex
            BindFieldMappingGrid()
        End Sub

        ' 編集キャンセル
        Protected Sub gvFieldMappingList_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)
            e.Cancel = True
            gvFieldMappingList.EditIndex = -1
            BindFieldMappingGrid()
        End Sub

        ' 編集確定（画面内 DataTable を更新）
        Protected Sub gvFieldMappingList_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)
            Dim dt As DataTable = GetSessionTable()
            Dim row As GridViewRow = gvFieldMappingList.Rows(e.RowIndex)

            Dim tempId As String = gvFieldMappingList.DataKeys(e.RowIndex).Values("TempId").ToString()

            Dim dr As DataRow = dt.AsEnumerable().
                First(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            ' DDL は FindControl、テキストは e.NewValues
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
            r("ProfileId") = ProfileId

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

            ' ダミー列があれば False（実データ）
            If dt.Columns.Contains(DUMMY_ROW_FLAG) Then r(DUMMY_ROW_FLAG) = False

            dt.Rows.Add(r) ' RowState=Added
            BindFieldMappingGrid()
        End Sub

        ' 削除（画面内：既存行は Delete マーク、新規行は Remove）
        Protected Sub gvFieldMappingList_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
            Dim dt As DataTable = GetSessionTable()
            Dim keys = gvFieldMappingList.DataKeys(e.RowIndex)
            Dim tempId As String = keys("TempId").ToString()

            Dim dr As DataRow = dt.AsEnumerable().
                FirstOrDefault(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            If dr IsNot Nothing Then
                If tempId.StartsWith("NEW:", StringComparison.OrdinalIgnoreCase) Then
                    dt.Rows.Remove(dr)
                Else
                    dr.Delete()
                End If
            End If

            BindFieldMappingGrid()
            e.Cancel = True
        End Sub
#End Region

#Region "保存（DBへ一括反映）"
        Protected Sub btnSaveMappingSetting_Click(sender As Object, e As EventArgs)

            lblResult.Text = ""
            lblError.Text = ""

            ' セッション確認
            Dim userId As String = PageHelpers.GetUserId(Me.Page)
            If String.IsNullOrEmpty(userId) Then Exit Sub

            ' 以降は認証済み前提（9桁制限）
            If userId.Length > 9 Then
                userId = userId.Substring(0, 9)
            End If

            Dim dt As DataTable = GetSessionTable()
            Dim repo As New MappingRepository(Utils.GetConnectionString())
            Dim pgId As String = "SaveMappingSetting"
            Dim pgIdDetail As String = String.Empty

            ' 入力取得
            Dim version As String = (If(txtVersion.Text, "")).Trim()
            Dim headerRowIndex As String = (If(txtHeaderRowIndex.Text, "")).Trim()
            Dim dataStartRowIndex As String = (If(txtDataStartRowIndex.Text, "")).Trim()
            Dim defaultSheetName As String = (If(txtDefaultSheetName.Text, "")).Trim()
            Dim activeFlag As String = (If(ddlActiveFlag.SelectedValue, "")).Trim()

            Try
                '============================================================================================
                ' マッピングマスタ
                '============================================================================================
                ' UPDATE
                pgIdDetail = "Update"
                Dim affected As Integer = repo.UpdateMappingProfile(
                    profileId:=ProfileId.ToString(),
                    version:=version,
                    headerRowIndex:=headerRowIndex,
                    dataStartRowindex:=dataStartRowIndex,
                    defaultSheetName:=defaultSheetName,
                    activeFlag:=activeFlag,
                    userId:=userId,
                    pgId:=pgId & "(" & pgIdDetail & ")"
                )


                '============================================================================================
                ' マッピング明細マスタ
                '============================================================================================
                pgIdDetail = "Insert"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Added)
                    ' ダミーデータ判定（TempId = "DUMMY" の場合はスキップ）
                    If Not IsDBNull(r("TempId")) AndAlso r("TempId").ToString().Trim().ToUpper() = "DUMMY" Then
                        Continue For
                    End If

                    repo.InsertFieldMapping(
                        profileId:=r("ProfileId").ToString(),
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
                        pgId:=pgId & "(" & pgIdDetail & ")"
                    )
                Next

                pgIdDetail = "Update"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent)
                    Dim tempId As String = r("TempId").ToString()
                    If tempId.StartsWith("DB:", StringComparison.OrdinalIgnoreCase) Then
                        Dim id As String = r("FieldMappingId").ToString()
                        repo.UpdateFieldMapping(
                            fieldMappingId:=id,
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
                            activeFlag:=r("ActiveFlag").ToString(),
                            userId:=userId,
                            pgId:=pgId & "(" & pgIdDetail & ")"
                        )
                    End If
                Next

                pgIdDetail = "Delete"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Deleted)
                    Dim id As String = r("FieldMappingId", DataRowVersion.Original).ToString()
                    If Not String.IsNullOrEmpty(id) Then
                        repo.DeleteFieldMapping(fieldMappingId:=id)
                    End If
                Next

                ' 最新読込 → 画面リフレッシュ
                LoadProfileHeader(ProfileId)
                Dim fresh = LoadFieldMappingTableFromDb()
                RelaxDataTableForScreen(fresh)
                fresh.AcceptChanges()
                SetSessionTable(fresh)
                BindFieldMappingGrid()

                lblResult.Text = "編集内容を保存しました。"
            Catch ex As Exception
                lblError.Text = "保存時にエラーが発生しました: " & Server.HtmlEncode(ex.Message)
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

        ' 型の既定値（NULL 不可列のダミー埋め用）
        Private Function GetDefault(t As Type) As Object
            If t Is GetType(String) Then Return String.Empty
            If t Is GetType(Integer) Then Return 0
            If t Is GetType(Long) Then Return 0L
            If t Is GetType(Decimal) Then Return 0D
            If t Is GetType(Double) Then Return 0.0
            If t Is GetType(Boolean) Then Return False
            If t Is GetType(DateTime) Then Return DateTime.MinValue
            Return Nothing
        End Function
#End Region

    End Class
End Namespace