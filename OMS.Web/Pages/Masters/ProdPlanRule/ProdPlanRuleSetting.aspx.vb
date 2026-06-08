Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.ProdPlanRule
    Public Class ProdPlanRuleSetting
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        Private Const SESSION_KEY As String = "SplitCaseDT"
        Private Const DUMMY_ROW_FLAG As String = "__IsDummy__"

        ' id1 = ProdPlanRuleId（NUMBER(10,0) → Long）※未指定は 0
        Private ReadOnly Property ProdPlanRuleId As Long
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
                BindDropDown(ddlSplitFlag, SplitFlagMap, sortByLabel:=False)
                BindDropDown(ddlSplitCaseFlag, SplitCaseFlagMap, sortByLabel:=False)
                BindDropDown(ddlSplitMethodType, SplitMethodTypeMap, sortByLabel:=False)
                BindDropDown(ddlSplitRationType, SplitRationTypeMap, sortByLabel:=False)
                BindDropDown(ddlCarryToType, CarryToTypeMap, sortByLabel:=False)
                BindDropDown(ddlSplitStartType, SplitStartTypeMap, sortByLabel:=False)
                BindDropDown(ddlActiveFlag, ActiveFlagMap, sortByLabel:=False)
                LoadProdPlanRuleHeader(ProdPlanRuleId)
                EnsureSessionTable()
                BindSplitCaseGrid()
            End If
        End Sub
        Protected Overrides Sub OnPreRender(e As EventArgs)
            MyBase.OnPreRender(e)
            ' lblError が DOM に出るタイミングで確実に ClientID が決まっている
            btnSaveProdPlanRuleSetting.Attributes("data-error-label-id") = lblError.ClientID
        End Sub
#End Region

#Region "生産計画マスタデータ"
        ' ProdPlanRuleId でヘッダー情報を読み、画面の各コントロールに流し込む
        Private Sub LoadProdPlanRuleHeader(prodPlanRuleId As Long)

            If prodPlanRuleId <= 0 Then
                SetHeaderControls(Nothing)
                ' 必要ならメッセージ
                ' lblError.Text = "キーが不正です。"
                Exit Sub
            End If

            Dim repo As New ProdPlanRuleRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetProdPlanRule(prodPlanRuleId)

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
            txtCustomerName.Text = If(GetStr(r, "CustomerName"), "")
            txtProfitCenter.Text = If(GetStr(r, "ProfitCenter"), "")
            txtCustomerUnitName.Text = If(GetStr(r, "CustomerUnitName"), "")

            'ddlSplitFlag.SelectedValue = Lookup(SplitFlagMap, GetStr(r, "SplitFlag"))
            Dim strSplitFlag As String = GetStr(r, "SplitFlag")
            ddlSplitFlag.ClearSelection()
            If Not String.IsNullOrEmpty(strSplitFlag) Then
                Dim item = ddlSplitFlag.Items.FindByValue(strSplitFlag)
                If item IsNot Nothing Then item.Selected = True
            End If

            'ddlSplitCaseFlag.SelectedValue = Lookup(SplitCaseFlagMap, GetStr(r, "SplitCaseFlag"))
            Dim strSplitCaseFlag As String = GetStr(r, "SplitCaseFlag")
            ddlSplitCaseFlag.ClearSelection()
            If Not String.IsNullOrEmpty(strSplitCaseFlag) Then
                Dim item = ddlSplitCaseFlag.Items.FindByValue(strSplitCaseFlag)
                If item IsNot Nothing Then item.Selected = True
            End If

            'ddlSplitMethodType.SelectedValue = Lookup(SplitMethodTypeMap, GetStr(r, "SplitMethodType"))
            Dim strSplitMethodType As String = GetStr(r, "SplitMethodType")
            ddlSplitMethodType.ClearSelection()
            If Not String.IsNullOrEmpty(strSplitMethodType) Then
                Dim item = ddlSplitMethodType.Items.FindByValue(strSplitMethodType)
                If item IsNot Nothing Then item.Selected = True
            End If

            'ddlSplitRationType.SelectedValue = Lookup(SplitRationTypeMap, GetStr(r, "SplitRationType"))
            Dim strSplitRationType As String = GetStr(r, "SplitRationType")
            ddlSplitRationType.ClearSelection()
            If Not String.IsNullOrEmpty(strSplitRationType) Then
                Dim item = ddlSplitRationType.Items.FindByValue(strSplitRationType)
                If item IsNot Nothing Then item.Selected = True
            End If

            txtSplitRoudingUnit.Text = If(GetStr(r, "RoudhingUnit"), "")

            'ddlCarryToType.SelectedValue = Lookup(CarryToTypeMap, GetStr(r, "CarryToType"))
            Dim strCarryToType As String = GetStr(r, "CarryToType")
            ddlCarryToType.ClearSelection()
            If Not String.IsNullOrEmpty(strCarryToType) Then
                Dim item = ddlCarryToType.Items.FindByValue(strCarryToType)
                If item IsNot Nothing Then item.Selected = True
            End If

            'ddlSplitStartType.SelectedValue = Lookup(SplitStartTypeMap, GetStr(r, "SplitStartType"))
            Dim strSplitStartType As String = GetStr(r, "SplitStartType")
            ddlSplitStartType.ClearSelection()
            If Not String.IsNullOrEmpty(strSplitStartType) Then
                Dim item = ddlSplitStartType.Items.FindByValue(strSplitStartType)
                If item IsNot Nothing Then item.Selected = True
            End If

            Dim strActiveFlag As String = GetStr(r, "ActiveFlag")
            ddlActiveFlag.ClearSelection()
            If Not String.IsNullOrEmpty(strActiveFlag) Then
                Dim item = ddlActiveFlag.Items.FindByValue(strActiveFlag)
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

#Region "Session管理 / 読み込み / 緩和 / バインド"
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

            dt = LoadProdPlanRuleTableFromDb()

            ' 画面内ユニークキー（TempId）を付与
            If Not dt.Columns.Contains("TempId") Then
                dt.Columns.Add("TempId", GetType(String))
            End If
            For Each r As DataRow In dt.Rows
                r("TempId") = "DB:" & r("SplitCaseId").ToString()
            Next

            ' 監査系を画面用に緩和（NULL許容など）
            RelaxDataTableForScreen(dt)

            ' 以降の差分検出の起点
            dt.AcceptChanges()
            SetSessionTable(dt)
        End Sub

        ' DB から一覧を読み込む（緩和や AcceptChanges は呼び出し元で実施）
        Private Function LoadProdPlanRuleTableFromDb() As DataTable
            Dim repo As New ProdPlanRuleRepository(Utils.GetConnectionString())
            Dim dt As DataTable = repo.GetSplitCaseList(prodPlanRuleId:=ProdPlanRuleId, activeFlag:="Y")
            If Not dt.Columns.Contains("TempId") Then dt.Columns.Add("TempId", GetType(String))
            For Each r As DataRow In dt.Rows
                r("TempId") = "DB:" & r("SplitCaseId").ToString()
            Next
            Return dt
        End Function

        ' 画面内の一時管理のために列制約を緩める
        Private Sub RelaxDataTableForScreen(dt As DataTable)
            Dim nullableCols = New String() {
                "SplitCaseId",
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
        Private Sub BindSplitCaseGrid()
            EnsureSessionTable()

            Dim dt As DataTable = GetSessionTable()
            If dt Is Nothing Then
                dt = New DataTable()
            End If
            If Not dt.Columns.Contains(DUMMY_ROW_FLAG) Then
                dt.Columns.Add(DUMMY_ROW_FLAG, GetType(Boolean))
            End If

            Dim currentCount As Integer = dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length
            If currentCount = 0 Then
                Dim dummy As DataRow = dt.NewRow()
                For Each col As DataColumn In dt.Columns
                    Select Case col.ColumnName
                        Case "TempId"
                            dummy(col) = "DUMMY"
                        Case "SplitCaseId"
                            dummy(col) = DBNull.Value
                        Case DUMMY_ROW_FLAG
                            dummy(col) = True
                        Case "SplitMethodType"
                            dummy(col) = GetSplitMethodTypeDefault()
                        Case Else
                            If col.AllowDBNull Then
                                dummy(col) = DBNull.Value
                            Else
                                Dim v = GetDefaultSafe(col)
                                dummy(col) = v
                            End If
                    End Select
                Next

                dt.Rows.Add(dummy)
            End If

            Dim dv As New DataView(dt, Nothing, Nothing, DataViewRowState.CurrentRows)
            gvSplitCaseList.DataSource = dv
            gvSplitCaseList.DataBind()
        End Sub

        ' SplitMethodType の既定値
        Private Function GetSplitMethodTypeDefault() As Object
            If SplitMethodTypeMap IsNot Nothing AndAlso SplitMethodTypeMap.Count > 0 Then
                Return SplitMethodTypeMap.Keys.First()
            End If
            Return 0
        End Function

        'グリッドを初期化する（フォームを開く時に使用）
        Private Sub ResetGrid()
            Dim dt = GetSessionTable()
            If dt Is Nothing Then Exit Sub
            Session.Remove(SESSION_KEY)
        End Sub
#End Region

#Region "GridView イベント（明細）"
        ' 編集行/フッターの DDL 初期化 + 0件時ダミー行の非表示
        Protected Sub gvSplitCaseList_RowDataBound(sender As Object, e As GridViewRowEventArgs) _
            Handles gvSplitCaseList.RowDataBound

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
                    BindAndSelect(e.Row, "ddlQtyConditionType", QtyConditionTypeMap, DataBinder.Eval(e.Row.DataItem, "QtyConditionType"))
                    BindAndSelect(e.Row, "ddlSplitMethodType", SplitMethodTypeMap, DataBinder.Eval(e.Row.DataItem, "SplitMethodType"))
                    BindAndSelect(e.Row, "ddlActiveFlag", ActiveFlagMap, DataBinder.Eval(e.Row.DataItem, "ActiveFlag"))
                End If

            ElseIf e.Row.RowType = DataControlRowType.Footer Then
                ' フッター行（新規追加用）
                BindDropDown(TryCast(e.Row.FindControl("ddlQtyConditionType_F"), DropDownList), QtyConditionTypeMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlSplitMethodType_F"), DropDownList), SplitMethodTypeMap, sortByLabel:=False)
                BindDropDown(TryCast(e.Row.FindControl("ddlActiveFlag_F"), DropDownList), ActiveFlagMap, sortByLabel:=False)
            End If
        End Sub

        ' 編集開始
        Protected Sub gvSplitCaseList_RowEditing(sender As Object, e As GridViewEditEventArgs)
            gvSplitCaseList.EditIndex = e.NewEditIndex
            BindSplitCaseGrid()
        End Sub

        ' 編集キャンセル
        Protected Sub gvSplitCaseList_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)
            e.Cancel = True
            gvSplitCaseList.EditIndex = -1
            BindSplitCaseGrid()
        End Sub

        ' 編集確定（画面内 DataTable を更新）
        Protected Sub gvSplitCaseList_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)
            Dim dt As DataTable = GetSessionTable()
            Dim row As GridViewRow = gvSplitCaseList.Rows(e.RowIndex)

            Dim tempId As String = gvSplitCaseList.DataKeys(e.RowIndex).Values("TempId").ToString()

            Dim dr As DataRow = dt.AsEnumerable().
                First(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            ' ドロップダウンリスト：FindControl
            dr("QtyConditionType") = GetDDL(row, "ddlQtyConditionType")
            dr("SplitMethodType") = GetDDL(row, "ddlSplitMethodType")
            dr("ActiveFlag") = GetDDL(row, "ddlActiveFlag")

            ' テキスト：e.NewValues
            dr("Qty") = ToIntOrDBNull(SafeGet(e.NewValues, "Qty"))

            gvSplitCaseList.EditIndex = -1
            BindSplitCaseGrid()
            e.Cancel = True
        End Sub

        ' 追加（フッター行から）
        Protected Sub gvSplitCaseList_RowCommand(sender As Object, e As GridViewCommandEventArgs) _
            Handles gvSplitCaseList.RowCommand

            If Not String.Equals(e.CommandName, "Insert", StringComparison.OrdinalIgnoreCase) Then Return

            Dim dt As DataTable = GetSessionTable()
            Dim f As GridViewRow = gvSplitCaseList.FooterRow

            ' フッター入力値取得
            Dim txtQtyF = TryCast(f.FindControl("txtQty_F"), TextBox)
            Dim ddlQtyConditionTypeF = TryCast(f.FindControl("ddlQtyConditionType_F"), DropDownList)
            Dim ddlSplitMethodTypeF = TryCast(f.FindControl("ddlSplitMethodType_F"), DropDownList)
            Dim ddlActiveFlagF = TryCast(f.FindControl("ddlActiveFlag_F"), DropDownList)

            Dim qtyText As String = If(txtQtyF?.Text, String.Empty).Trim()
            Dim qty As Integer
            Dim hasQty As Boolean = Integer.TryParse(qtyText, qty)

            Dim qtyCond As String = If(ddlQtyConditionTypeF?.SelectedValue, String.Empty).Trim()
            Dim splitMethod As String = If(ddlSplitMethodTypeF?.SelectedValue, String.Empty).Trim()
            Dim activeFlag As String = If(ddlActiveFlagF?.SelectedValue, String.Empty).Trim()
            Dim errors As New List(Of String)

            If Not hasQty Then
                errors.Add("数量（Qty）は数値で入力してください。")
            End If

            If errors.Count > 0 Then
                lblDetailError.Text = Server.HtmlEncode(String.Join(" / ", errors))
                lblResult.Text = String.Empty
                If Not hasQty Then
                    txtQtyF?.Focus()
                End If
                Return
            End If

            Dim r As DataRow = dt.NewRow()
            r("TempId") = "NEW:" & Guid.NewGuid().ToString("N")
            r("ProdPlanRuleId") = ProdPlanRuleId

            r("Qty") = qty
            r("QtyConditionType") = qtyCond
            r("SplitMethodType") = splitMethod
            r("ActiveFlag") = activeFlag

            If dt.Columns.Contains(DUMMY_ROW_FLAG) Then r(DUMMY_ROW_FLAG) = False

            dt.Rows.Add(r)
            BindSplitCaseGrid()

            lblDetailError.Text = ""
            lblDetailResult.Text = ""
        End Sub

        ' 削除（画面内：既存行は Delete マーク、新規行は Remove）
        Protected Sub gvSplitCaseList_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
            Dim dt As DataTable = GetSessionTable()
            Dim keys = gvSplitCaseList.DataKeys(e.RowIndex)
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

            BindSplitCaseGrid()
            e.Cancel = True
        End Sub
#End Region

#Region "保存（DBへ一括反映）"
        Protected Sub btnSaveProdPlanRuleSetting_Click(sender As Object, e As EventArgs)

            ' セッション確認
            Dim userId As String = PageHelpers.GetUserId(Me.Page)
            If String.IsNullOrEmpty(userId) Then
                Exit Sub
            End If

            ' 以降は認証済み前提
            If userId.Length > 9 Then
                userId = userId.Substring(0, 9)
            End If

            Dim dt As DataTable = GetSessionTable()
            Dim repo As New ProdPlanRuleRepository(Utils.GetConnectionString())
            Dim pgId As String = "SaveProdPlanRuleSetting"
            Dim pgIdDetail As String = String.Empty

            ' 入力取得
            Dim splitFlag As String = (If(ddlSplitFlag.SelectedValue, "")).Trim()
            Dim splitCaseFlag As String = (If(ddlSplitCaseFlag.SelectedValue, "")).Trim()
            Dim splitMethodType As String = (If(ddlSplitMethodType.SelectedValue, "")).Trim()
            Dim splitRationType As String = (If(ddlSplitRationType.SelectedValue, "")).Trim()
            Dim splitRoudingUnit As String = (If(txtSplitRoudingUnit.Text, "")).Trim()
            Dim carryToType As String = (If(ddlCarryToType.SelectedValue, "")).Trim()
            Dim splitStartType As String = (If(ddlSplitStartType.SelectedValue, "")).Trim()
            Dim activeFlag As String = (If(ddlActiveFlag.SelectedValue, "")).Trim()

            Try
                '============================================================================================
                ' 生産計画条件マスタ
                '============================================================================================
                pgIdDetail = "Update"
                Dim affected As Integer = repo.UpdateProdPlanRule(
                    prodPlanRuleId:=ProdPlanRuleId.ToString(),
                    splitFlag:=splitFlag,
                    splitCaseFlag:=splitCaseFlag,
                    splitMethodType:=splitMethodType,
                    splitRationType:=splitRationType,
                    splitRoudingUnit:=splitRoudingUnit,
                    carryToType:=carryToType,
                    splitStartType:=splitStartType,
                    activeFlag:=activeFlag,
                    userId:=userId,
                    pgId:=pgId & "(" & pgIdDetail & ")"
                )


                '============================================================================================
                ' 分割パターンマスタ
                '============================================================================================
                pgIdDetail = "Insert"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Added)
                    ' ダミーデータ判定（TempId = "DUMMY" の場合はスキップ）
                    If Not IsDBNull(r("TempId")) AndAlso r("TempId").ToString().Trim().ToUpper() = "DUMMY" Then
                        Continue For
                    End If

                    repo.InsertSplitCase(
                        prodPlanRuleId:=r("ProdPlanRuleId").ToString(),
                        qty:=r("Qty").ToString(),
                        qtyConditionType:=r("QtyConditionType").ToString(),
                        splitMethodType:=r("SplitMethodType").ToString(),
                        userId:=userId,
                        pgId:=pgId & "(" & pgIdDetail & ")"
                    )
                Next

                pgIdDetail = "Update"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent)
                    Dim tempId As String = r("TempId").ToString()
                    If tempId.StartsWith("DB:", StringComparison.OrdinalIgnoreCase) Then
                        Dim splitCaseId As String = r("SplitCaseId").ToString() ' ← PKはこれを使う想定
                        repo.UpdateSplitCase(
                            splitCaseId:=splitCaseId,
                            prodPlanRuleId:=r("ProdPlanRuleId").ToString(),
                            qty:=r("Qty").ToString(),
                            qtyConditionType:=r("QtyConditionType").ToString(),
                            splitMethodType:=r("SplitMethodType").ToString(),
                            activeFlag:=r("ActiveFlag").ToString(),
                            userId:=userId,
                            pgId:=pgId & "(" & pgIdDetail & ")"
                        )
                    End If
                Next

                pgIdDetail = "Delete"
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Deleted)
                    Dim splitCaseId As String = r("SplitCaseId", DataRowVersion.Original).ToString()
                    If Not String.IsNullOrEmpty(splitCaseId) Then
                        repo.DeleteSplitCase(splitCaseId:=splitCaseId)
                    End If
                Next

                LoadProdPlanRuleHeader(ProdPlanRuleId)
                Dim fresh = LoadProdPlanRuleTableFromDb()
                RelaxDataTableForScreen(fresh)
                fresh.AcceptChanges()
                SetSessionTable(fresh)
                BindSplitCaseGrid()

                lblResult.Text = "編集内容を保存しました。"
                lblError.Text = ""
            Catch ex As Exception
                lblError.Text = "保存時にエラーが発生しました: " & Server.HtmlEncode(ex.Message)
                lblResult.Text = ""
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
            If String.IsNullOrWhiteSpace(text) Then
                Return DBNull.Value
            End If
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

        ' 型の既定値を安全に返す（NULL不可列に Nothing を返さない）
        Private Function GetDefaultSafe(col As DataColumn) As Object
            Dim t As Type = col.DataType

            ' 列名に依存する既定値を優先して返す例（必要なら増やす）
            Select Case col.ColumnName
                'Case "ProratedType" : Return 1 ' 例：既定を1にしたい場合など
            End Select

            If t Is GetType(String) Then Return String.Empty
            If t Is GetType(Integer) OrElse t Is GetType(Int32) Then Return 0
            If t Is GetType(Long) Then Return 0L
            If t Is GetType(Decimal) Then Return 0D
            If t Is GetType(Double) Then Return 0.0
            If t Is GetType(Boolean) Then Return False
            If t Is GetType(DateTime) Then Return DateTime.MinValue
            If t.IsEnum Then
                Try
                    Return System.Enum.ToObject(t, 0)
                Catch
                    Return Activator.CreateInstance(t)
                End Try
            End If
            Try
                Return Activator.CreateInstance(t)
            Catch
                If col.AllowDBNull Then Return DBNull.Value
                Return String.Empty
            End Try
        End Function
#End Region

    End Class
End Namespace