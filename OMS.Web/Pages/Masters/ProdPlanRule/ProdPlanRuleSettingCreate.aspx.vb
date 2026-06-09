Imports OMS.Common
Imports OMS.Data

Namespace Pages.Masters.ProdPlanRule
    Public Class ProdPlanRuleSettingCreate
        Inherits System.Web.UI.Page

#Region "定数・フィールド"
        Private Const SESSION_KEY As String = "SplitCaseDT"
        Private Const DUMMY_ROW_FLAG As String = "__IsDummy__"
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
                LoadSearchConditionLists()
                EnsureSessionTable()
                BindSplitCaseGrid()
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

            lstCustomerCode.InnerHtml = BuildOptions(customerCodeList)
            lstProfitCenter.InnerHtml = BuildOptions(profitCenterList)
            lstCustomerUnitName.InnerHtml = BuildOptions(customerUnitNameList)
        End Sub
#End Region

#Region "Session管理 / バインド（新規専用）"
        ' 新規専用：画面で使う空スキーマを作成（DBからは読み込まない）
        Private Function BuildEmptySplitCaseTableSchema() As DataTable
            Dim dt As New DataTable("SplitCase")

            ' 画面内ユニークキー
            dt.Columns.Add("TempId", GetType(String))

            ' DB主キー（新規は DBNull でOK）
            dt.Columns.Add("SplitCaseId", GetType(Long))

            ' 保存時に確定するため、初期は DBNull
            dt.Columns.Add("ProdPlanRuleId", GetType(Long))

            ' 明細項目（Gridで扱う列）
            dt.Columns.Add("Qty", GetType(Integer))
            dt.Columns.Add("QtyConditionType", GetType(String))
            dt.Columns.Add("SplitMethodType", GetType(Integer))
            dt.Columns.Add("ActiveFlag", GetType(String))

            ' 0件時ダミー行判定
            dt.Columns.Add(DUMMY_ROW_FLAG, GetType(Boolean))

            ' 既定値／NULL許容
            dt.Columns("ActiveFlag").DefaultValue = "Y"

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

            dt = BuildEmptySplitCaseTableSchema()
            dt.AcceptChanges() ' 起点
            SetSessionTable(dt)
        End Sub

        ' この画面の明細グリッドだけをバインド（0件時はダミー行を注入してFooterを出す）
        Private Sub BindSplitCaseGrid()
            EnsureSessionTable()

            Dim dt As DataTable = GetSessionTable()
            If dt Is Nothing Then
                dt = BuildEmptySplitCaseTableSchema()
                SetSessionTable(dt)
            End If

            ' 0件（CurrentRows）であれば、Footerを出すためのダミー行を1件追加
            Dim currentCount As Integer = dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length
            If currentCount = 0 Then
                Dim dummy As DataRow = dt.NewRow()
                dummy("TempId") = "DUMMY"
                dummy("SplitCaseId") = DBNull.Value
                dummy("ProdPlanRuleId") = DBNull.Value
                dummy(DUMMY_ROW_FLAG) = True
                dt.Rows.Add(dummy)
            End If

            Dim dv As New DataView(dt, Nothing, Nothing, DataViewRowState.CurrentRows)
            gvSplitCaseList.DataSource = dv
            gvSplitCaseList.DataBind()
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

                ' 編集行のみ DDL をバインド（Createでも編集は可能）
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
        Protected Sub gvSplitCaseList_RowEditing(sender As Object, e As GridViewEditEventArgs) _
            Handles gvSplitCaseList.RowEditing
            gvSplitCaseList.EditIndex = e.NewEditIndex
            BindSplitCaseGrid()
        End Sub

        ' 編集キャンセル
        Protected Sub gvSplitCaseList_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) _
            Handles gvSplitCaseList.RowCancelingEdit
            e.Cancel = True
            gvSplitCaseList.EditIndex = -1
            BindSplitCaseGrid()
        End Sub

        ' 編集確定（画面内 DataTable を更新）
        Protected Sub gvSplitCaseList_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) _
            Handles gvSplitCaseList.RowUpdating
            Dim dt As DataTable = GetSessionTable()
            Dim row As GridViewRow = gvSplitCaseList.Rows(e.RowIndex)

            ' TempId で行特定（Createは NEW: のみ）
            Dim tempId As String = gvSplitCaseList.DataKeys(e.RowIndex).Values("TempId").ToString()
            Dim dr As DataRow = dt.AsEnumerable().
                First(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            '数量にマイナスや文字が入力されていたら入力を0にする
            Dim qtyText As String = SafeGet(e.NewValues, "Qty")
            Dim qty As Integer
            Dim hasQty As Boolean = Integer.TryParse(qtyText, qty)
            If Not hasQty Then
                qtyText = "0"
            ElseIf qtyText < 0 Then
                qtyText = "0"
            End If

            'dr("Qty") = ToIntOrDBNull(SafeGet(e.NewValues, "Qty"))
            dr("Qty") = ToIntOrDBNull(qtyText)
            dr("QtyConditionType") = GetDDL(row, "ddlQtyConditionType")
            dr("SplitMethodType") = GetDDL(row, "ddlSplitMethodType")
            dr("ActiveFlag") = GetDDL(row, "ddlActiveFlag")

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

            Dim r As DataRow = dt.NewRow()
            r("TempId") = "NEW:" & Guid.NewGuid().ToString("N")
            r("SplitCaseId") = DBNull.Value
            r("ProdPlanRuleId") = DBNull.Value

            Dim txtQtyF = TryCast(f.FindControl("txtQty_F"), TextBox)
            Dim qtyText As String = If(txtQtyF?.Text, String.Empty).Trim()
            Dim qty As Integer
            Dim hasQty As Boolean = Integer.TryParse(qtyText, qty)
            Dim errors As New List(Of String)

            If Not hasQty Then
                errors.Add("数量（Qty）は数値で入力してください。")
            ElseIf qtyText < 0 Then
                errors.Add("数量（Qty）が不正です。")
            End If

            If errors.Count > 0 Then
                lblDetailError.Text = Server.HtmlEncode(String.Join(" / ", errors))
                lblResult.Text = String.Empty
                If Not hasQty Then
                    txtQtyF?.Focus()
                End If
                Return
            End If

            r("Qty") = ToIntOrDBNull(TryCast(f.FindControl("txtQty_F"), TextBox)?.Text)
            r("QtyConditionType") = TryCast(f.FindControl("ddlQtyConditionType_F"), DropDownList)?.SelectedValue
            r("SplitMethodType") = TryCast(f.FindControl("ddlSplitMethodType_F"), DropDownList)?.SelectedValue
            r("ActiveFlag") = TryCast(f.FindControl("ddlActiveFlag_F"), DropDownList)?.SelectedValue

            If dt.Columns.Contains(DUMMY_ROW_FLAG) Then r(DUMMY_ROW_FLAG) = False

            dt.Rows.Add(r)
            BindSplitCaseGrid()

            lblDetailError.Text = ""
            lblDetailResult.Text = ""
        End Sub

        ' 削除（新規：Remove のみ）
        Protected Sub gvSplitCaseList_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) _
            Handles gvSplitCaseList.RowDeleting
            Dim dt As DataTable = GetSessionTable()
            Dim tempId As String = gvSplitCaseList.DataKeys(e.RowIndex).Values("TempId").ToString()

            Dim dr As DataRow = dt.AsEnumerable().
                FirstOrDefault(Function(r) r.RowState <> DataRowState.Deleted AndAlso r.Field(Of String)("TempId") = tempId)

            If dr IsNot Nothing Then
                dt.Rows.Remove(dr)
            End If

            BindSplitCaseGrid()
            e.Cancel = True
        End Sub
#End Region

#Region "ボタンイベント"
        ' 一覧へボタン
        Protected Sub btnProdPlanRuleList_Click(sender As Object, e As EventArgs)
            Response.Redirect("ProdPlanRuleList.aspx")
        End Sub

        ' 登録ボタン
        Protected Sub btnCreateProdPlanRuleSetting_Click(sender As Object, e As EventArgs)

            lblResult.Text = ""
            lblError.Text = ""

            Try
                ' 入力取得
                Dim customerCode As String = (If(txtCustomerCode.Value, "")).Trim()
                Dim profitCenterRaw As String = (If(txtProfitCenter.Value, "")).Trim()
                Dim customerUnitName As String = (If(txtCustomerUnitName.Value, "")).Trim()
                Dim splitFlag As String = (If(ddlSplitFlag.SelectedValue, "")).Trim()
                Dim splitCaseFlag As String = (If(ddlSplitCaseFlag.SelectedValue, "")).Trim()
                Dim splitMethodType As String = (If(ddlSplitMethodType.SelectedValue, "")).Trim()
                Dim splitRationType As String = (If(ddlSplitRationType.SelectedValue, "")).Trim()
                Dim splitRoudingUnit As String = (If(txtSplitRoudingUnit.Text, "")).Trim()
                Dim carryToType As String = (If(ddlCarryToType.SelectedValue, "")).Trim()
                Dim splitStartType As String = (If(ddlSplitStartType.SelectedValue, "")).Trim()

                Dim repoCust As New CustomerRepository(Utils.GetConnectionString())
                Dim repoProdPlanRule As New ProdPlanRuleRepository(Utils.GetConnectionString())

                Dim userId As String = PageHelpers.GetUserId(Me.Page)
                If String.IsNullOrEmpty(userId) Then
                    lblError.Text = "ログイン情報が見つかりません。"
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
                Dim existsOther As Boolean = repoProdPlanRule.ExistsProdPlanRule(
                    customerSettingId:=customerSettingId,
                    excludeProdPlanRuleId:=0 ' 新規のため除外なし
                )
                If existsOther Then
                    lblError.Text = "同一（取引先コード、PC、注文工場／担当者名）の登録が見つかりました。"
                    Return
                End If

                ' まるめ数　入力チェック
                If splitRoudingUnit < 0 Then
                    lblError.Text = "まるめ数が不正です。"
                    Return
                End If

                ' ▼ 1) ヘッダーをINSERTして新しい ProfileId を採番（ヘッダー入力欄のID確定後に実装）
                Dim repo As New ProdPlanRuleRepository(Utils.GetConnectionString())
                Dim newProdPlanRuleId As Long = repo.InsertProdPlanRule(
                    customerSettingId:=customerSettingId.ToString(),
                    splitFlag:=splitFlag,
                    splitCaseFlag:=splitCaseFlag,
                    splitMethodType:=splitMethodType,
                    splitRationType:=splitRationType,
                    splitRoudingUnit:=splitRoudingUnit,
                    carryToType:=carryToType,
                    splitStartType:=splitStartType,
                    userId:=userId,
                    pgId:="CreateProdPlanRule(Header)"
                )


                ' ▼ 2) 明細を一括INSERT（RowState=Added）
                Dim dt As DataTable = GetSessionTable()
                For Each r As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.Added)
                    ' ダミーデータ判定（TempId = "DUMMY" の場合はスキップ）
                    If Not IsDBNull(r("TempId")) AndAlso r("TempId").ToString().Trim().ToUpper() = "DUMMY" Then
                        Continue For
                    End If

                    ' 数量が空欄の場合はスキップ
                    If IsDBNull(r("Qty")) Then
                        Continue For
                    End If

                    repo.InsertSplitCase(
                        prodPlanRuleId:=newProdPlanRuleId.ToString(),
                        qty:=r("Qty").ToString(),
                        qtyConditionType:=r("QtyConditionType").ToString(),
                        splitMethodType:=r("SplitMethodType").ToString(),
                        userId:=userId,
                        pgId:="CreateProdPlanRule(Detail)"
                    )
                Next

                If newProdPlanRuleId <= 0 Then
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

#Region "OnPreRender"
        Protected Overrides Sub OnPreRender(e As EventArgs)
            MyBase.OnPreRender(e)
            ' lblError が DOM に出るタイミングで確実に ClientID が決まっている
            btnCreateProdPlanRuleSetting.Attributes("data-error-label-id") = lblError.ClientID
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