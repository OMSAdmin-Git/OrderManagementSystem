
Imports System.Data
Imports System.Text
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports OMS.Data.UsrDeffIdfRepository
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data

    ''' <summary>
    ''' 分割 設定
    ''' </summary>
    Public Class SplitCaseRepository

            Private ReadOnly _connectionString As String

            Public Sub New(connectionString As String)
                _connectionString = connectionString
            End Sub
        ' [Field]
        'SPLIT_CASE_ID
        'PROD_PLAN_RULE_ID
        'QTY
        'QTY_CONDITION_TYPE
        'SPLIT_METHOD_TYPE
        'ACTIVE_FLAG
        'CREATED_AT
        'CREATED_USER_ID
        'CREATED_PG_ID
        'UPDATED_AT
        'UPDATED_USER_ID
        'UPDATED_PG_ID
        ''' <summary>
        ''' 分割条件マスタ一覧取得
        ''' </summary>
        ''' <param name="prodPlanRuleId"></param>
        ''' <returns></returns>
        Public Function GetSplitCaseRuleList(prodPlanRuleId As Long) As List(Of SplitCaseRow)
            Dim dt As New DataTable()
            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    dt = GetSplitCaseRuleList(conn, tran, prodPlanRuleId)
                    tran.Commit()
                End Using
            End Using
            Dim dtl = ToClass(dt)
            Return dtl
        End Function

        ''' <summary>
        ''' 分割条件マスタ一覧取得 優先/数量 ソート
        ''' </summary>
        ''' <param name="prodPlanRuleId"></param>
        ''' <returns></returns>
        Public Function GetSplitCaseRuleListSortByPriolity(conn As OracleConnection, tran As OracleTransaction, prodPlanRuleId As Long) As List(Of SplitCaseRow)

            ' 分割条件 ソート優先順位
            Dim ConditionType As New List(Of (type As String, priority As Integer)) From {
                    ("GT", 1), ("GE", 2), ("LT", 3), ("LE", 4)
                }

            Dim splitCaseRuleList = ToClass(GetSplitCaseRuleList(conn, tran, prodPlanRuleId))
            If (splitCaseRuleList.Count = 0) Then
                ' データなし
            End If
            ' 分割条件(QTY_CONDITION_TYPE)でソートを行い 同じ条件内では数量(QTY)の多い順に 
            ' SplitCaseRow を並べ替える
            ' GT > GE > LT > LE の順 (未満を先に判定するため以下より優先が高い)
            'splitCaseRuleList = splitCaseRuleList _
            '                    .Join(ConditionType,
            '                          Function(sd) sd.QtyConditionType,
            '                          Function(ct) ct.type,
            '                          Function(sd, ct) New With {sd, ct.priority}) _
            '                    .OrderBy(Function(x) x.priority) _
            '                    .ThenByDescending(Function(x) x.sd.Qty) _
            '                    .Select(Function(x) x.sd) _
            '                    .ToList()

            splitCaseRuleList = splitCaseRuleList _
                                .Join(ConditionType,
                                      Function(sd) sd.QtyConditionType,
                                      Function(ct) ct.type,
                                      Function(sd, ct) New With {sd, ct.priority}) _
                                .OrderBy(Function(x) x.priority) _
                                .ThenBy(Function(x)
                                            ' conditionごとにソート用の値を計算
                                            If x.sd.QtyConditionType = "GE" Then
                                                ' GEの場合：数量が多い順(降順)にするためマイナスを掛ける
                                                Return -x.sd.Qty
                                            Else
                                                ' それ以外：数量が少ない順(昇順)
                                                Return x.sd.Qty
                                            End If
                                        End Function) _
                                .Select(Function(x) x.sd) _
                                .ToList()
            Return splitCaseRuleList

        End Function
        ''' <summary>
        ''' 分割条件条件マスタ一覧取得
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="prodPlanRuleId"></param>
        ''' <returns></returns>
        Public Function GetSplitCaseRuleList(conn As OracleConnection, tran As OracleTransaction, prodPlanRuleId As Long) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine("  split_case_id          AS ""SplitCaseId"",")
            sb.AppendLine("  prod_plan_rule_id      AS ""ProdPlanRuleId"",")
            sb.AppendLine("  qty                    AS ""Qty"",")
            sb.AppendLine("  qty_condition_type     AS ""QtyConditionType"",")
            sb.AppendLine("  split_method_type      AS ""SplitMethodType"",")
            sb.AppendLine("  active_flag            AS ""ActiveFlag"",")
            sb.AppendLine("  created_at             AS ""CreatedAt"",")
            sb.AppendLine("  created_user_id        AS ""CreatedUserId"",")
            sb.AppendLine("  created_pg_id          AS ""CreatedPgId"",")
            sb.AppendLine("  updated_at             AS ""UpdatedAt"",")
            sb.AppendLine("  updated_user_id        AS ""UpdatedUserId"",")
            sb.AppendLine("  updated_pg_id          AS ""UpdatedPgId"" ")
            sb.AppendLine("FROM split_case_mst ")
            sb.AppendLine("WHERE 1=1 ")
            Dim prm As New List(Of OracleParameter)()
            sb.AppendLine("AND prod_plan_rule_id = :p_prodPlanRuleId ")
            prm.Add(New OracleParameter(":p_prodPlanRuleId", OracleDbType.Long) With {.Value = prodPlanRuleId})

            Using cmd As New OracleCommand(sb.ToString(), conn)
                cmd.BindByName = True
                cmd.CommandType = CommandType.Text
                If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                'conn.Open()
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
            Return dt

        End Function

        ''' <summary>
        ''' DataRow to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As SplitCaseRow

            ' GetSplitCaseRuleList メソッド用 フィールド名が別名注意

            Dim udir = New SplitCaseRow
            udir.SplitCaseId = dt.Field(Of Long)("SplitCaseId")
            udir.ProdPlanRuleId = dt.Field(Of Long?)("ProdPlanRuleId")
            udir.Qty = dt.Field(Of Long?)("Qty")
            udir.QtyConditionType = dt.Field(Of String)("QtyConditionType")
            udir.SplitMethodType = dt.Field(Of Int16?)("SplitMethodType")
            udir.ActiveFlag = dt.Field(Of String)("ActiveFlag")
            udir.CreatedAt = dt.Field(Of Date?)("CreatedAt")
            udir.CreatedUserId = dt.Field(Of String)("CreatedUserId")
            udir.CreatedPgId = dt.Field(Of String)("CreatedPgId")
            udir.UpdatedAt = dt.Field(Of Date?)("UpdatedAt")
            udir.UpdatedUserId = dt.Field(Of String)("UpdatedUserId")
            udir.UpdatedPgId = dt.Field(Of String)("UpdatedPgId")
            Return udir
        End Function

        ''' <summary>
        ''' DataTable to class
        ''' R.sagisaka create
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As IEnumerable(Of SplitCaseRow)

            Dim osrs = New List(Of SplitCaseRow)()

            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next

            Return osrs
        End Function

    End Class

    Public Class SplitCaseRow
        Property SplitCaseId As Long        ' SPLIT_CASE_ID         分割パターンID      NUMBER(10,0)
        Property ProdPlanRuleId As Long?    ' PROD_PLAN_RULE_ID     生産計画条件ID      NUMBER(10,0)
        Property Qty As Long?               ' QTY                   数量                NUMBER(10,0)
        Property QtyConditionType As String ' QTY_CONDITION_TYPE    数量条件区分        VARCHAR2(2)
        Property SplitMethodType As Int16?  ' SPLIT_METHOD_TYPE     分割方法            NUMBER(1,0)
        Property ActiveFlag As String       ' ACTIVE_FLAG           有効フラグ          CHAR(1)
        Property CreatedAt As Date?         ' CREATED_AT            登録日時            DATE
        Property CreatedUserId As String    ' CREATED_USER_ID       登録ユーザーID      VARCHAR2(9)
        Property CreatedPgId As String      ' CREATED_PG_ID         登録プログラムID    VARCHAR2(150)
        Property UpdatedAt As Date?         ' UPDATED_AT            更新日時            DATE
        Property UpdatedUserId As String    ' UPDATED_USER_ID       更新ユーザーID      VARCHAR2(9)
        Property UpdatedPgId As String      ' UPDATED_PG_ID         更新プログラムID    VARCHAR2(150)
    End Class

End Namespace