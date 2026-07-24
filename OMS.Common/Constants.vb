Imports System.Collections.Generic
Imports System.Web.UI.WebControls
Imports System.Linq ' ← ToList などを使う場合

Namespace OMS.Common
    ''' <summary>
    ''' アプリ全体で使用する定数・マップ・簡易UIヘルパ。
    ''' </summary>
    Public Module Constants

#Region "管理者（ログイン）"
        Public Const AdminUserID As String = "admin"
        Public Const AdminUserPass As String = "admin"
#End Region

#Region "CSV 出力：区切り/囲み/改行/文字コード（コード値）"
        Public Const DELIMITER_COMMA As String = "COMMA"
        Public Const DELIMITER_TAB As String = "TAB"
        Public Const DELIMITER_SEMICOLON As String = "SEMICOLON"
        Public Const DELIMITER_PIPE As String = "PIPE"
        Public Const DELIMITER_SPACE As String = "SPACE"
        Public Const DELIMITER_COLON As String = "COLON"

        Public Const QUOTE_D As String = "D_QUOTE"
        Public Const QUOTE_S As String = "S_QUOTE"
        Public Const QUOTE_NONE As String = "NONE"

        Public Const NEWLINE_CRLF As String = "CRLF"
        Public Const NEWLINE_LF As String = "LF"
        Public Const NEWLINE_CR As String = "CR"

        Public Const ENCODING_SJIS As String = "SJIS"
        Public Const ENCODING_UTF8 As String = "UTF8"
        Public Const ENCODING_UTF8_BOM As String = "UTF8-BOM"
#End Region

#Region "ドメイン：コード → 表示名のマップ"
        Public ReadOnly FolderTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "内示"},
                {2, "確定"},
                {3, "納入指示"},
                {4, "混在"}
            }

        Public ReadOnly OrderTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "内示"},
                {2, "確定"},
                {3, "納入指示"}
            }

        Public ReadOnly FormatTypeMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"LIST", "リスト"},
                {"MATRIX", "マトリックス"}
            }

        Public ReadOnly InfoTypeMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"I", "新規"},
                {"U", "更新"},
                {"D", "取消"},
                {"N", "打切"}
            }

        Public ReadOnly TargetFieldMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"CUSTOMER_ORDER_NO", "客先発注No"},
                {"ORDER_DATE", "受注日"},
                {"DUE_DATE", "希望納期"},
                {"CUSTOMER_ITEM_NO", "客先品目No"},
                {"DEMAND_QTY", "需要数"},
                {"CURRENCY_CODE", "通貨コード"},
                {"PRODUCT_CODE", "製品コード"},
                {"REMARKS", "コメント"},
                {"DELIVERY_CODE", "納入先コード"},
                {"ORDER_TYPE", "受注区分"},
                {"PRORATED_TYPE", "分割区分"},
                {"CUSTOMER_INFO_TYPE", "取引先情報区分"},
                {"SELF_FCST_FLAG", "自社予測フラグ"},
                {"SELF_FCST_DELETE_FLAG", "自社予測削除フラグ"}
            }

        Public ReadOnly RowSelectorTypeMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"FIXED", "固定行を指定"},
                {"BY_HEADER", "ヘッダー名で列を特定"},
                {"BY_KEY", "キー値で行を検索"},
                {"OFFSET", "基準から相対オフセット"}
            }

        Public ReadOnly DataTypeMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"STRING", "文字列"},
                {"NUMBER", "数値"},
                {"DATE", "日付／時刻"}
            }

        Public ReadOnly TrimFlagMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"Y", "必要"},
                {"N", "不要"}
            }

        Public ReadOnly ActiveFlagMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"Y", "有効"},
                {"N", "無効"}
            }

        Public ReadOnly QtyConditionTypeMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"GE", "以上"},
                {"LE", "以下"},
                {"GT", "越え"},
                {"LT", "未満"}
            }

        Public ReadOnly SplitMethodTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {0, "分割しない"},
                {1, "日割り"},
                {2, "4分割"},
                {3, "3分割"},
                {4, "2分割"},
                {5, "週まるめ"},
                {6, "ハンド"}
            }

        Public ReadOnly ProratedTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "日割り"},
                {2, "日割り以外"}
            }

        Public ReadOnly SplitFlagMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String) From {
                {"Y", "する"},
                {"N", "しない"}
            }

        Public ReadOnly FcstRollupTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {0, "しない"},
                {1, "月まとめ"},
                {2, "全数まとめ"}
            }

        Public ReadOnly SplitCaseFlagMap As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String) From {
                {"Y", "使用する"},
                {"N", "使用しない"}
            }

        Public ReadOnly SplitRationTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "1:1"},
                {2, "2:1"}
            }

        Public ReadOnly CarryToTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "先頭"},
                {2, "最後尾"}
            }

        Public ReadOnly SplitStartTypeMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {1, "月初"},
                {2, "前月第4週"},
                {3, "納期の4週前"}
            }
        Public ReadOnly AuthorityLevelMap As IReadOnlyDictionary(Of Integer, String) =
            New Dictionary(Of Integer, String) From {
                {0, "作業者"},
                {1, "管理者"}
            }
#End Region

#Region "変換/バインド ヘルパ"

        ''' <summary>
        ''' 文字列コードを表示名へ変換（ヒットしなければ fallback→code の順で返す）。
        ''' </summary>
        Public Function Lookup(Of TKey)(
            map As IReadOnlyDictionary(Of TKey, String),
            keyObj As Object,
            Optional fallback As String = ""
        ) As String
            If map Is Nothing OrElse keyObj Is Nothing OrElse keyObj Is DBNull.Value Then
                Return fallback
            End If

            Dim key As TKey
            Try
                ' 基本の型変換（String→TKey / Object→TKey）
                key = DirectCast(Convert.ChangeType(keyObj, GetType(TKey)), TKey)
            Catch
                ' 変換不可のときは keyObj をそのまま返すか fallback
                Dim s = keyObj.ToString().Trim()
                Return If(fallback <> "", fallback, s)
            End Try

            Dim name As String = Nothing
            If map.TryGetValue(key, name) Then Return name
            Return If(fallback <> "", fallback, key?.ToString())
        End Function

        ''' <summary>
        ''' DropDownList にマップをバインドする（既定はラベル昇順）。
        ''' </summary>
        Public Sub BindDropDown(Of TKey)(
            ddl As DropDownList,
            map As IReadOnlyDictionary(Of TKey, String),
            Optional sortByLabel As Boolean = True
        )
            If ddl Is Nothing Then Exit Sub
            ddl.Items.Clear()
            If map Is Nothing OrElse map.Count = 0 Then Exit Sub

            Dim items As List(Of KeyValuePair(Of TKey, String)) = map.ToList()

            If sortByLabel Then
                items.Sort(Function(a, b) StringComparer.CurrentCulture.Compare(a.Value, b.Value))
            End If

            For Each kv In items
                ddl.Items.Add(New ListItem(kv.Value, kv.Key?.ToString()))
            Next
        End Sub

#End Region

    End Module
End Namespace