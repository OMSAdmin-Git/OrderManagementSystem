Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Web.UI.WebControls
Imports OMS.Common

<TestClass>
Public Class ConstantsTests

    ' ========== Lookup のテスト ==========
    <TestMethod>
    Public Sub Lookup_Hit_ReturnsDisplayName()
        Dim name = Constants.Lookup(Constants.FormatTypeMap, "LIST", fallback:="")
        Assert.AreEqual("リスト", name)
    End Sub

    <TestMethod>
    Public Sub Lookup_Miss_WithFallback_ReturnsFallback()
        Dim name = Constants.Lookup(Constants.FormatTypeMap, "UNKNOWN", fallback:="（未定義）")
        Assert.AreEqual("（未定義）", name)
    End Sub

    <TestMethod>
    Public Sub Lookup_Miss_WithoutFallback_ReturnsCode()
        Dim name = Constants.Lookup(Constants.FormatTypeMap, "UNKNOWN", fallback:="")
        Assert.AreEqual("UNKNOWN", name)
    End Sub

    <TestMethod>
    Public Sub Lookup_NullOrWhitespace_ReturnsFallback()
        Assert.AreEqual("N/A", Constants.Lookup(Constants.TrimFlagMap, Nothing, "N/A"))
        Assert.AreEqual("N/A", Constants.Lookup(Constants.TrimFlagMap, "   ", "N/A"))
    End Sub

    <TestMethod>
    Public Sub Lookup_CaseInsensitive_And_Trimmed()
        Dim name = Constants.Lookup(Constants.DataTypeMap, "  string  ", fallback:="")
        Assert.AreEqual("文字列", name)
        Dim name2 = Constants.Lookup(Constants.DataTypeMap, "dAtE", fallback:="")
        Assert.AreEqual("日付／時刻", name2)
    End Sub

    ' ========== BindDropDown のテスト ==========
    <TestMethod>
    Public Sub BindDropDown_SortsByLabel_WhenDefault()
        Dim ddl As New DropDownList()
        ' {Y:"必要", N:"不要"}
        Constants.BindDropDown(ddl, Constants.TrimFlagMap)   ' 既定: sortByLabel=True

        ' 期待順をカルチャに合わせて生成（ハードコードしない）
        Dim expected = Constants.TrimFlagMap.OrderBy(Function(kv) kv.Value, StringComparer.CurrentCulture) _
                                     .Select(Function(kv) New With {.Text = kv.Value, .Value = kv.Key}) _
                                     .ToList()

        Assert.AreEqual(expected.Count, ddl.Items.Count)
        For i = 0 To expected.Count - 1
            Assert.AreEqual(expected(i).Text, ddl.Items(i).Text, $"index={i}")
            Assert.AreEqual(expected(i).Value, ddl.Items(i).Value, $"index={i}")
        Next
    End Sub

    <TestMethod>
    Public Sub BindDropDown_KeepsInsertOrder_WhenSortByLabelFalse()
        ' 独自の小さいマップ（挿入順 A→C→B）を用意
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"A", "あ"}, {"C", "う"}, {"B", "い"}
        }
        Dim ddl As New DropDownList()
        Constants.BindDropDown(ddl, map, sortByLabel:=False)

        Assert.AreEqual(3, ddl.Items.Count)
        Assert.AreEqual("A", ddl.Items(0).Value) ' 挿入順を維持
        Assert.AreEqual("C", ddl.Items(1).Value)
        Assert.AreEqual("B", ddl.Items(2).Value)
    End Sub

    <TestMethod>
    Public Sub BindDropDown_ClearsExistingItems()
        Dim ddl As New DropDownList()
        ddl.Items.Add(New ListItem("old", "old"))
        Assert.AreEqual(1, ddl.Items.Count)

        Constants.BindDropDown(ddl, Constants.ActiveFlagMap)

        Dim expected = Constants.ActiveFlagMap.OrderBy(Function(kv) kv.Value, StringComparer.CurrentCulture) _
                                   .Select(Function(kv) New With {.Text = kv.Value, .Value = kv.Key}) _
                                   .ToList()

        Assert.AreEqual(expected.Count, ddl.Items.Count)
        For i = 0 To expected.Count - 1
            Assert.AreEqual(expected(i).Text, ddl.Items(i).Text, $"index={i}")
            Assert.AreEqual(expected(i).Value, ddl.Items(i).Value, $"index={i}")
        Next
    End Sub

    <TestMethod>
    Public Sub BindDropDown_NullOrEmpty_DoesNothing()
        Dim ddl As New DropDownList()
        ddl.Items.Add(New ListItem("keep", "keep"))
        Assert.AreEqual(1, ddl.Items.Count)

        ' map Nothing → クリアのみ実行、追加なし
        'Constants.BindDropDown(ddl, Nothing)
        Constants.BindDropDown(Of String)(ddl, Nothing)
        Assert.AreEqual(0, ddl.Items.Count)

        ' 空マップ → クリアのみ実行、追加なし
        ddl.Items.Add(New ListItem("keep2", "keep2"))
        Dim empty As New Dictionary(Of String, String)()
        Constants.BindDropDown(ddl, empty)
        Assert.AreEqual(0, ddl.Items.Count)
    End Sub

    <TestMethod>
    Public Sub BindDropDown_NullDdl_DoesNotThrow()
        ' 例外にならないことを確認（操作対象なし）
        Constants.BindDropDown(Nothing, Constants.DataTypeMap)
    End Sub

    ' ========== （任意）契約テストの例 ==========
    ' 外部仕様と結びついた「コード値」を将来誤って変更しないための“検知用”スモーク。
    ' 頻繁に変わらない前提のコードだけを最小限チェックするのがおすすめです。
    <TestMethod>
    Public Sub Contract_SomeDelimiterCodes_MustRemain()
        Assert.AreEqual("COMMA", Constants.DELIMITER_COMMA)
        Assert.AreEqual("TAB", Constants.DELIMITER_TAB)
        Assert.AreEqual("NONE", Constants.QUOTE_NONE)
    End Sub

End Class