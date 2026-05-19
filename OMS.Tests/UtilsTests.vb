Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OMS.Common

<TestClass>
Public Class UtilsTests

    <TestMethod>
    Public Sub BuildLikePattern_Contains_EscapesAndWraps()
        Dim actual = Utils.BuildLikePattern("a%b_c\z", Utils.LikeMode.Contains)
        ' バックスラッシュは二重化、%/_ はエスケープ
        Assert.AreEqual("%a\%b\_c\\z%", actual)
    End Sub

    <TestMethod>
    Public Sub BuildLikePattern_NullOrWhite_ReturnsNothing()
        Assert.IsNull(Utils.BuildLikePattern(Nothing, Utils.LikeMode.Contains))
        Assert.IsNull(Utils.BuildLikePattern("   ", Utils.LikeMode.Contains))
    End Sub

    <TestMethod>
    Public Sub SafeFolderName_RemovesInvalidChars_TruncatesAndDomainStrips()
        Dim s = "DOMAIN\user:inv*?<>|. "
        Dim actual = Utils.SafeFolderName(s, 10)
        ' DOMAIN\ 切除→ user:inv*?<>|. から禁止文字除去＆末尾の '.' 削除 → "userinv"
        Assert.AreEqual("userinv", actual)
    End Sub

    <TestMethod>
    Public Sub NormalizeString_Upper()
        Assert.AreEqual("ABC", Utils.NormalizeString("  Abc "))
        Assert.AreEqual(String.Empty, Utils.NormalizeString(Nothing))
    End Sub

    <TestMethod>
    Public Sub NormalizeYN_OnlyYorN_DefaultN()
        Assert.AreEqual("Y", Utils.NormalizeYN("y"))
        Assert.AreEqual("N", Utils.NormalizeYN("n"))
        Assert.AreEqual("N", Utils.NormalizeYN(""))
        Assert.AreEqual("N", Utils.NormalizeYN("true"))
    End Sub

    <TestMethod>
    Public Sub SafeVarchar_TrimAndCut()
        Assert.AreEqual(DBNull.Value, Utils.SafeVarchar(Nothing, 10))
        Assert.AreEqual("abc", Utils.SafeVarchar("  abc  ", 10))
        Assert.AreEqual("12345", Utils.SafeVarchar("1234567", 5))
    End Sub

    <TestMethod>
    Public Sub NullIfWhite_EmptyToNothing()
        Assert.IsNull(Utils.NullIfWhite(Nothing))
        Assert.IsNull(Utils.NullIfWhite("   "))
        Assert.AreEqual("A", Utils.NullIfWhite(" A "))
    End Sub

    <TestMethod>
    Public Sub FormatDate_YyyyMMdd()
        Dim dt As New DateTime(2025, 11, 4)
        Assert.AreEqual("2025/11/04", Utils.FormatDate(dt))
    End Sub

    <TestMethod>
    Public Sub BuildOptions_EncodesValues()
        Dim values = New String() {"A&B", """Q""", "<tag>"}
        Dim html = Utils.BuildOptions(values)
        ' 例: <option value="A&amp;B" />
        StringAssert.Contains(html, "<option value=""A&amp;B"" />")
        StringAssert.Contains(html, "<option value=""&quot;Q&quot;"" />")
        StringAssert.Contains(html, "<option value=""&lt;tag>"" />")
    End Sub

End Class