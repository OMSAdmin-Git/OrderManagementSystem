Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.IO
Imports System.Text
Imports Ref = System.Reflection
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.SessionState
Imports System.Web.Hosting
Imports OMS.Common

<TestClass>
Public Class PageHelpersTests

    '========================================================
    ' 本番に近い HttpContext（SimpleWorkerRequest） + Session + ApplicationInstance
    '========================================================
    Private Shared Function CreateHttpContextWithSession(Optional pagePath As String = "Pages/TopMenu.aspx") As HttpContext
        ' SimpleWorkerRequest の (appVirtualDir, appPhysicalDir, page, query, writer) 版で作成
        Dim output As New StringWriter()
        Dim appVirtualDir As String = "/"
        Dim appPhysicalDir As String = IO.Path.GetTempPath() ' 実在パスなら何でもOK（テストでは触りません）
        Dim swr As New SimpleWorkerRequest(appVirtualDir, appPhysicalDir, pagePath, "", output)

        Dim ctx As New HttpContext(swr)
        HttpContext.Current = ctx

        ' ---- Session を公式ユーティリティで紐付け ----
        Dim items As New SessionStateItemCollection()
        Dim staticObjects As New HttpStaticObjectsCollection()
        Dim container As New HttpSessionStateContainer(
            "test-session-id",
            items,
            staticObjects,
            20,
            True,
            HttpCookieMode.UseCookies,
            SessionStateMode.InProc,
            False
        )
        SessionStateUtility.AddHttpSessionStateToContext(ctx, container)

        ' ---- ApplicationInstance を差し込み（CompleteRequest 用）----
        Try
            Dim app As New HttpApplication()
            Dim fi = GetType(HttpContext).GetField("_applicationInstance",
                                                   Ref.BindingFlags.NonPublic Or Ref.BindingFlags.Instance)
            If fi IsNot Nothing Then fi.SetValue(ctx, app)
        Catch
            ' 環境差があっても無視（多くの環境で不要）
        End Try

        Return ctx
    End Function

    ' Page に HttpContext/Request/Response を反映（ライフサイクルを回さない代替）
    Private Shared Function CreatePageWithContext() As Page
        Dim ctx = CreateHttpContextWithSession()

        Dim p As New Page()
        Dim bf = Ref.BindingFlags.NonPublic Or Ref.BindingFlags.Instance

        ' 非公開フィールドに注入（Response.Redirect が内部で参照するため必須）
        Dim fContext = GetType(Page).GetField("_context", bf)
        If fContext IsNot Nothing Then
            fContext.SetValue(p, ctx)
        Else
            Throw New InvalidOperationException("Page._context フィールドが見つかりません。")
        End If

        Dim fRequest = GetType(Page).GetField("_request", bf)
        If fRequest IsNot Nothing Then
            fRequest.SetValue(p, ctx.Request)
        End If

        Dim fResponse = GetType(Page).GetField("_response", bf)
        If fResponse IsNot Nothing Then
            fResponse.SetValue(p, ctx.Response)
        Else
            Throw New InvalidOperationException("Page._response フィールドが見つかりません。")
        End If

        ' 現在のハンドラとして Page を設定（内部参照が安定）
        ctx.Handler = p

        Return p
    End Function

    <TestCleanup>
    Public Sub Cleanup()
        ' 他テストへの影響を避けるため都度クリア
        HttpContext.Current = Nothing
    End Sub

    '========================================================
    ' テスト：SetUserName
    '========================================================
    <TestMethod>
    Public Sub SetUserName_WhenPresent_SetsLabelAndNoRedirect()
        Dim page = CreatePageWithContext()
        page.Session("UserName") = "山田 太郎"

        Dim lbl As New System.Web.UI.WebControls.Label()
        PageHelpers.SetUserName(page, lbl)

        Assert.AreEqual("山田 太郎", lbl.Text)

        ' Redirect されていないこと（Location が空）
        Dim loc = page.Response.RedirectLocation
        Assert.IsTrue(String.IsNullOrEmpty(loc))
    End Sub

    <TestMethod>
    Public Sub SetUserName_WhenMissing_RedirectsToLogin()
        Dim page = CreatePageWithContext()
        Dim lbl As New System.Web.UI.WebControls.Label()

        PageHelpers.SetUserName(page, lbl)

        ' ログインへリダイレクト
        Dim loc = page.Response.RedirectLocation
        Assert.AreEqual("~/Pages/Login/Login.aspx", loc)
    End Sub

    '========================================================
    ' テスト：GetUserId
    '========================================================
    <TestMethod>
    Public Sub GetUserId_WhenPresent_ReturnsValue()
        Dim page = CreatePageWithContext()
        page.Session("UserId") = "kawada"

        Dim v = PageHelpers.GetUserId(page)
        Assert.AreEqual("kawada", v)

        Dim loc = page.Response.RedirectLocation
        Assert.IsTrue(String.IsNullOrEmpty(loc))
    End Sub

    <TestMethod>
    Public Sub GetUserId_WhenMissing_RedirectsAndReturnsEmpty()
        Dim page = CreatePageWithContext()

        Dim v = PageHelpers.GetUserId(page)

        Assert.AreEqual(String.Empty, v)
        Dim loc = page.Response.RedirectLocation
        Assert.AreEqual("~/Pages/Login/Login.aspx", loc)
    End Sub

End Class