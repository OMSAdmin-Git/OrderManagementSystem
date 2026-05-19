Imports System
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Web
Imports System.Web.Hosting
Imports System.Web.UI.WebControls
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OMS.Common
Imports OMS.Web
Imports System.Collections.Generic  ' ★ 追加
' Imports OMS.Data  ' ← 認証系テストを有効化するときに参照追加の上、コメント解除

<TestClass>
Public Class LoginPageTests

    '========================================================
    ' カスタム WorkerRequest：ServerVariables を供給する
    '========================================================
    Private NotInheritable Class TestWorkerRequest
        Inherits SimpleWorkerRequest

        Private ReadOnly _serverVars As Dictionary(Of String, String)

        Public Sub New(appVDir As String,
                   appPDir As String,
                   page As String,
                   query As String,
                   output As TextWriter,
                   Optional serverVars As Dictionary(Of String, String) = Nothing)
            MyBase.New(appVDir, appPDir, page, query, output)
            ' ★ Null合体は OrElse ではなく If(a, b) を使う
            _serverVars = If(serverVars, New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))
        End Sub

        Public Overrides Function GetServerVariable(name As String) As String
            Dim v As String = Nothing
            If _serverVars.TryGetValue(name, v) Then
                Return v
            End If
            Return MyBase.GetServerVariable(name)
        End Function
    End Class


    '========================================================
    ' HttpContext / Page ヘルパ
    '========================================================
    Private Shared Function CreateContext(Optional xff As String = Nothing,
                                          Optional remoteAddr As String = "127.0.0.1") As HttpContext
        Dim output As New StringWriter()
        Dim vars As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If Not String.IsNullOrEmpty(xff) Then vars("HTTP_X_FORWARDED_FOR") = xff
        If Not String.IsNullOrEmpty(remoteAddr) Then vars("REMOTE_ADDR") = remoteAddr

        ' SimpleWorkerRequest( appVDir, appPDir, page, query, writer )
        Dim swr As New TestWorkerRequest("/",
                                         Path.GetTempPath(),
                                         "Pages/Login/Login.aspx",
                                         "",
                                         output,
                                         vars)
        Dim ctx As New HttpContext(swr)
        HttpContext.Current = ctx

        ' CompleteRequest 対策：ApplicationInstance を差し込む
        Try
            Dim app As New HttpApplication()
            Dim fi = GetType(HttpContext).GetField("_applicationInstance",
                                                   BindingFlags.NonPublic Or BindingFlags.Instance)
            If fi IsNot Nothing Then fi.SetValue(ctx, app)
        Catch
        End Try
        Return ctx
    End Function

    Private Shared Function CreateLoginPage(Optional xff As String = Nothing,
                                            Optional remoteAddr As String = "127.0.0.1") As Pages.Login.Login
        Dim ctx = CreateContext(xff, remoteAddr)
        Dim p As New Pages.Login.Login()

        ' Page 内部フィールドに Context/Request/Response を注入
        Dim bf = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim tPage = GetType(System.Web.UI.Page)
        tPage.GetField("_context", bf).SetValue(p, ctx)
        tPage.GetField("_request", bf).SetValue(p, ctx.Request)
        tPage.GetField("_response", bf).SetValue(p, ctx.Response)
        ctx.Handler = p
        Return p
    End Function

    Private Shared Sub SetPrivateField(target As Object, name As String, value As Object)
        Dim bf = BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.Public
        Dim fi = target.GetType().GetField(name, bf)
        If fi Is Nothing Then Throw New MissingFieldException(target.GetType().FullName, name)
        fi.SetValue(target, value)
    End Sub

    ' ログインページが参照するコントロール（フィールド）に実体を割当
    Private Shared Sub WireControls(p As Pages.Login.Login)
        Dim txtUser As New TextBox() With {.ID = "txtUser"}
        Dim txtPass As New TextBox() With {.ID = "txtPass"}
        Dim lblError As New Label() With {.ID = "lblError"}
        Dim btnLogin As New Button() With {.ID = "btnLogin"}

        ' コントロールツリーにも載せておく（FindControl 等で参照できるように）
        p.Controls.Add(txtUser)
        p.Controls.Add(txtPass)
        p.Controls.Add(lblError)
        p.Controls.Add(btnLogin)

        ' ★ ページのフィールド（designer が生成する WithEvents フィールド）へ反射で割当
        '    これをしないと、btnLogin_Click 内の txtUser 等が Nothing になり NRE になる
        SetPrivateField(p, "txtUser", txtUser)
        SetPrivateField(p, "txtPass", txtPass)
        SetPrivateField(p, "lblError", lblError)
        SetPrivateField(p, "btnLogin", btnLogin)
    End Sub

    Private Shared Function GetRedirectLocation() As String
        Return HttpContext.Current.Response.RedirectLocation
    End Function

    '========================================================
    ' 1) 入力未設定 → エラーメッセージ、Redirectなし
    '========================================================
    <TestMethod>
    Public Sub Login_EmptyInput_ShowsError_NoRedirect()
        Dim p = CreateLoginPage()
        WireControls(p)

        Dim txtUser = DirectCast(p.FindControl("txtUser"), TextBox)
        Dim txtPass = DirectCast(p.FindControl("txtPass"), TextBox)
        Dim lbl = DirectCast(p.FindControl("lblError"), Label)
        txtUser.Text = "" : txtPass.Text = ""

        Dim mi = GetType(Pages.Login.Login).GetMethod("btnLogin_Click",
                    BindingFlags.NonPublic Or BindingFlags.Instance)
        mi.Invoke(p, New Object() {Nothing, EventArgs.Empty})

        StringAssert.Contains(lbl.Text, "入力してください")
        Assert.IsTrue(String.IsNullOrEmpty(GetRedirectLocation()))
    End Sub

    '========================================================
    ' 2) GetClientIp：X-Forwarded-For 優先
    '========================================================
    <TestMethod>
    Public Sub GetClientIp_TakesForwardedFor()
        ' ServerVariables は書換不可なので、Context 構築時に供給する
        Dim p = CreateLoginPage("203.0.113.10, 198.51.100.5", "192.0.2.1")
        WireControls(p)

        Dim mi = GetType(Pages.Login.Login).GetMethod("GetClientIp",
                    BindingFlags.NonPublic Or BindingFlags.Instance)
        Dim ip = DirectCast(mi.Invoke(p, Nothing), String)
        Assert.AreEqual("203.0.113.10", ip)
    End Sub

    '========================================================
    ' 3) GetClientIp：ForwardedFor 無 → REMOTE_ADDR
    '========================================================
    <TestMethod>
    Public Sub GetClientIp_UsesRemoteAddrWhenNoForwardedFor()
        Dim p = CreateLoginPage(Nothing, "192.0.2.99")
        WireControls(p)

        Dim mi = GetType(Pages.Login.Login).GetMethod("GetClientIp",
                    BindingFlags.NonPublic Or BindingFlags.Instance)
        Dim ip = DirectCast(mi.Invoke(p, Nothing), String)
        Assert.AreEqual("192.0.2.99", ip)
    End Sub

    '========================================================
    ' ▼ 認証シナリオ（DB 依存）は DI 整備後に有効化
    '========================================================
    <TestMethod, Ignore("UserRepository/Logger の差し替え可能化後に有効化してください。")>
    Public Sub Login_Success_RedirectsAndSetsSession()
        Assert.Inconclusive()
    End Sub

    <TestMethod, Ignore("UserRepository 失敗 + admin 直書き一致の差し替え後に有効化してください。")>
    Public Sub Login_AdminCredentials_Success()
        Assert.Inconclusive()
    End Sub

    <TestMethod, Ignore("UserRepository 失敗の差し替え後に有効化してください。")>
    Public Sub Login_Fail_ShowsError_NoRedirect()
        Assert.Inconclusive()
    End Sub

    <TestMethod, Ignore("例外を投げる差し替え後に有効化してください。")>
    Public Sub Login_Exception_ShowsGenericError()
        Assert.Inconclusive()
    End Sub

End Class