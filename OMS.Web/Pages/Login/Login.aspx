<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="OMS.Web.Pages.Login.Login" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <title>ログイン</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Login.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="login-container">
            <h2>ログイン</h2>

            <asp:Label ID="lblError" runat="server" CssClass="error-message"></asp:Label>

            <label for="txtUser">ユーザー名</label>
            <asp:TextBox ID="txtUser" runat="server"></asp:TextBox>

            <label for="txtPass">パスワード</label>
            <asp:TextBox ID="txtPass" runat="server" TextMode="Password"></asp:TextBox>

            <asp:Button ID="btnLogin" runat="server" Text="ログイン" OnClick="btnLogin_Click" />
            <asp:Button ID="btnPassword" runat="server" Text="パスワード設定" OnClick="btnPassword_Click" />
        </div>
    </form>
</body>
</html>