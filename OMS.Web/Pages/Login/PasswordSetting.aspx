<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PasswordSetting.aspx.vb" Inherits="OMS.Web.PasswordSetting" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>パスワード設定</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Input.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Menu.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server" novalidate="novalidate">

    <!-- ヘッダー -->
    <div class="menu-wrapper">
        <div class="menu-header">
            <h1>パスワード設定</h1>
            <div class="user-info">
                <asp:Button ID="btnReturn" runat="server" CssClass="btn-logout" Text="ログインへ" OnClick="btnReturn_Click" />
            </div>
        </div>
        <!-- 入力 -->
<div>
    <span class="label-title">ユーザーID*</span>
    <asp:TextBox ID="txtUserID" runat="server" />
</div>
<div>
    <span class="label-title">現在のパスワード</span>
    <asp:TextBox ID="txtCurrentPassword" runat="server" TextMode="Password" />
</div>
<div>
    <span class="label-title">新しいパスワード*</span>
    <asp:TextBox ID="txtNewPassword" runat="server" TextMode="Password" />
</div>
<div>
    <span class="label-title">新しいパスワード（確認用）*</span>
    <asp:TextBox ID="txtConfirmPassword" runat="server" TextMode="Password" />
</div>

<!-- アクションボタン -->
<div class="action-buttons">
    <asp:Button ID="btnSavePasswordSetting" runat="server"
        CssClass="btn-asti btn-asti-process"
        Text="更新" OnClick="btnSavePasswordSetting_Click" />
</div>

<!-- 結果表示 -->
<div>
    <br />
    <asp:Label ID="lblResult" runat="server" ForeColor="Green" />
    <asp:Label ID="lblError" runat="server" CssClass="validation-summary" />
</div>
    </div>

    
    </form>
</body>
</html>
