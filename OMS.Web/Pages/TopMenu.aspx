<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TopMenu.aspx.vb" Inherits="OMS.Web.Pages.TopMenu" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>トップメニュー</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Menu.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>トップメニュー</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;
                    <asp:Button ID="btnLogout" runat="server" CssClass="btn-logout" Text="ログアウト" OnClick="btnLogout_Click" />
                </div>
            </div>

            <div class="menu-grid">
                <button id="btnOrder" runat="server" class="menu-tile btn-asti" onserverclick="btnOrder_ServerClick">
                    <div class="tile-title">受注管理</div>
                    <div class="tile-desc">注文登録・検索・CSV作成</div>
                </button>
                <button id="btnMaster" runat="server" class="menu-tile btn-asti" onserverclick="btnMaster_ServerClick">
                    <div class="tile-title">マスタ管理</div>
                    <div class="tile-desc">取込先情報・システム情報</div>
                </button>
            </div>

            <div class="menu-actions">
            </div>
        </div>
    </form>
</body>
</html>
