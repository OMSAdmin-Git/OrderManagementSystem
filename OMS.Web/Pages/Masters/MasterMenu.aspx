<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MasterMenu.aspx.vb" Inherits="OMS.Web.Pages.Masters.MasterMenu" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>マスタ管理メニュー</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Menu.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>マスタ管理</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                <asp:Button ID="btnTopMenu" runat="server" CssClass="btn-logout" Text="トップへ" OnClick="btnTopMenu_Click" />
                </div>
            </div>
            <div class="menu-grid">
                <button id="btnCustomerList" runat="server" class="menu-tile btn-asti" onserverclick="btnCustomerList_ServerClick">
                    <div class="tile-title">取引先マスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnCustomerUnitList" runat="server" class="menu-tile btn-asti" onserverclick="btnCustomerUnitList_ServerClick">
                    <div class="tile-title">注文工場／担当者マスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnFolderList" runat="server" class="menu-tile btn-asti" onserverclick="btnFolderList_ServerClick">
                    <div class="tile-title">フォルダマスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnFileList" runat="server" class="menu-tile btn-asti" onserverclick="btnFileList_ServerClick">
                    <div class="tile-title">ファイルマスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnInfoTypeList" runat="server" class="menu-tile btn-asti" onserverclick="btnInfoTypeList_ServerClick">
                    <div class="tile-title">情報区分マスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnMappingList" runat="server" class="menu-tile btn-asti" onserverclick="btnMappingList_ServerClick">
                    <div class="tile-title">マッピングマスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnImpRuleList" runat="server" class="menu-tile btn-asti" onserverclick="btnImpRuleList_ServerClick">
                    <div class="tile-title">取込条件マスタ</div>
                    <div class="tile-descList">マスタデータ登録・編集</div>
                </button>
                <button id="btnProdPlanRuleList" runat="server" class="menu-tile btn-asti" onserverclick="btnProdPlanRuleList_ServerClick">
                    <div class="tile-title">生産計画条件マスタ</div>
                    <div class="tile-desc">マスタデータ登録・編集</div>
                </button>
                <button id="btnUserList" runat="server" class="menu-tile btn-asti" onserverclick="btnUserList_ServerClick">
                    <div class="tile-title">ユーザーマスタ</div>
                    <div class="tile-desc">マスタデータ編集</div>
                </button>
            </div>
        </div>
    </form>
</body>
</html>
