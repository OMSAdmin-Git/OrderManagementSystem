<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OrderMenu.aspx.vb" Inherits="OMS.Web.Pages.Orders.OrderMenu" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>受注管理メニュー</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Menu.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>受注取込</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnTopMenu" runat="server" CssClass="btn-back" Text="トップへ" OnClick="btnTopMenu_Click" />
                </div>
            </div>
            <div class="menu-grid">
                <button id="btnOrderImportStage" runat="server" class="menu-tile btn-asti" onserverclick="btnOrderImportStage_ServerClick">
                    <div class="tile-title">受注ファイル取込準備</div>
                    <div class="tile-desc">ファイル検索</div>
                </button>
                <button id="btnOrderImport" runat="server" class="menu-tile btn-asti" onserverclick="btnOrderImport_ServerClick">
                    <div class="tile-title">受注ファイル取込</div>
                    <div class="tile-desc">取引先CSV・汎用マトリックス</div>
                </button>
                <button id="btnDueSetFromImport" runat="server" class="menu-tile btn-asti" onserverclick="btnDueSetFromImport_ServerClick">
                    <div class="tile-title">STRA納期設定</div>
                    <div class="tile-desc">受注差異リスト</div>
                </button>
            </div>
        </div>
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>生産計画</h1>
            </div>
            <div class="menu-grid">
                <button id="btnProdPlan" runat="server" class="menu-tile btn-asti" onserverclick="btnProdPlan_ServerClick">
                    <div class="tile-title">生産計画</div>
                    <div class="tile-desc">数量丸め・日割り・分割等</div>
                </button>
                <button id="btnDueSetFromProdPlan" runat="server" class="menu-tile btn-asti" onserverclick="btnDueSetFromProdPlan_ServerClick">
                    <div class="tile-title">STRA納期設定</div>
                    <div class="tile-desc">生産計画差異リスト</div>
                </button>
                <button id="Button1" runat="server" class="menu-tile btn-asti empty-menu-button" enabled="false">
                    <div class="tile-title"></div>
                    <div class="tile-desc"></div>
                </button>
            </div>
        </div>
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>受注出力</h1>
            </div>
            <div class="menu-grid">
                <button id="btnOrderExport" runat="server" class="menu-tile btn-asti" onserverclick="btnOrderExport_ServerClick">
                    <div class="tile-title">受注ファイル出力</div>
                    <div class="tile-desc">基幹システム用CSV</div>
                </button>
                <button id="Button2" runat="server" class="menu-tile btn-asti empty-menu-button" enabled="false">
                    <div class="tile-title"></div>
                    <div class="tile-desc"></div>
                </button>
                <button id="Button3" runat="server" class="menu-tile btn-asti empty-menu-button" enabled="false">
                    <div class="tile-title"></div>
                    <div class="tile-desc"></div>
                </button>
            </div>
        </div>
        <!--
        <div class="menu-wrapper">
            <div class="menu-header">
                <h1>社内調整</h1>
            </div>
            <div class="menu-grid">
                <button id="btnSelfFcst" runat="server" class="menu-tile btn-asti" onserverclick="btnSelfFcst_ServerClick">
                    <div class="tile-title">内示調整</div>
                    <div class="tile-desc">追加内示閲覧・削除</div>
                </button>
                <button id="Button5" runat="server" class="menu-tile btn-asti empty-menu-button" enabled="false">
                    <div class="tile-title"></div>
                    <div class="tile-desc"></div>
                </button>
                <button id="Button6" runat="server" class="menu-tile btn-asti empty-menu-button" enabled="false">
                    <div class="tile-title"></div>
                    <div class="tile-desc"></div>
                </button>
            </div>
        </div>
        -->
    </form>
</body>
</html>
