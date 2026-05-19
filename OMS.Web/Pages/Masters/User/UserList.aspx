<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="UserList.aspx.vb" Inherits="OMS.Web.UserList" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>ユーザーマスタ</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Search.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="process-container">

    <!-- ヘッダー -->
    <div class="process-header">
        <h1>ユーザーマスタ</h1>
        <div class="user-info">
            <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
            &nbsp;               
            <asp:Button ID="btnMasterMenu" runat="server" CssClass="btn-logout" Text="メニューへ" OnClick="btnMasterMenu_Click" />
        </div>
    </div>

    <!-- 検索 -->
    <div class="search-section">
        <div class="search-item">
            <label for="txtSearchUserId">ユーザーID</label>
            <input type="text" id="txtSearchUserId" list="lstSearchUserId" runat="server" />
            <datalist id="lstSearchUserId" runat="server"></datalist>
        </div>
        <div class="search-item">
            <label for="txtSearchUserName">ユーザー名</label>
            <input type="text" id="txtSearchUserName" list="lstSearchUserName" runat="server" />
            <datalist id="lstSearchUserName" runat="server"></datalist>
        </div>
    </div>
    <div class="search-section">
        <div class="search-item">
            <input type="checkbox" id="chkSearchActiveOnly" runat="server" checked="checked" />
            <label for="chkSearchActiveOnly">有効のみ</label>
        </div>
    </div>
    <div class="action-buttons">
        <div class="search-item button-item">
            <button id="btnSearchGv" runat="server" class="btn-search " onserverclick="btnSearchGv_ServerClick">検索</button>
            <button id="btnDefaultGv" runat="server" class="btn-search secondary" onserverclick="btnDefaultGv_ServerClick">クリア</button>
        </div>
    </div>
    <!-- ユーザーマスタ一覧 -->
    <div class="data-list">
        <div class="data-grid-wrapper">
            <asp:GridView ID="gvUserList" runat="server"
                AutoGenerateColumns="False"
                OnRowCommand="gvUserList_RowCommand"
                CssClass="data-grid"
                BackColor="White"
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                CellPadding="4" ForeColor="Black" GridLines="Both"
                DataKeyNames="ProdMgmtUserId">
                <Columns>
                    <asp:BoundField DataField="ProdMgmtUserId" HeaderText="ユーザーID" />
                    <asp:BoundField DataField="ProdMgmtUserName" HeaderText="ユーザー名" />
                    <asp:TemplateField HeaderText="権限レベル">
                        <ItemTemplate>
                            <%# OMS.Common.Utils.ToAuthorityLevelNameSafe(Eval("AuthorityLevel")) %>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="UpdatedAt" HeaderText="更新日時" />
                    <asp:BoundField DataField="ActiveFlag" HeaderText="有効" ItemStyle-HorizontalAlign="Center" />
                    <asp:TemplateField HeaderText="パスワード">
                        <ItemTemplate>
                            <asp:Button ID="btnReset" runat="server" Text="リセット" CommandName="PassReset" CommandArgument='<%# Eval("ProdMgmtUserId") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                </Columns>
                <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
                <HeaderStyle BackColor="#333333" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
                <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F7F7F7" />
                <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
                <SortedDescendingCellStyle BackColor="#E5E5E5" />
                <SortedDescendingHeaderStyle BackColor="#242121" />
            </asp:GridView>
        </div>
    </div>
    <!-- 結果表示 -->
    <div>
        <br />
        <asp:Label ID="lblResult" runat="server" ForeColor="Green" />
        <asp:Label ID="lblError" runat="server" CssClass="validation-summary" />
    </div>
    </form>
</body>
</html>
