<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="InfoTypeList.aspx.vb" Inherits="OMS.Web.Pages.Masters.InfoType.InfoTypeList" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>情報区分マスタ</title>
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
                <h1>情報区分マスタ</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnMasterMenu" runat="server" CssClass="btn-logout" Text="メニューへ" OnClick="btnMasterMenu_Click" />
                </div>
            </div>

            <!-- 検索 -->
            <div class="search-section">
                <div class="search-item">
                    <label for="txtSearchCustomerCode">取引先コード</label>
                    <input type="text" id="txtSearchCustomerCode" list="lstSearchCustomerCode" runat="server" />
                    <datalist id="lstSearchCustomerCode" runat="server"></datalist>
                </div>
                <div class="search-item">
                    <label for="txtSearchCustomerName">取引先名</label>
                    <input type="text" id="txtSearchCustomerName" list="lstSearchCustomerName" runat="server" />
                    <datalist id="lstSearchCustomerName" runat="server"></datalist>
                </div>
                <div class="search-item">
                    <label for="txtSearchProfitCenter">PC</label>
                    <input type="text" id="txtSearchProfitCenter" list="lstSearchProfitCenter" runat="server" />
                    <datalist id="lstSearchProfitCenter" runat="server"></datalist>
                </div>
                <div class="search-item">
                    <label for="txtSearchCustomerUnitName">注文工場／担当者名</label>
                    <input type="text" id="txtSearchCustomerUnitName" list="lstSearchCustomerUnitName" runat="server" />
                    <datalist id="lstSearchCustomerUnitName" runat="server"></datalist>
                </div>
                <div class="search-item">
                    <label for="txtSearchCustomerInfoType">取引先情報区分</label>
                    <input type="text" id="txtSearchCustomerInfoType" runat="server" />
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

            <!-- 情報区分マスタ一覧-->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvInfoTypeList" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="InfoTypeId">
                        <Columns>
                            <asp:BoundField DataField="InfoTypeId" HeaderText="情報区分ID" Visible="false" />
                            <asp:BoundField DataField="CustomerSettingId" HeaderText="取引先設定ID" Visible="false" />
                            <asp:BoundField DataField="CustomerCode" HeaderText="取引先CD" />
                            <asp:BoundField DataField="CustomerName" HeaderText="取引先名" />
                            <asp:BoundField DataField="ProfitCenter" HeaderText="PC" />
                            <asp:BoundField DataField="CustomerUnitId" HeaderText="注文工場／担当者ID" Visible="false" />
                            <asp:BoundField DataField="CustomerUnitName" HeaderText="注文工場／担当者名" />
                            <asp:BoundField DataField="CustomerInfoType" HeaderText="取引先情報区分" />
                            <asp:TemplateField HeaderText="情報区分">
                                <ItemTemplate>
                                    <%# OMS.Common.Utils.ToInfoTypeNameSafe(Eval("InfoType")) %>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="ActiveFlag" HeaderText="有効" ItemStyle-HorizontalAlign="Center" />
                            <asp:TemplateField HeaderText="詳細">
                                <ItemTemplate>
                                    <asp:HyperLink ID="hlOpenSetting" runat="server"
                                        Text="開く"
                                        NavigateUrl='<%# "InfoTypeSetting.aspx?id1=" & System.Web.HttpUtility.UrlEncode(Eval("InfoTypeId").ToString()) %>'
                                        Target="_blank" />
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

                <!-- アクションボタン -->
                <div class="action-buttons">
                    <asp:Button ID="btnInfoTypeSettingCreate" runat="server" CssClass="btn-asti btn-asti-process" Text="新規登録" OnClick="btnInfoTypeSettingCreate_Click" />
                </div>

            </div>
        </div>
    </form>
</body>
</html>
