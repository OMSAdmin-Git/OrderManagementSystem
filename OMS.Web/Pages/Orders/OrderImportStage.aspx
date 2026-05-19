<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OrderImportStage.aspx.vb" Inherits="OMS.Web.Pages.Orders.OrderImportStage" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>受注ファイル取込準備</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Search.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/GridCheckAll.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/DropDownColor.js") %>"></script>

</head>
<body>
    <form id="form1" runat="server">
        <div class="process-container">

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>受注ファイル取込準備</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnOrderMenu" runat="server" CssClass="btn-back" Text="メニューへ" OnClick="btnOrderMenu_Click" />
                </div>
            </div>

            <!-- 検索条件 -->
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
                <div class="search-item button-item">
                    <asp:Button ID="btnSearchGv" runat="server" CssClass="btn-search" Text="検索" OnClick="btnSearchGv_Click" />
                    <asp:Button ID="btnDefaultGv" runat="server" CssClass="btn-search secondary" Text="クリア" OnClick="btnDefaultGv_Click" />
                </div>
            </div>

            <!-- 取込準備対象選択 -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvSelectCustomers" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="CustomerSettingId, CustomerCode, ProfitCenter, CustomerUnitId">
                        <Columns>
                            <asp:BoundField DataField="CustomerSettingId" HeaderText="取引先設定ID" Visible="false" />
                            <asp:BoundField DataField="CustomerCode" HeaderText="取引先コード" />
                            <asp:BoundField DataField="CustomerName" HeaderText="取引先名" />
                            <asp:BoundField DataField="ProfitCenter" HeaderText="PC" />
                            <asp:BoundField DataField="CustomerUnitId" HeaderText="注文工場／担当者ID" Visible="false" />
                            <asp:BoundField DataField="CustomerUnitName" HeaderText="注文工場／担当者名" />
                            <asp:TemplateField HeaderText="消込処理" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:DropDownList ID="ddlReconcileFlag"
                                        runat="server"
                                        CssClass="ddl-flag"
                                        SelectedValue='<%# Bind("ReconcileFlag") %>'>
                                        <asp:ListItem Text="する" Value="Y"></asp:ListItem>
                                        <asp:ListItem Text="しない" Value="N"></asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="消込対象" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:DropDownList ID="ddlFcstReconcileFlag"
                                        runat="server"
                                        CssClass="ddl-flag"
                                        SelectedValue='<%# Bind("FcstReconcileFlag") %>'>
                                        <asp:ListItem Text="内示を含める" Value="Y"></asp:ListItem>
                                        <asp:ListItem Text="内示を含めない" Value="N"></asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="処理対象">
                                <HeaderTemplate>
                                    <input type="checkbox" id="chkStageImportAll"
                                        onclick="OMS.Grid.toggleAll('<%= gvSelectCustomers.ClientID %>', this, 'chkStageImport')" />
                                    <label for="chkStageImportAll">処理対象</label>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkStageImport" runat="server" />
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
            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnStageImport" runat="server" CssClass="btn-asti btn-asti-process" Text="取込準備" OnClick="btnStageImport_Click" />
            </div>
            <!-- 実行結果 -->
            <div>
                <br />
                <asp:Label ID="lblResult" runat="server" ForeColor="Green" /><br />
                <asp:Label ID="lblError" runat="server" ForeColor="Red" />
            </div>
        </div>
    </form>
</body>
</html>
