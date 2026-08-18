<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SuzukiSpiritsConversionList.aspx.vb" Inherits="OMS.Web.Pages.Masters.SuzukiSpiritsConversion.SuzukiSpiritsConversionList" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>スズキSPIRITS納入ホーム変換マスタ</title>
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
                <h1>スズキSPIRITS納入ホーム変換マスタ</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnMasterMenu" runat="server" CssClass="btn-logout" Text="メニューへ" OnClick="btnMasterMenu_Click" />
                </div>
            </div>

            <!-- 検索 -->
            <div class="search-section">
                <div class="search-item">
                    <label for="txtSearchDeliveryCodeOrder">納入指示納入先コード</label>
                    <input type="text" id="txtSearchDeliveryCodeOrder" list="lstSearchDeliveryCodeOrder" runat="server" />
                    <datalist id="lstSearchDeliveryCodeOrder" runat="server"></datalist>
                </div>
                <div class="search-item">
                    <label for="txtSearchDeliveryCodePlan">内示納入先コード</label>
                    <input type="text" id="txtSearchDeliveryCodePlan" list="lstSearchDeliveryCodePlan" runat="server" />
                    <datalist id="lstSearchDeliveryCodePlan" runat="server"></datalist>
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

            <!-- スズキSPIRITS納入ホーム変換マスタ -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvSuzukiSpiritsConversionList" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="ConversionId">
                        <Columns>
                            <asp:BoundField DataField="ConversionId" HeaderText="変換ID" Visible="false" />
                            <asp:BoundField DataField="deliveryCodeOrder" HeaderText="納入指示納入先コード" />
                            <asp:BoundField DataField="deliveryCodePlan" HeaderText="内示納入先コード" />
                            <asp:BoundField DataField="ActiveFlag" HeaderText="有効" ItemStyle-HorizontalAlign="Center" />
                            <asp:TemplateField HeaderText="詳細">
                                <ItemTemplate>
                                    <asp:HyperLink ID="hlOpenSetting" runat="server"
                                        Text="開く"
                                        NavigateUrl='<%# "SuzukiSpiritsConversionSetting.aspx?id1=" & System.Web.HttpUtility.UrlEncode(Eval("ConversionId").ToString()) %>'
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
                    <asp:Button ID="btnSuzukiSpiritsConversionSettingCreate" runat="server" CssClass="btn-asti btn-asti-process" Text="新規登録" OnClick="btnSuzukiSpiritsConversionSettingCreate_Click" />
                </div>

            </div>
        </div>
    </form>
</body>
</html>
