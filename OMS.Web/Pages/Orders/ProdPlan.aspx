<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ProdPlan.aspx.vb" Inherits="OMS.Web.Pages.Orders.ProdPlan" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>生産計画</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Search.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/CustomFileUpload.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/GridCheckAll.js") %>"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="process-container">

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>生産計画</h1>
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

            <!-- 処理対象選択 -->
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
                            <asp:TemplateField HeaderText="処理対象">
                                <HeaderTemplate>
                                    <input type="checkbox" id="chkProdPlanAll"
                                        onclick="OMS.Grid.toggleAll('<%= gvSelectCustomers.ClientID %>', this, 'chkProdPlan')" checked="checked" />
                                    <label for="chkDueDateSettingAll">処理対象</label>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkProdPlan" runat="server" Checked="True" />
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
            <!-- アクションボタン R.Sagisaka Modified -->
            <div class="action-buttons">
                <asp:Button ID="btnProdPlan"           runat="server" CssClass="btn-asti btn-asti-process" Text="生産計画"  
                    OnClick="btnProdPlan_Click" />
                <asp:Button ID="btnExportProdPlanList" runat="server" CssClass="btn-asti secondary-excel " Text="Excel出力" 
                    OnClick="btnExportProdPlanList_Click" />

                <!-- <asp:Button ID="btnTest1" runat="server" CssClass="btn-asti btn-asti-process" Text="テスト1" OnClick="btnTest1_Click" />
                <asp:Button ID="btnTest2" runat="server" CssClass="btn-asti btn-asti-process" Text="テスト2" OnClick="btnTest2_Click" /> -->

                <!-- <asp:Button ID="btnImportProdPlanList" runat="server" CssClass="btn-asti secondary-excel " Text="Excel取込" OnClick="btnImportProdPlanList_Click" /> -->

                 <div class="custom-file-upload"> 
                    <asp:FileUpload ID="FileUpload1" runat="server"  />
                    <!-- <label for="<%= FileUpload1.ClientID %>" class="btn-asti secondary-excel ">Excel取込</label> -->
                    <asp:Button ID="Button1" runat="server" CssClass="btn-asti secondary-excel " Text="Excel取込" 
                        OnClick="btnImportProdPlanList_Click" />
                 </div> 
            </div>
            <!-- 生産計画 ボタン押下 データなし 継続確認 -->
           <div>
                <asp:Label ID="lblProdPlanContinueMessage" runat="server" Text="生産計画条件マスタに登録されていないデータを検出しました。​生産計画を続行しますか？​"></asp:Label>
                <!-- <label for="txtSearchCustomerUnitName">生産計画条件マスタに登録されていないデータを検出しました。​生産計画を続行しますか？​</label> -->
                <asp:Button ID="btnProdPlanOK" runat="server" CssClass="btn-asti btn-asti-process" Text="はい" CausesValidation="false"  
                    OnClick="btnProdPlanOK_Click"
                    OnClientClick="btnProdPlanOK.style.display='none';btnProdPlanNO.style.display='none';lblProdPlanContinueMessage.style.display='none';"
                    />
                <asp:Button ID="btnProdPlanNO"  runat="server" CssClass="btn-asti btn-asti-process" Text="いいえ" CausesValidation="false" OnClick="btnProdPlanNO_Click" />
                <!-- ダウンロードを裏で実行するための隠し iframe -->
                <iframe id = "downloadFrame" style="display:none;"></iframe>
           </div>
            <!-- 結果表示 -->
            <div>
                <br />
                <asp:Label ID="lblResult" runat="server" ForeColor="Green" />
                <asp:Label ID="lblError" runat="server" ForeColor="Red" />
           </div>
    </form>
</body>
</html>
