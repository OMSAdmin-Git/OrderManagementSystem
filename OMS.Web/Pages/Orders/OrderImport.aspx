<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OrderImport.aspx.vb" Inherits="OMS.Web.Pages.Orders.OrderImport" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>受注ファイル取込</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Search.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/GridCheckAll.js") %>"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="process-container">

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>受注ファイル取込</h1>
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

            <!-- 取込準備済みファイル一覧 -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvImpFilesStage" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="ImpFileStageId, CustomerSettingId, CustomerCode, FolderType, FolderPath, FileName, StagedFolderPath, StagedFileName, ReconcileFlag, FcstReconcileFlag,ProfitCenter,CustomerUnitName">
                        <Columns>
                            <asp:BoundField DataField="ImpFileStageId" HeaderText="一時取込ファイルID" Visible="false" />
                            <asp:BoundField DataField="CustomerSettingId" HeaderText="取引先設定ID" Visible="false" />
                            <asp:BoundField DataField="CustomerCode" HeaderText="取引先コード" />
                            <asp:BoundField DataField="CustomerName" HeaderText="取引先名" />
                            <asp:BoundField DataField="ProfitCenter" HeaderText="PC" />
                            <asp:BoundField DataField="CustomerUnitId" HeaderText="注文工場／担当者ID" Visible="false" />
                            <asp:BoundField DataField="CustomerUnitName" HeaderText="注文工場／担当者名" />
                            <asp:TemplateField HeaderText="フォルダ区分">
                                <ItemTemplate>
                                    <%# OMS.Common.Utils.ToFolderTypeNameSafe(Eval("FolderType")) %>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="FolderPath" HeaderText="フォルダパス" Visible="false" />
                            <asp:BoundField DataField="FileName" HeaderText="ファイル名" />
                            <asp:BoundField DataField="StagedFolderPath" HeaderText="WORKフォルダパス" Visible="false" />
                            <asp:BoundField DataField="StagedFileName" HeaderText="WORKファイル名" Visible="false" />
                            <asp:BoundField DataField="ReconcileFlag" HeaderText="消込フラグ" Visible="false" />
                            <asp:BoundField DataField="FcstReconcileFlag" HeaderText="内示消込フラグ" Visible="false" />
                            <asp:TemplateField HeaderText="ハンド">
                                <HeaderTemplate>
                                    <input type="checkbox" id="chkHandFlagAll"
                                        onclick="OMS.Grid.toggleAll('<%= gvImpFilesStage.ClientID %>', this, 'chkHandFlag')" />
                                    <label for="chkHandFlagAll">ハンド</label>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkHandFlag" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="処理対象">
                                <HeaderTemplate>
                                    <input type="checkbox" id="chkOrderImportAll"
                                        onclick="OMS.Grid.toggleAll('<%= gvImpFilesStage.ClientID %>', this, 'chkOrderImport')" checked="checked" />
                                    <label for="chkOrderImportAll">処理対象</label>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkOrderImport" runat="server" Checked="True" />
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
                <%--<asp:Button ID="btnImportCancel" runat="server" CssClass="btn-asti btn-asti-process" Text="破棄" OnClick="btnImportCancel_Click" />--%>
                <asp:Button ID="btnImportCancel" runat="server" CssClass="btn-asti btn-asti-process" Text="破棄" OnClick="btnImportCancel_Click" />
                <asp:Button ID="btnImportFile" runat="server" CssClass="btn-asti btn-asti-process" Text="取込実行" OnClick="btnImportFile_Click" />
            </div>

            <!-- 結果表示 -->
            <div>
                <br />
                <asp:Label ID="lblImportResult" runat="server" ForeColor="Green" />
                <br />
                <asp:Label ID="lblImportError" runat="server" ForeColor="Red" />
            </div>

            <!-- 取込済み受注一覧 -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvImportOrder" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="OrderId, CustomerSettingId, OrderType">
                        <Columns>
                            <asp:BoundField DataField="OrderId" HeaderText="受注ID" Visible="false" />
                            <asp:BoundField DataField="CustomerSettingId" HeaderText="取引先設定ID" Visible="false" />
                            <asp:BoundField DataField="CustomerCode" HeaderText="取引先コード" />
                            <asp:BoundField DataField="CustomerName" HeaderText="取引先名" />
                            <asp:BoundField DataField="ProfitCenter" HeaderText="PC" />
                            <asp:BoundField DataField="CustomerUnitId" HeaderText="注文工場／担当者ID" Visible="false" />
                            <asp:BoundField DataField="CustomerUnitName" HeaderText="注文工場／担当者名" />
                            <asp:BoundField DataField="OrderType" HeaderText="受注区分" />
                            <asp:BoundField DataField="CustomerOrderNo" HeaderText="客先発注No" Visible="false" />
                            <asp:BoundField DataField="OrderDate" HeaderText="受注日" />                            
                            <asp:BoundField DataField="DueDate" HeaderText="希望納期" />
                            <asp:BoundField DataField="CustomerItemNo" HeaderText="客先品目No" />
                            <asp:BoundField DataField="ItemNo" HeaderText="品目No" />
                            <asp:BoundField DataField="DemandQty" HeaderText="需要数" />
                            <asp:BoundField DataField="CurrencyCode" HeaderText="通貨コード" />
                            <asp:BoundField DataField="ProductCode" HeaderText="製品コード" />
                            <asp:BoundField DataField="DeliveryCode" HeaderText="納入先コード" />
                            <asp:BoundField DataField="Remarks" HeaderText="コメント" />
                            <asp:BoundField DataField="ProratedType" HeaderText="分割区分" />
                            <asp:BoundField DataField="CustomerInfoType" HeaderText="取引先情報区分" />
                            <asp:BoundField DataField="InfoType" HeaderText="情報区分" />
                            <asp:BoundField DataField="SalfFcstFlag" HeaderText="自社予測フラグ" />
                            <asp:BoundField DataField="SalfFcstDeleteFlag" HeaderText="自社予測削除フラグ" />
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
                <asp:Button ID="btnCancelOrder" runat="server" CssClass="btn-asti btn-asti-process" Text="破棄" OnClick="btnCancelOrder_Click" />
                <asp:Button ID="btnSaveOrder" runat="server" CssClass="btn-asti btn-asti-process" Text="保存" OnClick="btnSaveOrder_Click" />
            </div>

            <!-- 結果表示 -->
            <div>
                <br />
                <asp:Label ID="lblSaveResult" runat="server" ForeColor="Green" />
                <br />
                <asp:Label ID="lblSaveError" runat="server" ForeColor="Red" />
            </div>

        </div>
    </form>
</body>
</html>
