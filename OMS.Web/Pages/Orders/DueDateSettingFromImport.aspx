<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DueDateSettingFromImport.aspx.vb" Inherits="OMS.Web.Pages.Orders.DueDateSettingFromImport" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>STRA納期設定</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Search.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/ModalMessage.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/GridCheckAll.js") %>"></script>

    <script type="text/javascript">
        function showErrorModal() {
            var modal = document.getElementById('<%= errorModalOverlay.ClientID %>');
            if (modal) modal.style.display = 'flex';
        }
        function closeErrorModal() {
            var modal = document.getElementById('<%= errorModalOverlay.ClientID %>');
            if (modal) modal.style.display = 'none';
        }
        function onButtonClick(btn) {
            var overlay = document.getElementById('loadingOverlay');
            if (overlay) overlay.style.display = 'flex';
            setTimeout(function () {
                btn.disabled = true;
            }, 50);
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="process-container">

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>STRA納期設定</h1>
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
                                    <input type="checkbox" id="chkDueDateSettingAll"
                                        onclick="OMS.Grid.toggleAll('<%= gvSelectCustomers.ClientID %>', this, 'chkDueDateSetting')" checked="checked" />
                                    <label for="chkDueDateSettingAll">処理対象</label>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkDueDateSetting" runat="server" Checked="True" />
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
                <asp:Button ID="btnDueDateSetting" runat="server" CssClass="btn-asti btn-asti-process" Text="納期設定" OnClick="btnDueDateSetting_Click" OnClientClick="return onButtonClick(this);"/>
                <asp:Button ID="btnExportDiffList" runat="server" CssClass="btn-asti secondary-excel " Text="差異リスト出力" OnClick="btnExportDiffList_Click" OnClientClick="return onButtonClick(this);"/>
                <!-- ダウンロードを裏で実行するための隠し iframe -->
                <iframe id = "downloadFrame" style="display:none;"></iframe>
            </div>
            <!-- 結果表示 -->
            <div>
                <br />
                <asp:Label ID="lblResult" runat="server" ForeColor="Green" />
                <asp:Label ID="lblError" runat="server" ForeColor="Red" />
            </div>

            <!--------------------------------------------->
            <!-- エラー表示用ポップアップ (Modal Dialog) -->
            <div id="errorModalOverlay" class="modal-overlay" runat="server">
                <div class="modal-dialog-custom">
                    <div class="modal-header-custom">
                        <h2>納期設定エラー一覧</h2>
                        <button type="button" class="modal-close-btn" onclick="closeErrorModal();">&times;</button>
                    </div>
                    <div class="modal-body-custom">
                        <p style="color: #c9302c; font-weight: bold; margin-top: 0; margin-bottom: 15px;">
                            受注取込後 納期設定中にエラーが発生しました。詳細は下記および各フォルダの「エラーリスト」フォルダ内のCSVをご確認ください。
                        </p>
                        <div style="max-height: 400px; overflow-y: auto; border: 1px solid #ddd;">
                            <asp:GridView ID="gvErrorList" runat="server"
                                AutoGenerateColumns="False"
                                CssClass="data-grid"
                                BackColor="White"
                                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                CellPadding="6" ForeColor="Black" GridLines="Both" Width="100%">
                                <Columns>
                                    <asp:TemplateField HeaderText="No" ItemStyle-Width="50px" ItemStyle-HorizontalAlign="Center">
                                        <ItemTemplate>
                                            <%# Container.DataItemIndex + 1 %>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="ErrorMessage" HeaderText="エラー内容" />
                                </Columns>
                                <HeaderStyle BackColor="#333333" Font-Bold="True" ForeColor="White" />
                                <RowStyle BackColor="#F7F7F7" />
                                <AlternatingRowStyle BackColor="White" />
                            </asp:GridView>
                        </div>
                    </div>
                    <div class="modal-footer-custom">
                        <button type="button" class="btn-cancel" onclick="closeErrorModal();">閉じる</button>
                    </div>
                </div>
            </div>
            <!-- 処理中ローディング表示 (Loading Overlay) -->
            <div id="loadingOverlay" class="modal-overlay">
                <div class="loading-dialog">
                    <div class="spinner"></div>
                    <div style="font-size: 16px; font-weight: 600; color: #333;">受注取込後 納期設定中... しばらくお待ちください</div>
                </div>
            </div>
            <!--------------------------------------------->
        </div>
    </form>
</body>
</html>
