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
    <style type="text/css">
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            z-index: 9999;
            justify-content: center;
            align-items: center;
        }
        .modal-dialog-custom {
            background-color: #fff;
            width: 85%;
            max-width: 950px;
            max-height: 85vh;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            animation: modalFadeIn 0.2s ease-out;
        }
        @keyframes modalFadeIn {
            from { opacity: 0; transform: translateY(-20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .modal-header-custom {
            background-color: #c9302c;
            color: white;
            padding: 14px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .modal-header-custom h2 {
            margin: 0;
            font-size: 18px;
            font-weight: 600;
        }
        .modal-close-btn {
            background: none;
            border: none;
            color: white;
            font-size: 26px;
            cursor: pointer;
            line-height: 1;
            padding: 0;
        }
        .modal-close-btn:hover {
            opacity: 0.8;
        }
        .modal-body-custom {
            padding: 20px;
            overflow-y: auto;
            flex: 1;
        }
        .modal-footer-custom {
            padding: 12px 20px;
            border-top: 1px solid #e5e5e5;
            display: flex;
            justify-content: flex-end;
            background-color: #f9f9f9;
        }
    </style>
    <script type="text/javascript">
        function showErrorModal() {
            var modal = document.getElementById('<%= errorModalOverlay.ClientID %>');
            if (modal) modal.style.display = 'flex';
        }
        function closeErrorModal() {
            var modal = document.getElementById('<%= errorModalOverlay.ClientID %>');
            if (modal) modal.style.display = 'none';
        }
    </script>

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
                        DataKeyNames="CustomerSettingId, CustomerCode, ProfitCenter, CustomerUnitId, SpProcessType">
                        <Columns>
                            <asp:BoundField DataField="CustomerSettingId" HeaderText="取引先設定ID" Visible="false" />
                            <asp:BoundField DataField="CustomerCode" HeaderText="取引先コード" />
                            <asp:BoundField DataField="CustomerName" HeaderText="取引先名" />
                            <asp:BoundField DataField="ProfitCenter" HeaderText="PC" />
                            <asp:BoundField DataField="CustomerUnitId" HeaderText="注文工場／担当者ID" Visible="false" />
                            <asp:BoundField DataField="CustomerUnitName" HeaderText="注文工場／担当者名" />
                            <asp:BoundField DataField="SpProcessType" HeaderText="特殊加工区分(Debug用)"  />
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

            <!-- エラー表示用ポップアップ (Modal Dialog) -->
            <div id="errorModalOverlay" class="modal-overlay" runat="server">
                <div class="modal-dialog-custom">
                    <div class="modal-header-custom">
                        <h2>取込エラー一覧</h2>
                        <button type="button" class="modal-close-btn" onclick="closeErrorModal();">&times;</button>
                    </div>
                    <div class="modal-body-custom">
                        <p style="color: #c9302c; font-weight: bold; margin-top: 0; margin-bottom: 15px;">
                            取込前処理中にエラーが発生しました。詳細は下記および各フォルダの「エラーリスト」フォルダ内のCSVをご確認ください。
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
                        <button type="button" class="btn-asti btn-asti-process" onclick="closeErrorModal();">閉じる</button>
                    </div>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
