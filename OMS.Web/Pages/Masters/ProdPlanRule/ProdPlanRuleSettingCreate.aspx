<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ProdPlanRuleSettingCreate.aspx.vb" Inherits="OMS.Web.Pages.Masters.ProdPlanRule.ProdPlanRuleSettingCreate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>生産計画条件マスタ登録</title>
    <link href="~/Styles/Common.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Process.css" rel="stylesheet" type="text/css" />
    <link href="~/Styles/Input.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/PreventEnterSubmit.js") %>"></script>
    <script type="text/javascript" src="<%= ResolveUrl("~/Scripts/Custom/Validation.js") %>"></script>
</head>
<body>
    <form id="form1" runat="server" novalidate="novalidate">
        <div class="process-container">

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>生産計画条件マスタ登録</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnProdPlanRuleList" runat="server" CssClass="btn-logout" Text="一覧へ" OnClick="btnProdPlanRuleList_Click" />
                </div>
            </div>

            <!-- 入力 -->
            <div>
                <span class="label-title">取引先コード*</span>
                <input type="text" id="txtCustomerCode" list="lstCustomerCode" runat="server" class="field-required" required="required" />
                <datalist id="lstCustomerCode" runat="server"></datalist>
            </div>
            <div>
                <span class="label-title">PC</span>
                <input type="text" id="txtProfitCenter" list="lstProfitCenter" runat="server" />
                <datalist id="lstProfitCenter" runat="server"></datalist>
            </div>
            <div>
                <span class="label-title">注文工場／担当者名</span>
                <input type="text" id="txtCustomerUnitName" list="lstCustomerUnitName" runat="server" />
                <datalist id="lstCustomerUnitName" runat="server"></datalist>
            </div>
            <br />
            <div>
                <span class="label-title">分割フラグ*</span>
                <asp:DropDownList ID="ddlSplitFlag" runat="server" />
            </div>
            <div>
                <span class="label-title">分割パターンフラグ</span>
                <asp:DropDownList ID="ddlSplitCaseFlag" runat="server" />
            </div>
            <div>
                <span class="label-title">分割方法</span>
                <asp:DropDownList ID="ddlSplitMethodType" runat="server" />
            </div>
            <div>
                <span class="label-title">分割比</span>
                <asp:DropDownList ID="ddlSplitRationType" runat="server" />
            </div>
            <div>
                <span class="label-title">まるめ数</span>
                <asp:TextBox ID="txtSplitRoudingUnit" runat="server" />
            </div>
            <div>
                <span class="label-title">端数加算先</span>
                <asp:DropDownList ID="ddlCarryToType" runat="server" />
            </div>
            <div>
                <span class="label-title">分割開始区分</span>
                <asp:DropDownList ID="ddlSplitStartType" runat="server" />
            </div>
            <br />

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>分割パターン</h1>
            </div>

            <!-- 分割パターンマスタ一覧 -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvSplitCaseList" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="TempId,SplitCaseId"
                        EmptyDataText="該当データがありません。"
                        RowStyle-Wrap="False"
                        HeaderStyle-Wrap="False"
                        ShowFooter="True"
                        OnRowEditing="gvSplitCaseList_RowEditing"
                        OnRowCancelingEdit="gvSplitCaseList_RowCancelingEdit"
                        OnRowUpdating="gvSplitCaseList_RowUpdating"
                        OnRowDataBound="gvSplitCaseList_RowDataBound"
                        OnRowCommand="gvSplitCaseList_RowCommand"
                        OnRowDeleting="gvSplitCaseList_RowDeleting">

                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton runat="server" CommandName="Edit" Text="編集" />
                                    <asp:LinkButton runat="server" CommandName="Delete" Text="削除"
                                        OnClientClick="return confirm('この行を削除します。よろしいですか？');" />
                                </ItemTemplate>

                                <EditItemTemplate>
                                    <asp:LinkButton runat="server" CommandName="Update" Text="更新" />
                                    <asp:LinkButton runat="server" CommandName="Cancel" Text="取消" />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:LinkButton runat="server" CommandName="Insert" Text="追加" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:BoundField DataField="SplitCaseId" HeaderText="分割パターンID" Visible="false" ReadOnly="True" />
                            <asp:BoundField DataField="ProdPlanRuleId" HeaderText="生産計画条件ID" Visible="false" ReadOnly="True" />

                            <asp:TemplateField HeaderText="数量">
                                <ItemTemplate><%# Eval("Qty") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtQty" runat="server" Text='<%# Bind("Qty") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtQty_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="数量条件区分">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.QtyConditionTypeMap, Eval("QtyConditionType")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlQtyConditionType" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlQtyConditionType_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="分割方法">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.SplitMethodTypeMap, Eval("SplitMethodType")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlSplitMethodType" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlSplitMethodType_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="有効">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.ActiveFlagMap, Eval("ActiveFlag")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlActiveFlag" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlActiveFlag_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                                <ItemStyle HorizontalAlign="Center" />
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
            <div>
                <asp:Label ID="lblDetailResult" runat="server" ForeColor="Green" />
                <asp:Label ID="lblDetailError" runat="server" ForeColor="Red" CssClass="validation-summary" />
            </div>


            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnCreateProdPlanRuleSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="登録"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnCreateProdPlanRuleSetting_Click" />
            </div>

            <!-- 結果表示 -->
            <div>
                <br />
                <asp:Label ID="lblResult" runat="server" ForeColor="Green" />
                <asp:Label ID="lblError" runat="server" CssClass="validation-summary" />
            </div>

        </div>
    </form>
</body>
</html>
