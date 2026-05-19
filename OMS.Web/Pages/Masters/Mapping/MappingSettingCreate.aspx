<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MappingSettingCreate.aspx.vb" Inherits="OMS.Web.Pages.Masters.Mapping.MappingSettingCreate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>マッピングマスタ登録</title>
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
                <h1>マッピングマスタ登録</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnMappingList" runat="server" CssClass="btn-logout" Text="一覧へ" OnClick="btnMappingList_Click" />
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
            <div>
                <span class="label-title">フォルダ区分*</span>
                <asp:DropDownList ID="ddlFolderType" runat="server">
                    <asp:ListItem Text="内示" Value="1"></asp:ListItem>
                    <asp:ListItem Text="確定" Value="2"></asp:ListItem>
                    <asp:ListItem Text="納入指示" Value="3"></asp:ListItem>
                    <asp:ListItem Text="混在" Value="4"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <br />
            <div>
                <span class="label-title">バージョン</span>
                <asp:TextBox ID="txtVersion" runat="server" />
            </div>
            <div>
                <span class="label-title">ヘッダ行番号</span>
                <asp:TextBox ID="txtHeaderRowIndex" runat="server" />
            </div>
            <div>
                <span class="label-title">データ開始行番号</span>
                <asp:TextBox ID="txtDataStartRowIndex" runat="server" />
            </div>
            <div>
                <span class="label-title">シート名</span>
                <asp:TextBox ID="txtDefaultSheetName" runat="server" />
            </div>
            <br />

            <!-- ヘッダー -->
            <div class="process-header">
                <h1>詳細設定</h1>
            </div>

            <!-- マッピング明細マスタ一覧 -->
            <div class="data-list">
                <div class="data-grid-wrapper">
                    <asp:GridView ID="gvFieldMappingList" runat="server"
                        AutoGenerateColumns="False"
                        CssClass="data-grid"
                        BackColor="White"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="4" ForeColor="Black" GridLines="Both"
                        DataKeyNames="TempId,FieldMappingId"
                        EmptyDataText="該当データがありません。"
                        RowStyle-Wrap="False"
                        HeaderStyle-Wrap="False"
                        ShowFooter="True"
                        OnRowEditing="gvFieldMappingList_RowEditing"
                        OnRowCancelingEdit="gvFieldMappingList_RowCancelingEdit"
                        OnRowUpdating="gvFieldMappingList_RowUpdating"
                        OnRowDataBound="gvFieldMappingList_RowDataBound"
                        OnRowCommand="gvFieldMappingList_RowCommand"
                        OnRowDeleting="gvFieldMappingList_RowDeleting">

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

                            <asp:BoundField DataField="FieldMappingId" HeaderText="マッピング明細ID" Visible="false" ReadOnly="True" />
                            <asp:BoundField DataField="ProfileId" HeaderText="マッピングプロファイルID" Visible="false" ReadOnly="True" />

                            <asp:TemplateField HeaderText="フォーマット">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.FormatTypeMap, Eval("FormatType")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlFormatType" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlFormatType_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="マッピング先">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.TargetFieldMap, Eval("TargetField")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlTargetField" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlTargetField_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="列番号">
                                <ItemTemplate><%# Eval("SourceColumnIndex") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtSourceColumnIndex" runat="server" Text='<%# Bind("SourceColumnIndex") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtSourceColumnIndex_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="ヘッダ名">
                                <ItemTemplate><%# Eval("SourceHeaderName") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtSourceHeaderName" runat="server" Text='<%# Bind("SourceHeaderName") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtSourceHeaderName_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="シート名">
                                <ItemTemplate><%# Eval("SourceSheetName") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtSourceSheetName" runat="server" Text='<%# Bind("SourceSheetName") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtSourceSheetName_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="セルアドレス">
                                <ItemTemplate><%# Eval("SourceCellAddress") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtSourceCellAddress" runat="server" Text='<%# Bind("SourceCellAddress") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtSourceCellAddress_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="行選択タイプ">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.RowSelectorTypeMap, Eval("RowSelectorType")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlRowSelectorType" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlRowSelectorType_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="行選択値">
                                <ItemTemplate><%# Eval("RowSelectorValue") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtRowSelectorValue" runat="server" Text='<%# Bind("RowSelectorValue") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtRowSelectorValue_F" runat="server" />
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="データ型">
                                <ItemTemplate><%# OMS.Common.Constants.Lookup(OMS.Common.Constants.DataTypeMap, Eval("DataType")) %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlDataType" runat="server"></asp:DropDownList>
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:DropDownList ID="ddlDataType_F" runat="server"></asp:DropDownList>
                                </FooterTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="フォーマットパターン">
                                <ItemTemplate><%# Eval("FormatPattern") %></ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtFormatPattern" runat="server" Text='<%# Bind("FormatPattern") %>' />
                                </EditItemTemplate>
                                <FooterTemplate>
                                    <asp:TextBox ID="txtFormatPattern_F" runat="server" />
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

            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnCreateMappingSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="登録"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnCreateMappingSetting_Click" />
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
