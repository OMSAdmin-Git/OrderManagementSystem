<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ImpRuleSettingCreate.aspx.vb" Inherits="OMS.Web.Pages.Masters.ImpRule.ImpRuleSettingCreate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>取込条件マスタ登録</title>
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
                <h1>取込条件マスタ登録</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnImpRuleList" runat="server" CssClass="btn-logout" Text="一覧へ" OnClick="btnImpRuleList_Click" />
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
                <span class="label-title">分割区分*</span>
                <asp:DropDownList ID="ddlProratedType" runat="server">
                    <asp:ListItem Text="日割り" Value="1"></asp:ListItem>
                    <asp:ListItem Text="日割り以外" Value="2"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">消込*</span>
                <asp:DropDownList ID="ddlReconcileFlag" runat="server">
                    <asp:ListItem Text="消込処理をする" Value="Y"></asp:ListItem>
                    <asp:ListItem Text="消込処理をしない" Value="N"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">内示消込*</span>
                <asp:DropDownList ID="ddlFcstReconcileFlag" runat="server">
                    <asp:ListItem Text="内示を消込対象に含める" Value="Y"></asp:ListItem>
                    <asp:ListItem Text="内示を消込対象に含めない" Value="N"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">消込条件*</span>
                <asp:DropDownList ID="ddlReconcileType" runat="server">
                    <asp:ListItem Text="順次" Value="1"></asp:ListItem>
                    <asp:ListItem Text="同月まで" Value="2"></asp:ListItem>
                    <asp:ListItem Text="同月内のみ" Value="3"></asp:ListItem>
                </asp:DropDownList>
            </div>

            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnCreateImpRuleSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="登録"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnCreateImpRuleSetting_Click" />
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
