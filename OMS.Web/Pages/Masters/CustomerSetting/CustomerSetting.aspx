<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CustomerSetting.aspx.vb" Inherits="OMS.Web.Pages.Masters.CustomerSetting.CustomerSetting" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>取引先マスタ設定</title>
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
                <h1>取引先マスタ設定</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                </div>
            </div>

            <!-- 入力 -->
            <div>
                <span class="label-title">取引先*</span>
                <input type="text" id="txtCustomerCode" list="lstCustomerCode" runat="server" class="field-required" required="required" />
                <datalist id="lstCustomerCode" runat="server"></datalist>
                <asp:TextBox ID="txtCustomerName" runat="server" Width="400px" Enabled="false" />
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
                <span class="label-title">生産管理担当者名*</span>
                <input type="text" id="txtProdMgmtUserName" list="lstProdMgmtUserName" runat="server" class="field-required" required="required" />
                <datalist id="lstProdMgmtUserName" runat="server"></datalist>
            </div>
            <br />
            <div>
                <span class="label-title">有効／無効*</span>
                <asp:DropDownList ID="ddlActiveFlag" runat="server">
                    <asp:ListItem Text="有効" Value="Y"></asp:ListItem>
                    <asp:ListItem Text="無効" Value="N"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <br />
            <div>
                <span class="label-title secondary">最終更新日時</span>
                <asp:TextBox ID="txtUpdatedAt" runat="server" Enabled="False" />
            </div>
            <div>
                <span class="label-title secondary">最終更新者</span>
                <asp:TextBox ID="txtUpdatedUserName" runat="server" Enabled="False" />
            </div>

            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnSaveCustomerSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="保存"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnSaveCustomerSetting_Click" />
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
