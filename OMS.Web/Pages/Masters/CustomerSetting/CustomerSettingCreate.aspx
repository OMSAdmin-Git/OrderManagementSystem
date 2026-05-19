<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CustomerSettingCreate.aspx.vb" Inherits="OMS.Web.Pages.Masters.CustomerSetting.CustomerSettingCreate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>取引先マスタ登録</title>
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
                <h1>取引先マスタ登録</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnCustomerList" runat="server" CssClass="btn-logout" Text="一覧へ" OnClick="btnCustomerList_Click" />
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
                <span class="label-title">生産管理担当者名*</span>
                <input type="text" id="txtProdMgmtUserName" list="lstProdMgmtUserName" runat="server" class="field-required" required="required" />
                <datalist id="lstProdMgmtUserName" runat="server"></datalist>
            </div>

            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnCreateCustomerSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="登録"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnCreateCustomerSetting_Click" />
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
