<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SuzukiSpiritsConversionSetting.aspx.vb" Inherits="OMS.Web.Pages.Masters.SuzukiSpiritsConversion.SuzukiSpiritsConversionSetting" %>


<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>スズキSPIRITS納入ホーム変換マスタ設定</title>
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
                <h1>スズキSPIRITS納入ホーム変換マスタ設定</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                </div>
            </div>

            <!-- 入力 -->
            <div>
                <span class="label-title secondary">納入指示納入先コード*</span>
                <asp:TextBox ID="txtDeliveryCodeOrder" runat="server" Enabled="False" />
            </div>
            <br />
             <div>
                 <span class="label-title">内示納入先コード*</span>
                 <input type="text" id="txtDeliveryCodePlan" list="lstDeliveryCodePlan" runat="server" class="field-required" required="required" />
                 <datalist id="lstDeliveryCodePlan" runat="server"></datalist>
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
                <asp:Button ID="btnSaveSuzukiSpiritsConversionSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="保存"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnSaveSuzukiSpiritsConversionSetting_Click" />
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
