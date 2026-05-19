<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="InfoTypeSetting.aspx.vb" Inherits="OMS.Web.Pages.Masters.InfoType.InfoTypeSetting" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>情報区分マスタ設定</title>
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
                <h1>情報区分マスタ設定</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                </div>
            </div>

            <!-- 入力 -->
            <div>
                <span class="label-title secondary">取引先コード*</span>
                <asp:TextBox ID="txtCustomerCode" runat="server" Enabled="False" />
                <asp:TextBox ID="txtCustomerName" runat="server" Width="400px" Enabled="False" />
            </div>
            <div>
                <span class="label-title secondary">PC</span>
                <asp:TextBox ID="txtProfitCenter" runat="server" Enabled="False" />
            </div>
            <div>
                <span class="label-title secondary">注文工場／担当者名</span>
                <asp:TextBox ID="txtCustomerUnitName" runat="server" Enabled="False" />
            </div>
            <br />
            <div>
                <span class="label-title">取引先情報区分*</span>
                <input type="text" id="txtCustomerInfoType" runat="server" class="field-required" required="required" />
            </div>
            <div>
                <span class="label-title">情報区分*</span>
                <asp:DropDownList ID="ddlInfoType" runat="server">
                    <asp:ListItem Text="新規" Value="I"></asp:ListItem>
                    <asp:ListItem Text="変更" Value="U"></asp:ListItem>
                    <asp:ListItem Text="取消" Value="D"></asp:ListItem>
                    <asp:ListItem Text="打切" Value="N"></asp:ListItem>
                </asp:DropDownList>
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
                <asp:Button ID="btnSaveInfoTypeSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="保存"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnSaveInfoTypeSetting_Click" />
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
