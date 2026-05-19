<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FileSetting.aspx.vb" Inherits="OMS.Web.Pages.Masters.File.FileSetting" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>ファイルマスタ設定</title>
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
                <h1>ファイルマスタ設定</h1>
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
            <div>
                <span class="label-title secondary">フォルダ区分*</span>
                <asp:TextBox ID="txtFolderType" runat="server" Enabled="False" />
            </div>
            <br />
            <div>
                <span class="label-title">フォーマット*</span>
                <asp:DropDownList ID="ddlFormatType" runat="server">
                    <asp:ListItem Text="リスト" Value="LIST"></asp:ListItem>
                    <asp:ListItem Text="マトリックス" Value="MATRIX"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">ファイル形式*</span>
                <asp:DropDownList ID="ddlFileType" runat="server">
                    <asp:ListItem Text="CSV" Value="CSV"></asp:ListItem>
                    <asp:ListItem Text="TSV" Value="TSV"></asp:ListItem>
                    <asp:ListItem Text="固定長" Value="FIXED"></asp:ListItem>
                    <asp:ListItem Text="Excel" Value="EXCEL"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">区切り文字*</span>
                <asp:DropDownList ID="ddlDelimiter" runat="server">
                    <asp:ListItem Text="コンマ(,)" Value="COMMA"></asp:ListItem>
                    <asp:ListItem Text="タブ" Value="TAB"></asp:ListItem>
                    <asp:ListItem Text="セミコロン(;)" Value="SEMICOLON"></asp:ListItem>
                    <asp:ListItem Text="コロン(:)" Value="COLON"></asp:ListItem>
                    <asp:ListItem Text="パイプ(|)" Value="PIPE"></asp:ListItem>
                    <asp:ListItem Text="スペース( )" Value="SPACE"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">囲み文字*</span>
                <asp:DropDownList ID="ddlEnclosure" runat="server">
                    <asp:ListItem Text="ダブルクォーテーション(&quot;)" Value="D_QUOTE"></asp:ListItem>
                    <asp:ListItem Text="シングルクォーテーション(')" Value="S_QUOTE"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">ヘッダー*</span>
                <asp:DropDownList ID="ddlHeaderFlag" runat="server">
                    <asp:ListItem Text="あり" Value="Y"></asp:ListItem>
                    <asp:ListItem Text="なし" Value="N"></asp:ListItem>
                </asp:DropDownList>
            </div>
            <div>
                <span class="label-title">文字コード*</span>
                <asp:DropDownList ID="ddlCharset" runat="server">
                    <asp:ListItem Text="UTF-8" Value="UTF8"></asp:ListItem>
                    <asp:ListItem Text="Shift_JIS" Value="Shift_JIS"></asp:ListItem>
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
                <asp:Button ID="btnSaveFileSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="保存"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnSaveFileSetting_Click" />
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
