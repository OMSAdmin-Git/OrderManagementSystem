<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FileSettingCreate.aspx.vb" Inherits="OMS.Web.Pages.Masters.File.FileSettingCreate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="ja">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>ファイルマスタ登録</title>
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
                <h1>ファイルマスタ登録</h1>
                <div class="user-info">
                    <asp:Label ID="lblUser" runat="server" Text="ようこそ"></asp:Label>
                    &nbsp;               
                    <asp:Button ID="btnFileList" runat="server" CssClass="btn-logout" Text="一覧へ" OnClick="btnFileList_Click" />
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

            <!-- アクションボタン -->
            <div class="action-buttons">
                <asp:Button ID="btnCreateFileSetting" runat="server"
                    CssClass="btn-asti btn-asti-process"
                    Text="登録"
                    OnClientClick="return validateRequiredFromBtn(this);"
                    OnClick="btnCreateFileSetting_Click" />
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
