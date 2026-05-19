'------------------------------------------------------------------------------
' <自動生成>
'     このコードはツールによって生成されました。
'
'     このファイルへの変更は、正しくない動作の原因になる可能性があり、
'     コードが再生成されるときに損失したりします。 
' </自動生成>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Namespace Pages.Orders

    Partial Public Class OrderImport

        '''<summary>
        '''form1 コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents form1 As Global.System.Web.UI.HtmlControls.HtmlForm

        '''<summary>
        '''lblUser コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblUser As Global.System.Web.UI.WebControls.Label

        '''<summary>
        '''btnOrderMenu コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnOrderMenu As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''txtSearchCustomerCode コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents txtSearchCustomerCode As Global.System.Web.UI.HtmlControls.HtmlInputText

        '''<summary>
        '''lstSearchCustomerCode コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lstSearchCustomerCode As Global.System.Web.UI.HtmlControls.HtmlGenericControl

        '''<summary>
        '''txtSearchCustomerName コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents txtSearchCustomerName As Global.System.Web.UI.HtmlControls.HtmlInputText

        '''<summary>
        '''lstSearchCustomerName コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lstSearchCustomerName As Global.System.Web.UI.HtmlControls.HtmlGenericControl

        '''<summary>
        '''txtSearchProfitCenter コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents txtSearchProfitCenter As Global.System.Web.UI.HtmlControls.HtmlInputText

        '''<summary>
        '''lstSearchProfitCenter コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lstSearchProfitCenter As Global.System.Web.UI.HtmlControls.HtmlGenericControl

        '''<summary>
        '''txtSearchCustomerUnitName コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents txtSearchCustomerUnitName As Global.System.Web.UI.HtmlControls.HtmlInputText

        '''<summary>
        '''lstSearchCustomerUnitName コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lstSearchCustomerUnitName As Global.System.Web.UI.HtmlControls.HtmlGenericControl

        '''<summary>
        '''btnSearchGv コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnSearchGv As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''btnDefaultGv コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnDefaultGv As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''gvImpFilesStage コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents gvImpFilesStage As Global.System.Web.UI.WebControls.GridView

        '''<summary>
        '''btnImportCancel コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnImportCancel As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''btnImportFile コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnImportFile As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''lblImportResult コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblImportResult As Global.System.Web.UI.WebControls.Label

        '''<summary>
        '''lblImportError コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblImportError As Global.System.Web.UI.WebControls.Label

        '''<summary>
        '''gvImportOrder コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents gvImportOrder As Global.System.Web.UI.WebControls.GridView

        '''<summary>
        '''btnCancelOrder コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnCancelOrder As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''btnSaveOrder コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnSaveOrder As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''lblSaveResult コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblSaveResult As Global.System.Web.UI.WebControls.Label

        '''<summary>
        '''lblSaveError コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblSaveError As Global.System.Web.UI.WebControls.Label
    End Class
End Namespace
