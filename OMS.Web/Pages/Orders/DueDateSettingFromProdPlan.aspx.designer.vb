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

    Partial Public Class DueDateSettingFromProdPlan

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
        '''gvSelectCustomers コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents gvSelectCustomers As Global.System.Web.UI.WebControls.GridView

        '''<summary>
        '''btnDueDateSetting コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnDueDateSetting As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''btnExportDiffList コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents btnExportDiffList As Global.System.Web.UI.WebControls.Button

        '''<summary>
        '''lblResult コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblResult As Global.System.Web.UI.WebControls.Label

        '''<summary>
        '''lblError コントロール。
        '''</summary>
        '''<remarks>
        '''自動生成されたフィールド。
        '''変更するには、フィールドの宣言をデザイナー ファイルから分離コード ファイルに移動します。
        '''</remarks>
        Protected WithEvents lblError As Global.System.Web.UI.WebControls.Label
    End Class
End Namespace
