Imports System.Configuration
Imports System.Data
Imports System.Text
Imports System.Web
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.SqlServer
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data
    Public Class OrderProductionPlanExcelFile

        ''' <summary>
        ''' (差異リスト:内示) 
        ''' Index(Column), Cell width, Title
        ''' </summary>
        Private OrdersTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
            (1, 15, "受注ID"),                'ORDER_ID                NUMBER(10,0)                No
            (2, 15, "取引先設定ID"),          'CUSTOMER_SETTING_ID     NUMBER(10,0)                No
            (3, 15, "取引先コード"),          'CUSTOMER_CODE           VARCHAR2(25 BYTE)           No
            (4, 15, "請求先"),                'BILLING_TO              VARCHAR2(25 BYTE)           Yes
            (5, 15, "客先発注No"),            'CUSTOMER_ORDER_NO       VARCHAR2(40 BYTE)           No
            (6, 15, "需要ステイタス"),        'DEMAND_STATUS           CHAR(1 BYTE)                Yes
            (7, 15, "出荷先"),                'SHIP_TO                 VARCHAR2(25 BYTE)           No
            (8, 15, "受注日"),                'ORDER_DATE              DATE                        Yes
            (9, 15, "希望納期"),              'DUE_DATE                DATE                        No
            (10, 15, "出荷予定日"),           'SHIP_SCHEDULED_DATE     DATE                        Yes
            (11, 15, "客先品目No"),           'CUSTOMER_ITEM_NO        VARCHAR2(20 BYTE)           No
            (12, 15, "品目No"),               'ITEM_NO                 VARCHAR2(20 BYTE)           Yes
            (13, 15, "需要数"),               'DEMAND_QTY              NUMBER(10,0)                No
            (14, 15, "需要単位"),             'DEMAND_UNIT             VARCHAR2(4 BYTE)            Yes
            (15, 15, "通貨コード"),           'CURRENCY_CODE           VARCHAR2(3 BYTE)            Yes
            (16, 15, "出荷在庫場所"),         'SHIP_STOCK_LOCATION     VARCHAR2(25 BYTE)           Yes
            (17, 15, "会社ID"),               'COMPANY_ID              VARCHAR2(25 BYTE)           Yes
            (18, 15, "製品コード"),           'PRODUCT_CODE            VARCHAR2(20 BYTE)           Yes
            (19, 15, "請求基準"),             'BILLING_STANDARD        VARCHAR2(3 BYTE)            Yes
            (20, 15, "出荷プロセスタイプ"),   'SHIP_PROCESS_TYPE       CHAR(1 BYTE)                Yes
            (21, 15, "納入指示フラグ"),       'DELIVERY_INSTR_FLAG     CHAR(1 BYTE)                Yes
            (22, 15, "受注番号"),             'ORDER_NO                VARCHAR2(45 BYTE)           Yes
            (23, 15, "コメント"),             'REMARKS                 VARCHAR2(45 BYTE)           Yes
            (24, 15, "納入先コード"),         'DELIVERY_CODE           VARCHAR2(20 BYTE)           Yes
            (25, 15, "受注時刻"),             'ORDER_TIME              NUMBER(18,6)                Yes
            (26, 15, "売上単価"),             'SALES_UNIT_PRICE        NUMBER(18,6)                Yes
            (27, 15, "納入時間"),             'DELIVERY_TIME           NUMBER(18,6)                Yes
            (28, 15, "使用先"),               'USAGE_LOCATION          VARCHAR2(45 BYTE)           Yes
            (29, 15, "累計出荷数"),           'TOTAL_SHIP_QTY          NUMBER(18,6)                Yes
            (30, 15, "生産区分"),             'PRODUCTION_CATEGORY     VARCHAR2(45 BYTE)           Yes
            (31, 15, "文字 2ｹﾀ"),             'CHAR_2                  VARCHAR2(45 BYTE)           Yes
            (32, 15, "容器番号"),             'CONTAINER_NO            VARCHAR2(45 BYTE)           Yes
            (33, 15, "文字 3ｹﾀ"),             'CHAR_3                  VARCHAR2(45 BYTE)           Yes
            (34, 15, "文字 4ｹﾀ"),             'CHAR_4                  VARCHAR2(45 BYTE)           Yes
            (35, 15, "文字 4ｹﾀ"),             'CHAR_4_2                VARCHAR2(45 BYTE)           Yes
            (36, 15, "文字 5ｹﾀ"),             'CHAR_5                  VARCHAR2(45 BYTE)           Yes
            (37, 15, "文字 5ｹﾀ"),             'CHAR_5_2                VARCHAR2(45 BYTE)           Yes
            (38, 15, "文字 6ｹﾀ"),             'CHAR_6                  VARCHAR2(45 BYTE)           Yes
            (39, 15, "発注理由"),             'ORDER_REASON            VARCHAR2(45 BYTE)           Yes
            (40, 15, "容器収容数"),           'CONTAINER_CAPACITY      NUMBER(18,6)                Yes
            (41, 15, "得意先 ﾛｯﾄNO"),         'CUSTOMER_LOT_NO         VARCHAR2(45 BYTE)           Yes
            (42, 15, "初品区分"),             'INITIAL_FLAG            VARCHAR2(45 BYTE)           Yes
            (43, 15, "出荷日"),               'SHIP_DATE               DATE                        Yes
            (44, 15, "文字 50ｹﾀ"),            'CHAR_50                 VARCHAR2(60 BYTE)           Yes
            (45, 15, "輸送方法"),             'TRANSPORT_METHOD        VARCHAR2(3 BYTE)            Yes
            (46, 15, "出荷計画日"),           'SHIP_PLAN_DATE          DATE                        Yes
            (47, 15, "客先発注No行番号"),     'CUSTOMER_ORDER_LINE_NO  VARCHAR2(2 BYTE)            Yes
            (48, 15, "日割前受注数"),         'PRE_DAILY_ORDER_QTY     NUMBER(18,6)                Yes
            (49, 15, "日割前納期"),           'PRE_DAILY_DELIVERY_DATE DATE                        Yes
            (50, 15, "取込ファイルID"),       'IMP_FILE_ID             NUMBER(10,0)                No
            (51, 15, "受注区分"),             'ORDER_TYPE              NUMBER(1,0)                 No
            (52, 15, "分割区分"),             'PRORATED_TYPE           NUMBER(1,0)                 No
            (53, 15, "取引先情報区分"),       'CUSTOMER_INFO_TYPE      VARCHAR2(50 BYTE)           Yes
            (54, 15, "情報区分"),             'INFO_TYPE               CHAR(1 BYTE)                Yes
            (55, 15, "自社予測フラグ"),       'SELF_FCST_FLAG          CHAR(1 BYTE)                Yes
            (56, 15, "自社予測削除フラグ"),   'SELF_FCST_DELETE_FLAG   CHAR(1 BYTE)                Yes
            (57, 15, "消込条件区分"),         'RECONCILE_TYPE          NUMBER(1,0)                 Yes
            (58, 15, "取込実行ID"),           'IMP_RUN_ID              VARCHAR2(36 BYTE)           No
            (59, 15, "ステータス"),           'STATUS                  VARCHAR2(20 BYTE)           Yes
            (60, 15, "有効フラグ"),           'ACTIVE_FLAG             CHAR(1 BYTE)                No
            (61, 15, "登録日時"),             'CREATED_AT              DATE                        No
            (62, 15, "登録ユーザーID"),       'CREATED_USER_ID         VARCHAR2(9 BYTE)            No
            (63, 15, "登録プログラムID"),     'CREATED_PG_ID           VARCHAR2(150 BYTE)          No
            (64, 15, "更新日時"),             'UPDATED_AT              DATE                        No
            (65, 15, "更新ユーザーID"),       'UPDATED_USER_ID         VARCHAR2(9 BYTE)            No
            (66, 15, "更新プログラムID")      'UPDATED_PG_ID           VARCHAR2(150 BYTE)          No
        }
        ''' <summary>
        ''' (差異リスト:内示) 
        ''' Index(Column), Cell width, Title
        ''' </summary>
        Private OrdersStageTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
            (1, 15, "受注ID"),                'ORDER_ID                NUMBER(10,0)                No
            (2, 15, "取引先設定ID"),          'CUSTOMER_SETTING_ID     NUMBER(10,0)                No
            (3, 15, "取引先コード"),          'CUSTOMER_CODE           VARCHAR2(25 BYTE)           No
            (4, 15, "請求先"),                'BILLING_TO              VARCHAR2(25 BYTE)           Yes
            (5, 15, "客先発注No"),            'CUSTOMER_ORDER_NO       VARCHAR2(40 BYTE)           No
            (6, 15, "需要ステイタス"),        'DEMAND_STATUS           CHAR(1 BYTE)                Yes
            (7, 15, "出荷先"),                'SHIP_TO                 VARCHAR2(25 BYTE)           No
            (8, 15, "受注日"),                'ORDER_DATE              DATE                        Yes
            (9, 15, "希望納期"),              'DUE_DATE                DATE                        No
            (10, 15, "出荷予定日"),           'SHIP_SCHEDULED_DATE     DATE                        Yes
            (11, 15, "客先品目No"),           'CUSTOMER_ITEM_NO        VARCHAR2(20 BYTE)           No
            (12, 15, "品目No"),               'ITEM_NO                 VARCHAR2(20 BYTE)           Yes
            (13, 15, "需要数"),               'DEMAND_QTY              NUMBER(10,0)                No
            (14, 15, "需要単位"),             'DEMAND_UNIT             VARCHAR2(4 BYTE)            Yes
            (15, 15, "通貨コード"),           'CURRENCY_CODE           VARCHAR2(3 BYTE)            Yes
            (16, 15, "出荷在庫場所"),         'SHIP_STOCK_LOCATION     VARCHAR2(25 BYTE)           Yes
            (17, 15, "会社ID"),               'COMPANY_ID              VARCHAR2(25 BYTE)           Yes
            (18, 15, "製品コード"),           'PRODUCT_CODE            VARCHAR2(20 BYTE)           Yes
            (19, 15, "請求基準"),             'BILLING_STANDARD        VARCHAR2(3 BYTE)            Yes
            (20, 15, "出荷プロセスタイプ"),   'SHIP_PROCESS_TYPE       CHAR(1 BYTE)                Yes
            (21, 15, "納入指示フラグ"),       'DELIVERY_INSTR_FLAG     CHAR(1 BYTE)                Yes
            (22, 15, "受注番号"),             'ORDER_NO                VARCHAR2(45 BYTE)           Yes
            (23, 15, "コメント"),             'REMARKS                 VARCHAR2(45 BYTE)           Yes
            (24, 15, "納入先コード"),         'DELIVERY_CODE           VARCHAR2(20 BYTE)           Yes
            (43, 15, "出荷日"),               'SHIP_DATE               DATE                        Yes
            (45, 15, "輸送方法"),             'TRANSPORT_METHOD        VARCHAR2(3 BYTE)            Yes
            (46, 15, "出荷計画日"),           'SHIP_PLAN_DATE          DATE                        Yes
            (47, 15, "客先発注No行番号"),     'CUSTOMER_ORDER_LINE_NO  VARCHAR2(2 BYTE)            Yes
            (48, 15, "日割前受注数"),         'PRE_DAILY_ORDER_QTY     NUMBER(18,6)                Yes
            (49, 15, "日割前納期"),           'PRE_DAILY_DELIVERY_DATE DATE                        Yes
            (50, 15, "取込ファイルID"),       'IMP_FILE_ID             NUMBER(10,0)                No
            (51, 15, "受注区分"),             'ORDER_TYPE              NUMBER(1,0)                 No
            (52, 15, "分割区分"),             'PRORATED_TYPE           NUMBER(1,0)                 No
            (53, 15, "取引先情報区分"),       'CUSTOMER_INFO_TYPE      VARCHAR2(50 BYTE)           Yes
            (54, 15, "情報区分"),             'INFO_TYPE               CHAR(1 BYTE)                Yes
            (55, 15, "自社予測フラグ"),       'SELF_FCST_FLAG          CHAR(1 BYTE)                Yes
            (56, 15, "自社予測削除フラグ"),   'SELF_FCST_DELETE_FLAG   CHAR(1 BYTE)                Yes
            (57, 15, "消込条件区分"),         'RECONCILE_TYPE          NUMBER(1,0)                 Yes
            (58, 15, "取込実行ID"),           'IMP_RUN_ID              VARCHAR2(36 BYTE)           No
            (59, 15, "ステータス"),           'STATUS                  VARCHAR2(20 BYTE)           Yes
            (60, 15, "有効フラグ"),           'ACTIVE_FLAG             CHAR(1 BYTE)                No
            (61, 15, "登録日時"),             'CREATED_AT              DATE                        No
            (62, 15, "登録ユーザーID"),       'CREATED_USER_ID         VARCHAR2(9 BYTE)            No
            (63, 15, "登録プログラムID"),     'CREATED_PG_ID           VARCHAR2(150 BYTE)          No
            (64, 15, "更新日時"),             'UPDATED_AT              DATE                        No
            (65, 15, "更新ユーザーID"),       'UPDATED_USER_ID         VARCHAR2(9 BYTE)            No
            (66, 15, "更新プログラムID")      'UPDATED_PG_ID           VARCHAR2(150 BYTE)          No
        }
        ''' <summary>
        ''' (出荷状況エラーリスト) 
        ''' Index(Column), Cell width, Title
        ''' </summary>
#If False Then
        Private ErrorListExcelTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
                    (1, 15, "取引先コード"),
                    (2, 15, "請求先"),
                    (3, 15, "客先発注No"),
                    (4, 15, "需要ステイタス"),
                    (5, 15, "出荷先"),
                    (6, 15, "受注日"),
                    (7, 15, "希望納期"),
                    (8, 15, "出荷予定日"),
                    (9, 15, "客先品目No"),
                    (10, 15, "品目No"),
                    (11, 15, "需要数"),
                    (12, 15, "需要単位"),
                    (13, 15, "通貨コード"),
                    (14, 15, "出荷在庫場所"),
                    (15, 15, "会社ID"),
                    (16, 15, "製品コード"),
                    (17, 15, "請求基準"),
                    (18, 15, "出荷プロセスタイプ"),
                    (19, 15, "納入指示フラグ"),
                    (20, 15, "受注番号"),
                    (21, 15, "コメント"),
                    (22, 15, "納入先コード"),
                    (23, 15, "受注時刻"),
                    (24, 15, "売上単価"),
                    (25, 15, "納入時間"),
                    (26, 15, "使用先"),
                    (27, 15, "累計出荷数"),
                    (28, 15, "生産区分"),
                    (29, 15, "文字 2ｹﾀ"),
                    (30, 15, "容器番号"),
                    (31, 15, "文字 3ｹﾀ"),
                    (32, 15, "文字 4ｹﾀ"),
                    (33, 15, "文字 4ｹﾀ"),
                    (34, 15, "文字 5ｹﾀ"),
                    (35, 15, "文字 5ｹﾀ"),
                    (36, 15, "文字 6ｹﾀ"),
                    (37, 15, "発注理由"),
                    (38, 15, "容器収容数"),
                    (39, 15, "得意先 ﾛｯﾄNO"),
                    (40, 15, "初品区分"),
                    (41, 15, "出荷日"),
                    (42, 15, "文字 50ｹﾀ"),
                    (43, 15, "輸送方法"),
                    (44, 15, "出荷計画日"),
                    (45, 15, "客先発注No行番号"),
                    (46, 15, "日割前受注数"),
                    (47, 15, "日割前納期")
        }
#End If
        Private ErrorListExcelTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
                    (1, 15, "ステージID"),
                    (2, 15, "生産計画ID"),
                    (3, 15, "客先ID"),
                    (4, 15, "取引先コード"),
                    (5, 15, "取引先名"),
                    (6, 15, "PC"),
                    (7, 15, "取引先ユニットID"),
                    (8, 15, "取引先ユニット"),
                    (9, 15, "請求先"),
                    (10, 15, "客先発注No"),
                    (11, 15, "需要ステイタス"),
                    (12, 15, "出荷先"),
                    (13, 15, "受注日"),
                    (14, 15, "希望納期"),
                    (15, 15, "出荷予定日"),
                    (16, 15, "客先品目No"),
                    (17, 15, "品目No"),
                    (18, 15, "需要数"),
                    (19, 15, "需要単位"),
                    (20, 15, "通貨コード"),
                    (21, 15, "出荷在庫場所"),
                    (22, 15, "会社ID"),
                    (23, 15, "製品コード"),
                    (24, 15, "請求基準"),
                    (25, 15, "出荷プロセスタイプ"),
                    (26, 15, "納入指示フラグ"),
                    (27, 15, "受注番号"),
                    (28, 15, "コメント"),
                    (29, 15, "納入先コード"),
                    (30, 15, "累計出荷数"),
                    (31, 15, "出荷日"),
                    (32, 15, "輸送方法"),
                    (33, 15, "出荷計画日"),
                    (34, 15, "客先発注No行番号"),
                    (35, 15, "日割前受注数"),
                    (36, 15, "日割前納期"),
                    (37, 15, "ImpFileStageID"),
                    (38, 15, "ImpFileID"),
                    (39, 15, "OrderType"),
                    (40, 15, "ProratedType"),
                    (41, 15, "CustomerInfoType"),
                    (42, 15, "InfoType"),
                    (43, 15, "SelfFcstFlag"),
                    (44, 15, "SelfFcstDeleteFlag"),
                    (45, 15, "ReconcilType"),
                    (46, 15, "ImpRunID"),
                    (47, 15, "Status"),
                    (48, 15, "ActiveFlag"),
                    (49, 15, "CreatedAt"),
                    (50, 15, "CreatedUserID"),
                    (51, 15, "CreatedPGID"),
                    (52, 15, "UpdatedAt"),
                    (53, 15, "UpdatedUserID"),
                    (54, 15, "UpdatedPGID"),
                    (55, 15, "ProdMgmtUserID")
                    }
        ''' <summary>
        ''' 出荷状況エラーリスト出力 の excel 出力
        ''' </summary>
        ''' <param name="updateDate"></param>
        ''' <param name="rows"></param>
        ''' <returns></returns>
        Public Shared Function ShippingStatusErrorExcelOut(strPath As String, updateDate As Date, rows As List(Of OrderStageViewRow)) As String

            Dim excel = New OrderProductionPlanExcelFile()
            Return excel.OrdersExcelFile(GetErrorListExcelFilename(strPath, updateDate), rows)

        End Function


        ''' <summary>
        ''' Error list filename取得
        ''' </summary>
        ''' <param name="fileDate"></param>
        ''' <returns></returns>
        Public Shared Function GetErrorListExcelFilename(strPath As String, fileDate As DateTime) As String

            ' Server 側の File 保存Folder
            'Dim strPath = GetWorkPath()
            Dim filename = IO.Path.Combine(strPath, $"A-R-COEDI-F_{fileDate.ToString("yyyyMMssHHmmss")}.xlsx")
            Return filename

        End Function


        ''' <summary>
        ''' 生産計画 Excel ファイル出力
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="ordersRow"></param>
        ''' <returns></returns>
        Public Function OrdersExcelFile(filename As String, ordersRow As List(Of OrdersRow)) As String
            Dim rt = ""

            Try
                'ワークブックを作成
                Using objWBook As New XLWorkbook
                    ' (差異リスト:内示) 
                    Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add("生産計画")

                    For Each itemp In OrdersTitle
                        objSheet1.Column(itemp.index).Width = itemp.width
                        objSheet1.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    Dim offset = 1  ' Title 行 のかさまし分
                    For Each row In ordersRow.Select(Function(val, idx) (item:=val, index:=idx + 1 + offset))
                        objSheet1.Cell(row.index, 1).Value = row.item.OrderId
                        objSheet1.Cell(row.index, 2).Value = row.item.CustomerSettingId
                        objSheet1.Cell(row.index, 3).Value = row.item.CustomerCode
                        objSheet1.Cell(row.index, 4).Value = row.item.BillingTo
                        objSheet1.Cell(row.index, 5).Value = row.item.CustomerOrderNo
                        objSheet1.Cell(row.index, 6).Value = row.item.DemandStatus
                        objSheet1.Cell(row.index, 7).Value = row.item.ShipTo
                        objSheet1.Cell(row.index, 8).Value = row.item.OrderDate
                        objSheet1.Cell(row.index, 9).Value = row.item.DueDate
                        objSheet1.Cell(row.index, 10).Value = row.item.ShipScheduledDate
                        objSheet1.Cell(row.index, 11).Value = row.item.CustomerItemNo
                        objSheet1.Cell(row.index, 12).Value = row.item.ItemNo
                        objSheet1.Cell(row.index, 13).Value = row.item.DemandQty
                        objSheet1.Cell(row.index, 14).Value = row.item.DemandUnit
                        objSheet1.Cell(row.index, 15).Value = row.item.CurrencyCode
                        objSheet1.Cell(row.index, 16).Value = row.item.ShipStockLocation
                        objSheet1.Cell(row.index, 17).Value = row.item.CompanyId
                        objSheet1.Cell(row.index, 18).Value = row.item.ProductCode
                        objSheet1.Cell(row.index, 19).Value = row.item.BillingStandard
                        objSheet1.Cell(row.index, 20).Value = row.item.ShipProcessType
                        objSheet1.Cell(row.index, 21).Value = row.item.DeliveryInstrFlag
                        objSheet1.Cell(row.index, 22).Value = row.item.OrderNo
                        objSheet1.Cell(row.index, 23).Value = row.item.Remarks
                        objSheet1.Cell(row.index, 24).Value = row.item.DeliveryCode
                        objSheet1.Cell(row.index, 25).Value = row.item.OrderTime
                        objSheet1.Cell(row.index, 26).Value = row.item.SalesUnitPrice
                        objSheet1.Cell(row.index, 27).Value = row.item.DeliveryTime
                        objSheet1.Cell(row.index, 28).Value = row.item.UsageLocation
                        objSheet1.Cell(row.index, 29).Value = row.item.TotalShipQty
                        objSheet1.Cell(row.index, 30).Value = row.item.ProductionCategory
                        objSheet1.Cell(row.index, 31).Value = row.item.Char2
                        objSheet1.Cell(row.index, 32).Value = row.item.ContainerNo
                        objSheet1.Cell(row.index, 33).Value = row.item.Char3
                        objSheet1.Cell(row.index, 34).Value = row.item.Char4
                        objSheet1.Cell(row.index, 35).Value = row.item.Char4_2
                        objSheet1.Cell(row.index, 36).Value = row.item.Char5
                        objSheet1.Cell(row.index, 37).Value = row.item.Char5_2
                        objSheet1.Cell(row.index, 38).Value = row.item.Char6
                        objSheet1.Cell(row.index, 39).Value = row.item.OrderReason
                        objSheet1.Cell(row.index, 40).Value = row.item.ContainerCapacity
                        objSheet1.Cell(row.index, 41).Value = row.item.CustomerLotNo
                        objSheet1.Cell(row.index, 42).Value = row.item.InitialFlag
                        objSheet1.Cell(row.index, 43).Value = row.item.ShipDate
                        objSheet1.Cell(row.index, 44).Value = row.item.Char50
                        objSheet1.Cell(row.index, 45).Value = row.item.TransportMethod
                        objSheet1.Cell(row.index, 46).Value = row.item.ShipPlanDate
                        objSheet1.Cell(row.index, 47).Value = row.item.CustomerOrderLineNo
                        objSheet1.Cell(row.index, 48).Value = row.item.PreDailyOrderQty
                        objSheet1.Cell(row.index, 49).Value = row.item.PreDailyDeliveryDate
                        objSheet1.Cell(row.index, 50).Value = row.item.ImpFileId
                        objSheet1.Cell(row.index, 51).Value = row.item.OrderType
                        objSheet1.Cell(row.index, 52).Value = row.item.ProratedType
                        objSheet1.Cell(row.index, 53).Value = row.item.CustomerInfoType
                        objSheet1.Cell(row.index, 54).Value = row.item.InfoType
                        objSheet1.Cell(row.index, 55).Value = row.item.SelfFcstFlag
                        objSheet1.Cell(row.index, 56).Value = row.item.SelfFcstDeleteFlag
                        objSheet1.Cell(row.index, 57).Value = row.item.ReconcileType
                        objSheet1.Cell(row.index, 58).Value = row.item.ImpRunId
                        objSheet1.Cell(row.index, 59).Value = row.item.Status
                        objSheet1.Cell(row.index, 60).Value = row.item.ActiveFlag
                        objSheet1.Cell(row.index, 61).Value = row.item.CreatedAt
                        objSheet1.Cell(row.index, 62).Value = row.item.CreatedUserId
                        objSheet1.Cell(row.index, 63).Value = row.item.CreatedPgId
                        objSheet1.Cell(row.index, 64).Value = row.item.UpdatedAt
                        objSheet1.Cell(row.index, 65).Value = row.item.UpdatedUserId
                        objSheet1.Cell(row.index, 66).Value = row.item.UpdatedPgId
                    Next
                    'ファイルに保存 
                    objWBook.SaveAs(filename)
                End Using
                rt = ""
            Catch ex As Exception
                rt = "OrderExcelFile error: 生産計画 Excel faile 作成エラー"
            End Try
            Return rt

        End Function
        ''' <summary>
        ''' 生産計画 Excel ファイル出力
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="ordersRow"></param>
        ''' <returns></returns>
        Public Function OrdersStageExcelFile(filename As String, ordersRow As List(Of OrdersStageRow)) As String
            Dim rt = ""

            Try
                'ワークブックを作成
                Using objWBook As New XLWorkbook
                    ' (差異リスト:内示) 
                    Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add("生産計画")

                    For Each itemp In OrdersStageTitle
                        objSheet1.Column(itemp.index).Width = itemp.width
                        objSheet1.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    Dim offset = 1  ' Title 行 のかさまし分
                    For Each row In ordersRow.Select(Function(val, idx) (item:=val, index:=idx + 1 + offset))
                        objSheet1.Cell(row.index, 1).Value = row.item.OrderId
                        objSheet1.Cell(row.index, 2).Value = row.item.CustomerSettingId
                        objSheet1.Cell(row.index, 3).Value = row.item.CustomerCode
                        objSheet1.Cell(row.index, 4).Value = row.item.BillingTo
                        objSheet1.Cell(row.index, 5).Value = row.item.CustomerOrderNo
                        objSheet1.Cell(row.index, 6).Value = row.item.DemandStatus
                        objSheet1.Cell(row.index, 7).Value = row.item.ShipTo
                        objSheet1.Cell(row.index, 8).Value = row.item.OrderDate
                        objSheet1.Cell(row.index, 9).Value = row.item.DueDate
                        objSheet1.Cell(row.index, 10).Value = row.item.ShipScheduledDate
                        objSheet1.Cell(row.index, 11).Value = row.item.CustomerItemNo
                        objSheet1.Cell(row.index, 12).Value = row.item.ItemNo
                        objSheet1.Cell(row.index, 13).Value = row.item.DemandQty
                        objSheet1.Cell(row.index, 14).Value = row.item.DemandUnit
                        objSheet1.Cell(row.index, 15).Value = row.item.CurrencyCode
                        objSheet1.Cell(row.index, 16).Value = row.item.ShipStockLocation
                        objSheet1.Cell(row.index, 17).Value = row.item.CompanyId
                        objSheet1.Cell(row.index, 18).Value = row.item.ProductCode
                        objSheet1.Cell(row.index, 19).Value = row.item.BillingStandard
                        objSheet1.Cell(row.index, 20).Value = row.item.ShipProcessType
                        objSheet1.Cell(row.index, 21).Value = row.item.DeliveryInstrFlag
                        objSheet1.Cell(row.index, 22).Value = row.item.OrderNo
                        objSheet1.Cell(row.index, 23).Value = row.item.Remarks
                        objSheet1.Cell(row.index, 24).Value = row.item.DeliveryCode
                        'objSheet1.Cell(row.index, 25).Value = row.item.OrderTime
                        'objSheet1.Cell(row.index, 26).Value = row.item.SalesUnitPrice
                        'objSheet1.Cell(row.index, 27).Value = row.item.DeliveryTime
                        'objSheet1.Cell(row.index, 28).Value = row.item.UsageLocation
                        'objSheet1.Cell(row.index, 29).Value = row.item.TotalShipQty
                        'objSheet1.Cell(row.index, 30).Value = row.item.ProductionCategory
                        'objSheet1.Cell(row.index, 31).Value = row.item.Char2
                        'objSheet1.Cell(row.index, 32).Value = row.item.ContainerNo
                        'objSheet1.Cell(row.index, 33).Value = row.item.Char3
                        'objSheet1.Cell(row.index, 34).Value = row.item.Char4
                        'objSheet1.Cell(row.index, 35).Value = row.item.Char4_2
                        'objSheet1.Cell(row.index, 36).Value = row.item.Char5
                        'objSheet1.Cell(row.index, 37).Value = row.item.Char5_2
                        'objSheet1.Cell(row.index, 38).Value = row.item.Char6
                        'objSheet1.Cell(row.index, 39).Value = row.item.OrderReason
                        'objSheet1.Cell(row.index, 40).Value = row.item.ContainerCapacity
                        'objSheet1.Cell(row.index, 41).Value = row.item.CustomerLotNo
                        'objSheet1.Cell(row.index, 42).Value = row.item.InitialFlag
                        objSheet1.Cell(row.index, 43).Value = row.item.ShipDate
                        'objSheet1.Cell(row.index, 44).Value = row.item.Char50
                        objSheet1.Cell(row.index, 45).Value = row.item.TransportMethod
                        objSheet1.Cell(row.index, 46).Value = row.item.ShipPlanDate
                        objSheet1.Cell(row.index, 47).Value = row.item.CustomerOrderLineNo
                        objSheet1.Cell(row.index, 48).Value = row.item.PreDailyOrderQty
                        objSheet1.Cell(row.index, 49).Value = row.item.PreDailyDeliveryDate
                        objSheet1.Cell(row.index, 50).Value = row.item.ImpFileId
                        objSheet1.Cell(row.index, 51).Value = row.item.OrderType
                        objSheet1.Cell(row.index, 52).Value = row.item.ProratedType
                        objSheet1.Cell(row.index, 53).Value = row.item.CustomerInfoType
                        objSheet1.Cell(row.index, 54).Value = row.item.InfoType
                        objSheet1.Cell(row.index, 55).Value = row.item.SelfFcstFlag
                        objSheet1.Cell(row.index, 56).Value = row.item.SelfFcstDeleteFlag
                        objSheet1.Cell(row.index, 57).Value = row.item.ReconcileType
                        objSheet1.Cell(row.index, 58).Value = row.item.ImpRunId
                        objSheet1.Cell(row.index, 59).Value = row.item.Status
                        objSheet1.Cell(row.index, 60).Value = row.item.ActiveFlag
                        objSheet1.Cell(row.index, 61).Value = row.item.CreatedAt
                        objSheet1.Cell(row.index, 62).Value = row.item.CreatedUserId
                        objSheet1.Cell(row.index, 63).Value = row.item.CreatedPgId
                        objSheet1.Cell(row.index, 64).Value = row.item.UpdatedAt
                        objSheet1.Cell(row.index, 65).Value = row.item.UpdatedUserId
                        objSheet1.Cell(row.index, 66).Value = row.item.UpdatedPgId
                    Next
                    'ファイルに保存 
                    objWBook.SaveAs(filename)
                End Using
                rt = ""
            Catch ex As Exception
                rt = "OrderExcelFile error: 生産計画 Excel faile 作成エラー"
            End Try
            Return rt

        End Function
        ''' <summary>
        ''' 出荷状況エラー Excel ファイル出力
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="ordersRow"></param>
        ''' <returns></returns>
        Public Function OrdersExcelFile(filename As String, ordersRow As List(Of OrderStageViewRow)) As String
            Dim rt = ""

            Try
                'ワークブックを作成
                Using objWBook As New XLWorkbook
                    Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add("出荷状況エラー")

                    For Each itemp In ErrorListExcelTitle
                        objSheet1.Column(itemp.index).Width = itemp.width
                        objSheet1.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    Dim offset = 1  ' Title 行 のかさまし分
                    For Each row In ordersRow.Select(Function(val, idx) (item:=val, index:=idx + 1 + offset))
                        objSheet1.Cell(row.index, 1).Value = row.item.StageID
                        objSheet1.Cell(row.index, 2).Value = row.item.ProdPlanID
                        objSheet1.Cell(row.index, 3).Value = row.item.CustomerSettingID
                        objSheet1.Cell(row.index, 4).Value = row.item.CustomerCode
                        objSheet1.Cell(row.index, 5).Value = row.item.CustomerName
                        objSheet1.Cell(row.index, 6).Value = row.item.ProfitCenter
                        objSheet1.Cell(row.index, 7).Value = row.item.CustomerUnitID
                        objSheet1.Cell(row.index, 8).Value = row.item.CustomerUnitName
                        objSheet1.Cell(row.index, 9).Value = row.item.BillingTo
                        objSheet1.Cell(row.index, 10).Value = row.item.CustomerOrderNo
                        objSheet1.Cell(row.index, 11).Value = row.item.DemandStatus
                        objSheet1.Cell(row.index, 12).Value = row.item.ShipTo
                        objSheet1.Cell(row.index, 13).Value = row.item.OrderDate
                        objSheet1.Cell(row.index, 14).Value = row.item.DueDate
                        objSheet1.Cell(row.index, 15).Value = row.item.ShipScheduledDate
                        objSheet1.Cell(row.index, 16).Value = row.item.CustomerItemNo
                        objSheet1.Cell(row.index, 17).Value = row.item.ItemNo
                        objSheet1.Cell(row.index, 18).Value = row.item.DemandQty
                        objSheet1.Cell(row.index, 19).Value = row.item.DemandUnit
                        objSheet1.Cell(row.index, 20).Value = row.item.CurrencyCode
                        objSheet1.Cell(row.index, 21).Value = row.item.ShipStockLocation
                        objSheet1.Cell(row.index, 22).Value = row.item.CompanyId
                        objSheet1.Cell(row.index, 23).Value = row.item.ProductCode
                        objSheet1.Cell(row.index, 24).Value = row.item.BillingStandard
                        objSheet1.Cell(row.index, 25).Value = row.item.ShipProcessType
                        objSheet1.Cell(row.index, 26).Value = row.item.DeliveryInstrFlag
                        objSheet1.Cell(row.index, 27).Value = row.item.OrderNo
                        objSheet1.Cell(row.index, 28).Value = row.item.Remarks
                        objSheet1.Cell(row.index, 29).Value = row.item.DeliveryCode
                        objSheet1.Cell(row.index, 30).Value = row.item.TotalShipQty
                        objSheet1.Cell(row.index, 31).Value = row.item.ShipDate
                        objSheet1.Cell(row.index, 32).Value = row.item.TransportMethod
                        objSheet1.Cell(row.index, 33).Value = row.item.ShipPlanDate
                        objSheet1.Cell(row.index, 34).Value = row.item.CustomerOrderLineNo
                        objSheet1.Cell(row.index, 35).Value = row.item.PreDailyOrderQty
                        objSheet1.Cell(row.index, 36).Value = row.item.PreDailyDeliveryDate
                        objSheet1.Cell(row.index, 37).Value = row.item.ImpFileStageID
                        objSheet1.Cell(row.index, 38).Value = row.item.ImpFileID
                        objSheet1.Cell(row.index, 39).Value = row.item.OrderType
                        objSheet1.Cell(row.index, 40).Value = row.item.ProratedType
                        objSheet1.Cell(row.index, 41).Value = row.item.CustomerInfoType
                        objSheet1.Cell(row.index, 42).Value = row.item.InfoType
                        objSheet1.Cell(row.index, 43).Value = row.item.SelfFcstFlag
                        objSheet1.Cell(row.index, 44).Value = row.item.SelfFcstDeleteFlag
                        objSheet1.Cell(row.index, 45).Value = row.item.ReconcilType
                        objSheet1.Cell(row.index, 46).Value = row.item.ImpRunID
                        objSheet1.Cell(row.index, 47).Value = row.item.Status
                        objSheet1.Cell(row.index, 48).Value = row.item.ActiveFlag
                        objSheet1.Cell(row.index, 49).Value = row.item.CreatedAt
                        objSheet1.Cell(row.index, 50).Value = row.item.CreatedUserID
                        objSheet1.Cell(row.index, 51).Value = row.item.CreatedPGID
                        objSheet1.Cell(row.index, 52).Value = row.item.UpdatedAt
                        objSheet1.Cell(row.index, 53).Value = row.item.UpdatedUserID
                        objSheet1.Cell(row.index, 54).Value = row.item.UpdatedPGID
                        objSheet1.Cell(row.index, 55).Value = row.item.ProdMgmtUserID
#If False Then
                        objSheet1.Cell(row.index, 1).Value = row.item.CustomerCode
                        objSheet1.Cell(row.index, 2).Value = row.item.BillingTo
                        objSheet1.Cell(row.index, 3).Value = row.item.CustomerOrderNo
                        objSheet1.Cell(row.index, 4).Value = row.item.DemandStatus
                        objSheet1.Cell(row.index, 5).Value = row.item.ShipTo
                        objSheet1.Cell(row.index, 6).Value = row.item.OrderDate
                        objSheet1.Cell(row.index, 7).Value = row.item.DueDate
                        objSheet1.Cell(row.index, 8).Value = row.item.ShipScheduledDate
                        objSheet1.Cell(row.index, 9).Value = row.item.CustomerItemNo
                        objSheet1.Cell(row.index, 10).Value = row.item.ItemNo
                        objSheet1.Cell(row.index, 11).Value = row.item.DemandQty
                        objSheet1.Cell(row.index, 12).Value = row.item.DemandUnit
                        objSheet1.Cell(row.index, 13).Value = row.item.CurrencyCode
                        objSheet1.Cell(row.index, 14).Value = row.item.ShipStockLocation
                        objSheet1.Cell(row.index, 15).Value = row.item.CompanyId
                        objSheet1.Cell(row.index, 16).Value = row.item.ProductCode
                        objSheet1.Cell(row.index, 17).Value = row.item.BillingStandard
                        objSheet1.Cell(row.index, 18).Value = row.item.ShipProcessType
                        objSheet1.Cell(row.index, 19).Value = row.item.DeliveryInstrFlag
                        objSheet1.Cell(row.index, 20).Value = row.item.OrderNo
                        objSheet1.Cell(row.index, 21).Value = row.item.Remarks
                        objSheet1.Cell(row.index, 22).Value = row.item.DeliveryCode
                        'objSheet1.Cell(row.index, 23).Value = row.item.OrderTime
                        objSheet1.Cell(row.index, 23).Value = row.item.SalesUnitPrice
                        objSheet1.Cell(row.index, 24).Value = row.item.DeliveryTime
                        objSheet1.Cell(row.index, 25).Value = row.item.UsageLocation
                        objSheet1.Cell(row.index, 26).Value = row.item.TotalShipQty
                        objSheet1.Cell(row.index, 27).Value = row.item.ProductionCategory
                        objSheet1.Cell(row.index, 28).Value = row.item.Char2
                        objSheet1.Cell(row.index, 29).Value = row.item.ContainerNo
                        objSheet1.Cell(row.index, 30).Value = row.item.Char3
                        objSheet1.Cell(row.index, 31).Value = row.item.Char4
                        objSheet1.Cell(row.index, 32).Value = row.item.Char42
                        objSheet1.Cell(row.index, 33).Value = row.item.Char5
                        objSheet1.Cell(row.index, 34).Value = row.item.Char52
                        objSheet1.Cell(row.index, 35).Value = row.item.Char6
                        objSheet1.Cell(row.index, 36).Value = row.item.OrderReason
                        objSheet1.Cell(row.index, 37).Value = row.item.ContainerCapacity
                        objSheet1.Cell(row.index, 38).Value = row.item.CustomerLotNo
                        objSheet1.Cell(row.index, 39).Value = row.item.InitialFlag
                        objSheet1.Cell(row.index, 40).Value = row.item.ShipDate
                        objSheet1.Cell(row.index, 41).Value = row.item.Char50
                        objSheet1.Cell(row.index, 42).Value = row.item.TransportMethod
                        objSheet1.Cell(row.index, 43).Value = row.item.ShipPlanDate
                        objSheet1.Cell(row.index, 44).Value = row.item.CustomerOrderLineNo
                        objSheet1.Cell(row.index, 45).Value = row.item.PreDailyOrderQty
                        objSheet1.Cell(row.index, 46).Value = row.item.PreDailyDeliveryDate
#End If
                    Next
                    'ファイルに保存 
                    objWBook.SaveAs(filename)
                End Using
                rt = ""
            Catch ex As Exception
                rt = "OrderExcelFile error: 出荷状況エラー Excel faile 作成エラー"
            End Try
            Return rt

        End Function

        ''' <summary>
        ''' 文字列変換
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Private Function ToValueString(value As String) As String

            Dim rt As String
            Try
                rt = value.Trim()
            Catch ex As Exception
                rt = Nothing
            End Try
            Return rt

        End Function
        ''' <summary>
        ''' 文字列変換
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Private Function ToDate(value As String) As Date

            Dim rt As DateTime
            Try
                rt = DateTime.Parse(value)
            Catch ex As Exception
                rt = DateTime.MinValue
            End Try
            Return rt

        End Function
        ''' <summary>
        ''' 文字列変換
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Private Function ToNumber(value As String) As Long
            Dim rt As Long = 0
            Try
                rt = Long.Parse(value)
            Catch ex As Exception
                rt = -1
            End Try
            Return rt
        End Function
        ''' <summary>
        ''' Excel ファイル読み込み
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <returns></returns>
        Public Function OrderExcelFile(filename As String) As List(Of OrdersRow)
            Dim ordersRow = New List(Of OrdersRow)()
            Dim errors = New List(Of String)()
            Try
                'Using stream = New System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                'ワークブックを作成
                Using objWBook As New XLWorkbook(filename)
                    ' (差異リスト:内示) 
                    'Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add("生産計画")

                    Dim objSheet1 = objWBook.Worksheets.FirstOrDefault()

                    If objSheet1 Is Nothing Then
                        Return ordersRow
                    End If
                    ' 最終行
                    Dim rowLast = objSheet1.LastRowUsed().RowNumber()
                    Dim offset = 1  ' Title 行 のかさまし分
                    For i = 1 + offset To rowLast
                        Dim orders = New OrdersRow()
                        orders.OrderId = ToNumber(objSheet1.Cell(i, 1).Value.ToString())
                        orders.CustomerSettingId = ToNumber(objSheet1.Cell(i, 2).Value.ToString())
                        orders.CustomerCode = ToValueString(objSheet1.Cell(i, 3).Value.ToString())
                        orders.BillingTo = ToValueString(objSheet1.Cell(i, 4).Value.ToString())
                        orders.CustomerOrderNo = ToValueString(objSheet1.Cell(i, 5).Value.ToString())
                        orders.DemandStatus = ToValueString(objSheet1.Cell(i, 6).Value.ToString())
                        orders.ShipTo = ToValueString(objSheet1.Cell(i, 7).Value.ToString())
                        orders.OrderDate = ToDate(objSheet1.Cell(i, 8).Value.ToString())
                        orders.DueDate = ToDate(objSheet1.Cell(i, 9).Value.ToString())
                        orders.ShipScheduledDate = ToDate(objSheet1.Cell(i, 10).Value.ToString())
                        orders.CustomerItemNo = ToValueString(objSheet1.Cell(i, 11).Value.ToString())
                        orders.ItemNo = ToValueString(objSheet1.Cell(i, 12).Value.ToString())
                        orders.DemandQty = ToNumber(objSheet1.Cell(i, 13).Value.ToString())
                        orders.DemandUnit = ToValueString(objSheet1.Cell(i, 14).Value.ToString())
                        orders.CurrencyCode = ToValueString(objSheet1.Cell(i, 15).Value.ToString())
                        orders.ShipStockLocation = ToValueString(objSheet1.Cell(i, 16).Value.ToString())
                        orders.CompanyId = ToValueString(objSheet1.Cell(i, 17).Value.ToString())
                        orders.ProductCode = ToValueString(objSheet1.Cell(i, 18).Value.ToString())
                        orders.BillingStandard = ToValueString(objSheet1.Cell(i, 19).Value.ToString())
                        orders.ShipProcessType = ToValueString(objSheet1.Cell(i, 20).Value.ToString())
                        orders.DeliveryInstrFlag = ToValueString(objSheet1.Cell(i, 21).Value.ToString())
                        orders.OrderNo = ToValueString(objSheet1.Cell(i, 22).Value.ToString())
                        orders.Remarks = ToValueString(objSheet1.Cell(i, 23).Value.ToString())
                        orders.DeliveryCode = ToValueString(objSheet1.Cell(i, 24).Value.ToString())
                        orders.OrderTime = ToNumber(objSheet1.Cell(i, 25).Value.ToString())
                        orders.SalesUnitPrice = ToNumber(objSheet1.Cell(i, 26).Value.ToString())
                        orders.DeliveryTime = ToNumber(objSheet1.Cell(i, 27).Value.ToString())
                        orders.UsageLocation = ToValueString(objSheet1.Cell(i, 28).Value.ToString())
                        orders.TotalShipQty = ToNumber(objSheet1.Cell(i, 29).Value.ToString())
                        orders.ProductionCategory = ToValueString(objSheet1.Cell(i, 30).Value.ToString())
                        orders.Char2 = ToValueString(objSheet1.Cell(i, 31).Value.ToString())
                        orders.ContainerNo = ToValueString(objSheet1.Cell(i, 32).Value.ToString())
                        orders.Char3 = ToValueString(objSheet1.Cell(i, 33).Value.ToString())
                        orders.Char4 = ToValueString(objSheet1.Cell(i, 34).Value.ToString())
                        orders.Char4_2 = ToValueString(objSheet1.Cell(i, 35).Value.ToString())
                        orders.Char5 = ToValueString(objSheet1.Cell(i, 36).Value.ToString())
                        orders.Char5_2 = ToValueString(objSheet1.Cell(i, 37).Value.ToString())
                        orders.Char6 = ToValueString(objSheet1.Cell(i, 38).Value.ToString())
                        orders.OrderReason = ToValueString(objSheet1.Cell(i, 39).Value.ToString())
                        orders.ContainerCapacity = ToNumber(objSheet1.Cell(i, 40).Value.ToString())
                        orders.CustomerLotNo = ToValueString(objSheet1.Cell(i, 41).Value.ToString())
                        orders.InitialFlag = ToValueString(objSheet1.Cell(i, 42).Value.ToString())
                        orders.ShipDate = ToDate(objSheet1.Cell(i, 43).Value.ToString())
                        orders.Char50 = ToValueString(objSheet1.Cell(i, 44).Value.ToString())
                        orders.TransportMethod = ToValueString(objSheet1.Cell(i, 45).Value.ToString())
                        orders.ShipPlanDate = ToDate(objSheet1.Cell(i, 46).Value.ToString())
                        orders.CustomerOrderLineNo = ToValueString(objSheet1.Cell(i, 47).Value.ToString())
                        orders.PreDailyOrderQty = ToNumber(objSheet1.Cell(i, 48).Value.ToString())
                        orders.PreDailyDeliveryDate = ToDate(objSheet1.Cell(i, 49).Value.ToString())
                        orders.ImpFileId = ToNumber(objSheet1.Cell(i, 50).Value.ToString())
                        orders.OrderType = ToNumber(objSheet1.Cell(i, 51).Value.ToString())
                        orders.ProratedType = ToNumber(objSheet1.Cell(i, 52).Value.ToString())
                        orders.CustomerInfoType = ToValueString(objSheet1.Cell(i, 53).Value.ToString())
                        orders.InfoType = ToValueString(objSheet1.Cell(i, 54).Value.ToString())
                        orders.SelfFcstFlag = ToValueString(objSheet1.Cell(i, 55).Value.ToString())
                        orders.SelfFcstDeleteFlag = ToValueString(objSheet1.Cell(i, 56).Value.ToString())
                        orders.ReconcileType = ToNumber(objSheet1.Cell(i, 57).Value.ToString())
                        orders.ImpRunId = ToValueString(objSheet1.Cell(i, 58).Value.ToString())
                        orders.Status = ToValueString(objSheet1.Cell(i, 59).Value.ToString())
                        orders.ActiveFlag = ToValueString(objSheet1.Cell(i, 60).Value.ToString())
                        orders.CreatedAt = ToDate(objSheet1.Cell(i, 61).Value.ToString())
                        orders.CreatedUserId = ToValueString(objSheet1.Cell(i, 62).Value.ToString())
                        orders.CreatedPgId = ToValueString(objSheet1.Cell(i, 63).Value.ToString())
                        orders.UpdatedAt = ToDate(objSheet1.Cell(i, 64).Value.ToString())
                        orders.UpdatedUserId = ToValueString(objSheet1.Cell(i, 65).Value.ToString())
                        orders.UpdatedPgId = ToValueString(objSheet1.Cell(i, 66).Value.ToString())
                        'errors.Add(CheckOrdersData(orders))
                        ordersRow.Add(orders)
                    Next
                End Using
                'End Using


            Catch ex As Exception
                ordersRow = Nothing
            End Try
            Return ordersRow

        End Function

        '''' <summary>
        '''' 取り込みデータの 正誤確認
        '''' </summary>
        '''' <param name="orders"></param>
        '''' <returns></returns>
        'Private Function CheckOrdersData(orders As OrdersRow) As String
        '    'CUSTOMER_SETTING_ID（取引先設定ID）
        '    'CUSTOMER_CODE（取引先コード）
        '    'CUSTOMER_ORDER_NO（客先発注No）
        '    '※ORDER_TYPE（受注区分）が[2 Or 3]のレコードのみ
        '    'SHIP_TO（出荷先）
        '    'DUE_DATE（希望納期）
        '    'SHIP_SCHEDULED_DATE（出荷予定日）
        '    'CUSTOMER_ITEM_NO（客先品目No）
        '    'DEMAND_QTY（需要数）
        '    'CUSTOMER_ORDER_LINE_NO(客先発注No行番号)
        '    '※ORDER_TYPE（受注区分）が[2 Or 3]のレコードのみ
        '    'PRE_DAILY_ORDER_QTY（日割前受注数）
        '    'PRE_DAILY_DELIVERY_DATE（日割前納期）
        '    'ORDER_TYPE（受注区分）
        '    'PRORATED_TYPE(分割区分)




        'End Function
        ''' <summary>
        ''' Work path 取得 (ない場合は作成する)
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetWorkPath() As String

            Dim rawWork As String = GetWorkFolderRoot()
            Dim workFolder = rawWork & "FileTempFolder"
            If (Not IO.File.Exists(workFolder)) Then
                EnsureDirectory(workFolder)
            End If
            Return workFolder
        End Function

    End Class
End Namespace
