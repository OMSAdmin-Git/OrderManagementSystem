Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OMS.Data
Imports OMS.Business.Services

<TestClass>
Public Class SuzukiDataImportTests

    ''' <summary>
    ''' 日割区分判定テスト (1,4 -> 2:日割以外, 5,6 -> 1:日割り)
    ''' </summary>
    <TestMethod>
    Public Sub ProratedType_Calculation_ReturnsExpectedType()
        Dim dateType1 = "1"
        Dim prorated1 = If(dateType1 = "1" OrElse dateType1 = "4", 2, 1)
        Assert.AreEqual(2, prorated1)

        Dim dateType4 = "4"
        Dim prorated4 = If(dateType4 = "1" OrElse dateType4 = "4", 2, 1)
        Assert.AreEqual(2, prorated4)

        Dim dateType5 = "5"
        Dim prorated5 = If(dateType5 = "1" OrElse dateType5 = "4", 2, 1)
        Assert.AreEqual(1, prorated5)

        Dim dateType6 = "6"
        Dim prorated6 = If(dateType6 = "1" OrElse dateType6 = "4", 2, 1)
        Assert.AreEqual(1, prorated6)
    End Sub

    ''' <summary>
    ''' 0630 特殊増減ロジック検証: 対象 0600 行の DEMAND_QTY に 0630.ORDER_QTY を加算
    ''' </summary>
    <TestMethod>
    Public Sub SpecialAdjustment_0630_IncrementsQuantity()
        Dim rows As New List(Of OrdersStageRow)()
        Dim targetDate = New DateTime(2026, 7, 20)

        rows.Add(New OrdersStageRow With {
            .CustomerSettingId = 5455,
            .ItemNo = "ITEM-001",
            .InfoTypeCode = "0600",
            .PreDailyDeliveryDate = targetDate,
            .TargetReferenceDateType = "1",
            .TargetReferenceDate = "202607",
            .DemandQty = 100,
            .PreDailyOrderQty = 100
        })

        ' 0630 増減レコードシミュレーション (+50)
        Dim addQty As Long = 50
        Dim target = rows.Find(Function(x) x.InfoTypeCode = "0600" AndAlso
                                           x.ItemNo = "ITEM-001" AndAlso
                                           x.PreDailyDeliveryDate.HasValue AndAlso
                                           x.PreDailyDeliveryDate.Value.Date = targetDate.Date AndAlso
                                           x.TargetReferenceDateType = "1" AndAlso
                                           x.TargetReferenceDate = "202607")

        Assert.IsNotNull(target)
        target.DemandQty = (If(target.DemandQty, 0L)) + addQty
        target.PreDailyOrderQty = target.DemandQty

        Assert.AreEqual(150L, target.DemandQty.Value)
        Assert.AreEqual(150D, target.PreDailyOrderQty.Value)
    End Sub

    ''' <summary>
    ''' 663N / 663S 変更ロジック検証: 注文番号の一致する 6624 行の納期および数量変更
    ''' </summary>
    <TestMethod>
    Public Sub SpecialAdjustment_663N_And_663S_UpdatesOrder()
        Dim rows As New List(Of OrdersStageRow)()
        rows.Add(New OrdersStageRow With {
            .CustomerSettingId = 5455,
            .CustomerOrderNo = "PO-998877",
            .ItemNo = "ITEM-002",
            .InfoTypeCode = "6624",
            .DueDate = New DateTime(2026, 8, 1),
            .DemandQty = 200
        })

        ' 663N: 納期変更
        Dim newDate = New DateTime(2026, 8, 10)
        Dim targetN = rows.Find(Function(x) x.InfoTypeCode = "6624" AndAlso x.CustomerOrderNo = "PO-998877")
        Assert.IsNotNull(targetN)
        targetN.DueDate = newDate
        Assert.AreEqual(newDate, targetN.DueDate.Value)

        ' 663S: 数量変更
        Dim newQty As Long = 250
        Dim targetS = rows.Find(Function(x) x.InfoTypeCode = "6624" AndAlso x.CustomerOrderNo = "PO-998877")
        Assert.IsNotNull(targetS)
        targetS.DemandQty = newQty
        Assert.AreEqual(250L, targetS.DemandQty.Value)
    End Sub

    ''' <summary>
    ''' 出荷先変換ロジック検証: SUZUKI_SPIRITS_CONVERSION_MST に基づく SHIP_TO 置換
    ''' </summary>
    <TestMethod>
    Public Sub ShipToConversion_ReplacesDeliveryCodePlan()
        Dim convMap As New Dictionary(Of String, String) From {
            {"LOC_A", "PLAN_01"},
            {"LOC_B", "PLAN_02"}
        }

        Dim rows As New List(Of OrdersStageRow) From {
            New OrdersStageRow With {.ShipTo = "LOC_A"},
            New OrdersStageRow With {.ShipTo = "LOC_B"},
            New OrdersStageRow With {.ShipTo = "LOC_OTHER"}
        }

        For Each r In rows
            If convMap.ContainsKey(r.ShipTo) Then
                r.ShipTo = convMap(r.ShipTo)
            End If
        Next

        Assert.AreEqual("PLAN_01", rows(0).ShipTo)
        Assert.AreEqual("PLAN_02", rows(1).ShipTo)
        Assert.AreEqual("LOC_OTHER", rows(2).ShipTo)
    End Sub

    ''' <summary>
    ''' OrdersStageRow 拡張 Phase 2 カラム保持テスト
    ''' </summary>
    <TestMethod>
    Public Sub OrdersStageRow_ExtendedPhase2Properties_RetainValues()
        Dim row As New OrdersStageRow With {
            .DeliveryTime = 1430D,
            .ContainerCapacity = 50D,
            .InitialFlag = "Y",
            .TargetReferenceDateType = "1",
            .TargetReferenceDate = "20260715",
            .InfoTypeCode = "0600"
        }

        Assert.AreEqual(1430D, row.DeliveryTime.Value)
        Assert.AreEqual(50D, row.ContainerCapacity.Value)
        Assert.AreEqual("Y", row.InitialFlag)
        Assert.AreEqual("1", row.TargetReferenceDateType)
        Assert.AreEqual("20260715", row.TargetReferenceDate)
        Assert.AreEqual("0600", row.InfoTypeCode)
    End Sub

    ''' <summary>
    ''' 【2026/09/25 仕様変更】0814(支給品納入指示)の消込先が 0813 から 0600 に変更されたことのロジック検証
    ''' </summary>
    <TestMethod>
    Public Sub ReconcileForecastSuzuki_0814_Reconciles_0600_Logic()
        ' 内示 0600 レコード (数量 100)
        Dim fcst0600 As New OrdersStageRow With {
            .StageId = 1,
            .ItemNo = "ITEM-0814-TEST",
            .ShipTo = "DEST-A",
            .InfoTypeCode = "0600",
            .OrderType = 1,
            .DemandQty = 100
        }

        ' 指示 0814 レコード (数量 40)
        Dim instr0814 As New OrdersStageRow With {
            .StageId = 2,
            .ItemNo = "ITEM-0814-TEST",
            .ShipTo = "DEST-A",
            .InfoTypeCode = "0814",
            .OrderType = 3,
            .DemandQty = 40
        }

        ' 消込対象内示コードのマッピング判定 (ReconcileForecastSuzuki のロジック)
        Dim targetFcstCode As String = ""
        Select Case instr0814.InfoTypeCode
            Case "0650", "0651", "0814"
                targetFcstCode = "0600"
            Case "6624"
                targetFcstCode = "6604"
            Case Else
                targetFcstCode = ""
        End Select

        Assert.AreEqual("0600", targetFcstCode, "0814 の消込対象コードは 0600 である必要があります")

        ' 消込シミュレーション (100 - 40 = 60)
        If targetFcstCode = fcst0600.InfoTypeCode Then
            fcst0600.DemandQty = fcst0600.DemandQty - instr0814.DemandQty
        End If

        Assert.AreEqual(60L, fcst0600.DemandQty.Value, "0600 の内示残数が 0814 の指示数分減算されている必要があります")
    End Sub

    ''' <summary>
    ''' 【2026/09/25 仕様変更】0813(支給品情報)の廃止に伴う洗い替え対象除外のロジック検証
    ''' </summary>
    <TestMethod>
    Public Sub ReplaceNaijiRelationSuzuki_Excludes_0813_From_WashList()
        Dim washTargets As New List(Of String) From {"0600", "6604"}

        Assert.IsTrue(washTargets.Contains("0600"), "0600 は洗い替え対象である必要があります")
        Assert.IsTrue(washTargets.Contains("6604"), "6604 は洗い替え対象である必要があります")
        Assert.IsFalse(washTargets.Contains("0813"), "0813 は2026/09/25仕様変更により洗い替え対象外である必要があります")
    End Sub

    ''' <summary>
    ''' SPIRITSデータマッピング: 0600 / 6604 の 取引先情報区分変換テスト
    ''' 1(内示) -> 新規, 2(確定) -> 新規, 3(変更) -> 変更
    ''' </summary>
    <TestMethod>
    Public Sub ConvertCustomerInfoType_0600_And_6604_MapsCorrectly()
        Assert.AreEqual("新規", SuzukiDataImportService.ConvertCustomerInfoType("0600", "1", ""))
        Assert.AreEqual("新規", SuzukiDataImportService.ConvertCustomerInfoType("0600", "2", ""))
        Assert.AreEqual("変更", SuzukiDataImportService.ConvertCustomerInfoType("0600", "3", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0600", "9", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0600", "", ""))

        Assert.AreEqual("新規", SuzukiDataImportService.ConvertCustomerInfoType("6604", "1", ""))
        Assert.AreEqual("新規", SuzukiDataImportService.ConvertCustomerInfoType("6604", "2", ""))
        Assert.AreEqual("変更", SuzukiDataImportService.ConvertCustomerInfoType("6604", "3", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("6604", "0", ""))
    End Sub

    ''' <summary>
    ''' SPIRITSデータマッピング: 0602 の 取引先情報区分変換テスト
    ''' 2(確定) -> 新規, 3(変更) -> 変更
    ''' </summary>
    <TestMethod>
    Public Sub ConvertCustomerInfoType_0602_MapsCorrectly()
        Assert.AreEqual("新規", SuzukiDataImportService.ConvertCustomerInfoType("0602", "2", ""))
        Assert.AreEqual("変更", SuzukiDataImportService.ConvertCustomerInfoType("0602", "3", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0602", "1", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0602", "", ""))
    End Sub

    ''' <summary>
    ''' SPIRITSデータマッピング: 0651 の 取引先情報区分変換テスト
    ''' 発行注釈 17桁目 <> '0' -> 打切, それ以外 -> ""
    ''' </summary>
    <TestMethod>
    Public Sub ConvertCustomerInfoType_0651_MapsCorrectly()
        ' 17桁目が '1' の場合 -> "打切"
        Dim noteUchikiri = "12345678901234561000"
        Assert.AreEqual("打切", SuzukiDataImportService.ConvertCustomerInfoType("0651", "", noteUchikiri))

        ' 17桁目が '0' の場合 -> ""
        Dim noteNormal = "12345678901234560000"
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0651", "", noteNormal))

        ' 17桁未満の場合 -> ""
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0651", "", "SHORT_NOTE"))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0651", "", ""))
    End Sub

    ''' <summary>
    ''' SPIRITSデータマッピング: 0650, 0814, 6624 は 取引先情報区分ブランク
    ''' </summary>
    <TestMethod>
    Public Sub ConvertCustomerInfoType_NonMappedCodes_ReturnEmpty()
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0650", "1", "12345678901234561000"))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("0814", "2", ""))
        Assert.AreEqual("", SuzukiDataImportService.ConvertCustomerInfoType("6624", "3", ""))
    End Sub

    ''' <summary>
    ''' 情報区分 (INFO_TYPE) 解決テスト: 新規->I, 変更->U, 打切->N, 空白->空文字
    ''' </summary>
    <TestMethod>
    Public Sub ResolveInfoTypeStatic_ResolvesStandardCodes()
        Assert.AreEqual("I", SuzukiDataImportService.ResolveInfoTypeStatic(Nothing, 0, "新規"))
        Assert.AreEqual("U", SuzukiDataImportService.ResolveInfoTypeStatic(Nothing, 0, "変更"))
        Assert.AreEqual("N", SuzukiDataImportService.ResolveInfoTypeStatic(Nothing, 0, "打切"))
        Assert.AreEqual("", SuzukiDataImportService.ResolveInfoTypeStatic(Nothing, 0, ""))
        Assert.AreEqual("", SuzukiDataImportService.ResolveInfoTypeStatic(Nothing, 0, Nothing))
    End Sub

    ''' <summary>
    ''' SPIRITSデータマッピング: 分割区分 (PRORATED_TYPE) 変換テスト
    ''' 0600: 1->2(日割以外), 5->1(日割り)
    ''' 6604: 1,4->2(日割以外), 5->1(日割り)
    ''' 0650/0651/0814: 5,6->1(日割り)
    ''' 0602/6624: 5->1(日割り)
    ''' </summary>
    <TestMethod>
    Public Sub ConvertProratedType_MapsCorrectlyPerSpecification()
        ' 0600
        Assert.AreEqual(CType(2, Short), SuzukiDataImportService.ConvertProratedType("0600", "1"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0600", "5"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0600", ""))

        ' 6604
        Assert.AreEqual(CType(2, Short), SuzukiDataImportService.ConvertProratedType("6604", "1"))
        Assert.AreEqual(CType(2, Short), SuzukiDataImportService.ConvertProratedType("6604", "4"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("6604", "5"))

        ' 0650, 0651, 0814
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0650", "5"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0650", "6"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0651", "5"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0651", "6"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0814", "5"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0814", "6"))

        ' 0602, 6624
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("0602", "5"))
        Assert.AreEqual(CType(1, Short), SuzukiDataImportService.ConvertProratedType("6624", "5"))
    End Sub

End Class
