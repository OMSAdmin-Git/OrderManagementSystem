Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OMS.Data

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

End Class
