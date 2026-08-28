Imports System.Globalization
Imports System.IO
Imports System.Text
Imports CsvHelper
Imports CsvHelper.Configuration
Imports OMS.Data.SUZUKI

Namespace Services

    ''' <summary>
    ''' スズキ (SPIRITS) CSV パーサー
    ''' </summary>
    Public Class SuzukiCsvParser

        ''' <summary>
        ''' CSV ファイルの1行目（1列目の INFO_TYPE_CODE）を読み取って種別コードを返します。
        ''' </summary>
        Public Shared Function PeekInfoTypeCode(filePath As String) As String
            If Not File.Exists(filePath) Then Return String.Empty

            Dim encoding As Encoding = Encoding.GetEncoding("shift-jis")
            Using reader As New StreamReader(filePath, encoding)
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim tokens As String() = SplitCsvLine(line)
                    If tokens.Length > 0 Then
                        Dim val As String = tokens(0).Trim(" "c, """"c)
                        If val.Length = 4 AndAlso val.All(Function(c) Char.IsLetterOrDigit(c)) Then
                            Return val
                        End If
                    End If
                End While
            End Using

            Return String.Empty
        End Function

        ''' <summary>
        ''' 0600 / 0630 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function Parse0600And0630(filePath As String) As List(Of Spirits0600And0630Row)
            Dim list As New List(Of Spirits0600And0630Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 35 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0600" AndAlso code <> "0630" Then Continue For

                Dim row As New Spirits0600And0630Row() With {
                    .InfoTypeCode = code,
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = ParseDate(CleanCol(cols, 6)),
                    .PublicationTime = ParseInt(CleanCol(cols, 7)),
                    .TargetReferenceDateType = CleanCol(cols, 8),
                    .TargetReferenceDate = CleanCol(cols, 9),
                    .CustomerItemNo = CleanCol(cols, 10),
                    .CustomerItemNoProcessNo = CleanCol(cols, 11),
                    .CustomerItemName = CleanCol(cols, 12),
                    .CriticalSafetyPartsCode = CleanCol(cols, 13),
                    .PackagingCode = CleanCol(cols, 14),
                    .Capacity = ParseInt(CleanCol(cols, 15)),
                    .DeliveryType = CleanCol(cols, 16),
                    .SupplierCode = CleanCol(cols, 17),
                    .SupplierFactoryCode = CleanCol(cols, 18),
                    .SupplierShippingLocation = CleanCol(cols, 19),
                    .DeliveryCode = CleanCol(cols, 20),
                    .DeliveryFactoryCode = CleanCol(cols, 21),
                    .ArrangeManager = CleanCol(cols, 22),
                    .PurchaseManager = CleanCol(cols, 23),
                    .CompleteFactory = CleanCol(cols, 24),
                    .FirstArticleType = CleanCol(cols, 25),
                    .LeadtimeType = CleanCol(cols, 26),
                    .Leadtime = CleanCol(cols, 27),
                    .JerseyNumber = CleanCol(cols, 28),
                    .DeliveryCycle = CleanCol(cols, 29),
                    .ManagementType = CleanCol(cols, 30),
                    .Reserve = CleanCol(cols, 31),
                    .OrderDataType = CleanCol(cols, 32),
                    .DeliveryDateType = CleanCol(cols, 33),
                    .ProductionMonthType = CleanCol(cols, 34),
                    .DeliveryDate = If(cols.Length > 35, ParseDate(CleanCol(cols, 35)), Nothing),
                    .OrderQty = If(cols.Length > 36, ParseLong(CleanCol(cols, 36)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' 0602 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function Parse0602(filePath As String) As List(Of Spirits0602Row)
            Dim list As New List(Of Spirits0602Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 25 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0602" Then Continue For

                Dim row As New Spirits0602Row() With {
                    .InfoTypeCode = code,
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = ParseDate(CleanCol(cols, 6)),
                    .PublicationTime = ParseInt(CleanCol(cols, 7)),
                    .CustomerItemNo = CleanCol(cols, 8),
                    .CustomerItemNoProcessNo = CleanCol(cols, 9),
                    .CustomerItemName = CleanCol(cols, 10),
                    .CriticalSafetyPartsCode = CleanCol(cols, 11),
                    .PackagingCode = CleanCol(cols, 12),
                    .Capacity = ParseInt(CleanCol(cols, 13)),
                    .DeliveryType = CleanCol(cols, 16),
                    .SupplierCode = CleanCol(cols, 17),
                    .SupplierFactoryCode = CleanCol(cols, 18),
                    .SupplierShippingLocation = CleanCol(cols, 19),
                    .DeliveryCode = CleanCol(cols, 20),
                    .DeliveryFactoryCode = CleanCol(cols, 21),
                    .ArrangeManager = CleanCol(cols, 22),
                    .PurchaseManager = CleanCol(cols, 23),
                    .OrderDataType = If(cols.Length > 26, CleanCol(cols, 26), String.Empty),
                    .DeliveryDateType = If(cols.Length > 27, CleanCol(cols, 27), String.Empty),
                    .ProductionMonthType = If(cols.Length > 28, CleanCol(cols, 28), String.Empty),
                    .DeliveryDate = If(cols.Length > 29, ParseDate(CleanCol(cols, 29)), Nothing),
                    .OrderQty = If(cols.Length > 31, ParseLong(CleanCol(cols, 31)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits0501And0502 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0501And0502(filePath As String) As List(Of Spirits0501And0502Row)
            Dim list As New List(Of Spirits0501And0502Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0501" AndAlso code <> "0502" Then Continue For

                Dim row As New Spirits0501And0502Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .PublicationTime = CleanCol(cols, 7),
                    .TargetReferenceDateType = CleanCol(cols, 8),
                    .TargetReferenceDate = CleanCol(cols, 9),
                    .CustomerItemNo = CleanCol(cols, 10),
                    .CustomerItemNoProcessNo = CleanCol(cols, 11),
                    .CustomerItemName = CleanCol(cols, 12),
                    .SuppliersCode = CleanCol(cols, 13),
                    .Reserve1 = CleanCol(cols, 14),
                    .ArrangeManager = CleanCol(cols, 15),
                    .PurchaseManager = CleanCol(cols, 16),
                    .Reserve2 = CleanCol(cols, 17),
                    .OrderDataType = CleanCol(cols, 18),
                    .DeliveryDateType = CleanCol(cols, 19),
                    .OrderQtyType = CleanCol(cols, 20),
                    .DeliveryDate = If(cols.Length > 21, ParseDate(CleanCol(cols, 21)), Nothing),
                    .OrderQty = If(cols.Length > 22, ParseLong(CleanCol(cols, 22)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits0650 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0650(filePath As String) As List(Of Spirits0650Row)
            Dim list As New List(Of Spirits0650Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0650" Then Continue For

                Dim row As New Spirits0650Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .PublicationTime = If(cols.Length > 7, ParseInt(CleanCol(cols, 7)), Nothing),
                    .CustomerItemNo = CleanCol(cols, 8),
                    .CustomerItemNoProcessNo = CleanCol(cols, 9),
                    .CustomerItemName = CleanCol(cols, 10),
                    .CriticalSafetyPartsCode = CleanCol(cols, 11),
                    .PackagingCode = CleanCol(cols, 12),
                    .Capacity = If(cols.Length > 13, ParseInt(CleanCol(cols, 13)), Nothing),
                    .SupplierCode = CleanCol(cols, 14),
                    .SupplierFactoryCode = CleanCol(cols, 15),
                    .SupplierShippingLocation = CleanCol(cols, 16),
                    .DeliveryCode = CleanCol(cols, 17),
                    .DeliveryFactoryCode = CleanCol(cols, 18),
                    .DeliveryLocation = CleanCol(cols, 19),
                    .ArrangeManager = CleanCol(cols, 20),
                    .PurchaseManager = CleanCol(cols, 21),
                    .JerseyNumber = CleanCol(cols, 22),
                    .ProcessProcessNo = CleanCol(cols, 23),
                    .Reserve = CleanCol(cols, 24),
                    .CustomerOrderNo = CleanCol(cols, 25),
                    .CustomerOrderNoProcessNo1 = CleanCol(cols, 26),
                    .CustomerOrderNoProcessNo2 = CleanCol(cols, 27),
                    .SupplyLine = CleanCol(cols, 28),
                    .SupplyProcess = CleanCol(cols, 29),
                    .FirstArticleType = CleanCol(cols, 30),
                    .OrderNotes = CleanCol(cols, 31),
                    .DeliveryType = CleanCol(cols, 32),
                    .OrderDataType = CleanCol(cols, 33),
                    .DeliveryDateType = CleanCol(cols, 34),
                    .ProductionMonthType = CleanCol(cols, 35),
                    .DeliveryDate = If(cols.Length > 36, ParseDate(CleanCol(cols, 36)), Nothing),
                    .DeliveryTime = CleanCol(cols, 37),
                    .OrderQty = If(cols.Length > 38, ParseInt(CleanCol(cols, 38)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits0651 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0651(filePath As String) As List(Of Spirits0651Row)
            Dim list As New List(Of Spirits0651Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0651" Then Continue For

                Dim row As New Spirits0651Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .PublicationTime = If(cols.Length > 7, ParseInt(CleanCol(cols, 7)), Nothing),
                    .CustomerItemNo = CleanCol(cols, 8),
                    .CustomerItemNoProcessNo = CleanCol(cols, 9),
                    .CustomerItemName = CleanCol(cols, 10),
                    .CriticalSafetyPartsCode = CleanCol(cols, 11),
                    .PackagingCode = CleanCol(cols, 12),
                    .Capacity = If(cols.Length > 13, ParseInt(CleanCol(cols, 13)), Nothing),
                    .SupplierCode = CleanCol(cols, 14),
                    .SupplierFactoryCode = CleanCol(cols, 15),
                    .SupplierShippingLocation = CleanCol(cols, 16),
                    .DeliveryCode = CleanCol(cols, 17),
                    .DeliveryFactoryCode = CleanCol(cols, 18),
                    .DeliveryLocation = CleanCol(cols, 19),
                    .ArrangeManager = CleanCol(cols, 20),
                    .PurchaseManager = CleanCol(cols, 21),
                    .JerseyNumber = CleanCol(cols, 22),
                    .InvertFlag = CleanCol(cols, 23),
                    .Reserve = CleanCol(cols, 24),
                    .CustomerOrderNo = CleanCol(cols, 25),
                    .CustomerOrderNoProcessNo1 = CleanCol(cols, 26),
                    .CustomerOrderNoProcessNo2 = CleanCol(cols, 27),
                    .SupplyLine = CleanCol(cols, 28),
                    .SupplyProcess = CleanCol(cols, 29),
                    .FirstArticleType = CleanCol(cols, 30),
                    .OrderNotes = CleanCol(cols, 31),
                    .DeliveryType = CleanCol(cols, 32),
                    .OrderDataType = CleanCol(cols, 33),
                    .DeliveryDateType = CleanCol(cols, 34),
                    .ProductionMonthType = CleanCol(cols, 35),
                    .DeliveryDate = If(cols.Length > 36, ParseDate(CleanCol(cols, 36)), Nothing),
                    .DeliveryTime = CleanCol(cols, 37),
                    .OrderQty = If(cols.Length > 38, ParseInt(CleanCol(cols, 38)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits0740 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0740(filePath As String) As List(Of Spirits0740Row)
            Dim list As New List(Of Spirits0740Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0740" Then Continue For

                Dim row As New Spirits0740Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .CustomerItemNo = CleanCol(cols, 6),
                    .CustomerItemNoProcessNo = CleanCol(cols, 7),
                    .CustomerItemName = CleanCol(cols, 8),
                    .SupplierCode = CleanCol(cols, 9),
                    .SupplierFactoryCode = CleanCol(cols, 10),
                    .DeliveryCode = CleanCol(cols, 11),
                    .DeliveryFactoryCode = CleanCol(cols, 12),
                    .DeliveryLocation = CleanCol(cols, 13),
                    .ArrangeManager = CleanCol(cols, 14),
                    .PurchaseManager = CleanCol(cols, 15),
                    .Reserve = CleanCol(cols, 16),
                    .CustomerOrderNo = CleanCol(cols, 17),
                    .DeliveryType = CleanCol(cols, 18),
                    .DeliveryDate = If(cols.Length > 19, ParseDate(CleanCol(cols, 19)), Nothing),
                    .DeliveryTime = CleanCol(cols, 20),
                    .ProductionMonthType = CleanCol(cols, 21),
                    .OrderQty = If(cols.Length > 22, ParseLong(CleanCol(cols, 22)), Nothing),
                    .AcceptanceDate = If(cols.Length > 23, ParseDate(CleanCol(cols, 23)), Nothing),
                    .AcceptanceTime = CleanCol(cols, 24),
                    .AcceptanceQty = CleanCol(cols, 25),
                    .DeliveryNo = CleanCol(cols, 26),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

''' <summary>
        ''' Spirits0813 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0813(filePath As String) As List(Of Spirits0813Row)
            Dim list As New List(Of Spirits0813Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0813" Then Continue For

                Dim row As New Spirits0813Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .Reserve1 = CleanCol(cols, 4),
                    .PublicationDate = If(cols.Length > 5, ParseDate(CleanCol(cols, 5)), Nothing),
                    .PublicationTime = If(cols.Length > 6, ParseInt(CleanCol(cols, 6)), Nothing),
                    .TargetReferenceDateType = CleanCol(cols, 7),
                    .TargetReferenceDate = CleanCol(cols, 8),
                    .CustomerItemNo = CleanCol(cols, 9),
                    .CustomerItemNoProcessNo = CleanCol(cols, 10),
                    .CustomerItemName = CleanCol(cols, 11),
                    .CriticalSafetyPartsCode = CleanCol(cols, 12),
                    .PackagingCode = CleanCol(cols, 13),
                    .Capacity = If(cols.Length > 14, ParseInt(CleanCol(cols, 14)), Nothing),
                    .SupplierCode = CleanCol(cols, 15),
                    .SupplierFactoryCode = CleanCol(cols, 16),
                    .SupplierShippingLocation = CleanCol(cols, 17),
                    .SupplierName = CleanCol(cols, 18),
                    .DeliveryCode = CleanCol(cols, 19),
                    .DeliveryFactoryCode = CleanCol(cols, 20),
                    .DeliveryLocation = CleanCol(cols, 21),
                    .ArrangeManager = CleanCol(cols, 22),
                    .PurchaseManager = CleanCol(cols, 23),
                    .FirstArticleType = CleanCol(cols, 24),
                    .LeadtimeType = CleanCol(cols, 25),
                    .Leadtime = CleanCol(cols, 26),
                    .JerseyNumber = CleanCol(cols, 27),
                    .DeliveryCycle = CleanCol(cols, 28),
                    .SupplyManager = CleanCol(cols, 29),
                    .SupplyPrimaryAcceptanceLocation = CleanCol(cols, 30),
                    .SupplyDepartureLocation = CleanCol(cols, 31),
                    .Reserve2 = CleanCol(cols, 32),
                    .CustomerOrderNo = CleanCol(cols, 33),
                    .CustomerOrderNoProcessNo1 = CleanCol(cols, 34),
                    .CustomerOrderNoProcessNo2 = CleanCol(cols, 35),
                    .SupplyLine = CleanCol(cols, 36),
                    .SupplyProcess = CleanCol(cols, 37),
                    .DeliveryType = CleanCol(cols, 38),
                    .OrderDataType = CleanCol(cols, 39),
                    .DeliveryDateType = CleanCol(cols, 40),
                    .OrderQtyType = CleanCol(cols, 41),
                    .DeliveryDate = If(cols.Length > 42, ParseDate(CleanCol(cols, 42)), Nothing),
                    .DeliveryTime = CleanCol(cols, 43),
                    .OrderQty = If(cols.Length > 44, ParseLong(CleanCol(cols, 44)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits0814 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits0814(filePath As String) As List(Of Spirits0814Row)
            Dim list As New List(Of Spirits0814Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "0814" Then Continue For

                Dim row As New Spirits0814Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .Reserve1 = CleanCol(cols, 4),
                    .PublicationDate = If(cols.Length > 5, ParseDate(CleanCol(cols, 5)), Nothing),
                    .PublicationTime = If(cols.Length > 6, ParseInt(CleanCol(cols, 6)), Nothing),
                    .TargetReferenceDateType = CleanCol(cols, 7),
                    .TargetReferenceDate = CleanCol(cols, 8),
                    .CustomerItemNo = CleanCol(cols, 9),
                    .CustomerItemNoProcessNo = CleanCol(cols, 10),
                    .CustomerItemName = CleanCol(cols, 11),
                    .CriticalSafetyPartsCode = CleanCol(cols, 12),
                    .PackagingCode = CleanCol(cols, 13),
                    .Capacity = If(cols.Length > 14, ParseInt(CleanCol(cols, 14)), Nothing),
                    .SupplierCode = CleanCol(cols, 15),
                    .SupplierFactoryCode = CleanCol(cols, 16),
                    .SupplierShippingLocation = CleanCol(cols, 17),
                    .SupplierName = CleanCol(cols, 18),
                    .DeliveryCode = CleanCol(cols, 19),
                    .DeliveryFactoryCode = CleanCol(cols, 20),
                    .DeliveryLocation = CleanCol(cols, 21),
                    .ArrangeManager = CleanCol(cols, 22),
                    .PurchaseManager = CleanCol(cols, 23),
                    .FirstArticleType = CleanCol(cols, 24),
                    .LeadtimeType = CleanCol(cols, 25),
                    .Leadtime = CleanCol(cols, 26),
                    .JerseyNumber = CleanCol(cols, 27),
                    .DeliveryCycle = CleanCol(cols, 28),
                    .SupplyManager = CleanCol(cols, 29),
                    .SupplyPrimaryAcceptanceLocation = CleanCol(cols, 30),
                    .SupplyDepartureLocation = CleanCol(cols, 31),
                    .Reserve2 = CleanCol(cols, 32),
                    .CustomerOrderNo = CleanCol(cols, 33),
                    .CustomerOrderNoProcessNo1 = CleanCol(cols, 34),
                    .CustomerOrderNoProcessNo2 = CleanCol(cols, 35),
                    .SupplyLine = CleanCol(cols, 36),
                    .SupplyProcess = CleanCol(cols, 37),
                    .DeliveryType = CleanCol(cols, 38),
                    .OrderDataType = CleanCol(cols, 39),
                    .DeliveryDateType = CleanCol(cols, 40),
                    .OrderQtyType = CleanCol(cols, 41),
                    .DeliveryDate = If(cols.Length > 42, ParseDate(CleanCol(cols, 42)), Nothing),
                    .DeliveryTime = CleanCol(cols, 43),
                    .OrderQty = If(cols.Length > 44, ParseLong(CleanCol(cols, 44)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits6604And6634 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits6604And6634(filePath As String) As List(Of Spirits6604And6634Row)
            Dim list As New List(Of Spirits6604And6634Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "6604" AndAlso code <> "6634" Then Continue For

                Dim row As New Spirits6604And6634Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .TargetReferenceDateType = CleanCol(cols, 7),
                    .TargetReferenceDate = CleanCol(cols, 8),
                    .CustomerItemNo = CleanCol(cols, 9),
                    .CustomerItemNoProcessNo1 = CleanCol(cols, 10),
                    .CustomerItemNoProcessNo2 = CleanCol(cols, 11),
                    .CustomerItemName = CleanCol(cols, 12),
                    .PackagingCode = CleanCol(cols, 13),
                    .Capacity = If(cols.Length > 14, ParseInt(CleanCol(cols, 14)), Nothing),
                    .SupplierCode = CleanCol(cols, 15),
                    .SupplierFactoryCode = CleanCol(cols, 16),
                    .SupplierShippingLocation = CleanCol(cols, 17),
                    .DeliveryCode = CleanCol(cols, 18),
                    .DeliveryFactoryCode = CleanCol(cols, 19),
                    .DeliveryLocation = CleanCol(cols, 20),
                    .DeliveryType = CleanCol(cols, 21),
                    .ConstructionNo = CleanCol(cols, 22),
                    .ArrangeManager = CleanCol(cols, 23),
                    .Reserve = CleanCol(cols, 24),
                    .OrderDataType = CleanCol(cols, 25),
                    .DeliveryDateType = CleanCol(cols, 26),
                    .DeliveryDate = If(cols.Length > 27, ParseDate(CleanCol(cols, 27)), Nothing),
                    .OrderQty = If(cols.Length > 28, ParseInt(CleanCol(cols, 28)), Nothing),
                    .OldOrderQty = If(cols.Length > 29, ParseInt(CleanCol(cols, 29)), Nothing),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits6624 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits6624(filePath As String) As List(Of Spirits6624Row)
            Dim list As New List(Of Spirits6624Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "6624" Then Continue For

                Dim row As New Spirits6624Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .CustomerItemNo = CleanCol(cols, 7),
                    .CustomerItemNoProcessNo1 = CleanCol(cols, 8),
                    .CustomerItemNoProcessNo2 = CleanCol(cols, 9),
                    .CustomerItemName = CleanCol(cols, 10),
                    .PackagingCode = CleanCol(cols, 11),
                    .Capacity = If(cols.Length > 12, ParseInt(CleanCol(cols, 12)), Nothing),
                    .SupplierCode = CleanCol(cols, 13),
                    .SupplierFactoryCode = CleanCol(cols, 14),
                    .SupplierShippingLocation = CleanCol(cols, 15),
                    .DeliveryCode = CleanCol(cols, 16),
                    .DeliveryFactoryCode = CleanCol(cols, 17),
                    .DeliveryLocation = CleanCol(cols, 18),
                    .DirectDeliveryType = CleanCol(cols, 19),
                    .CustomerOrderNo = CleanCol(cols, 20),
                    .CustomerOrderNoProcessNo1 = CleanCol(cols, 21),
                    .CustomerOrderNoProcessNo2 = CleanCol(cols, 22),
                    .DeliveryType = CleanCol(cols, 23),
                    .OrderDataType = CleanCol(cols, 24),
                    .DeliveryDateType = CleanCol(cols, 25),
                    .OrderQtyType = CleanCol(cols, 26),
                    .DeliveryDate = If(cols.Length > 27, ParseDate(CleanCol(cols, 27)), Nothing),
                    .DeliveryTime = CleanCol(cols, 28),
                    .ShipInstructionsDate = If(cols.Length > 29, ParseDate(CleanCol(cols, 29)), Nothing),
                    .OrderQty = If(cols.Length > 30, ParseInt(CleanCol(cols, 30)), Nothing),
                    .ArrangeManager = CleanCol(cols, 31),
                    .OrderReason = CleanCol(cols, 32),
                    .FirstArticleType = CleanCol(cols, 33),
                    .OrderNotes = CleanCol(cols, 34),
                    .SupplyLine = CleanCol(cols, 35),
                    .SupplyProcess = CleanCol(cols, 36),
                    .ConstructionNo = CleanCol(cols, 37),
                    .PackingType = CleanCol(cols, 38),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

        ''' <summary>
        ''' Spirits663NAnd663SAnd664T 形式の CSV をパースします。
        ''' </summary>
        Public Shared Function ParseSpirits663NAnd663SAnd664T(filePath As String) As List(Of Spirits663NAnd663SAnd66Row)
            Dim list As New List(Of Spirits663NAnd663SAnd66Row)()
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.GetEncoding("shift-jis"))

            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim cols As String() = SplitCsvLine(line)
                If cols.Length < 10 Then Continue For

                Dim code = CleanCol(cols, 0)
                If code <> "663N" AndAlso code <> "663S" AndAlso code <> "664T" Then Continue For

                Dim row As New Spirits663NAnd663SAnd66Row() With {
                    .InfoTypeCode = CleanCol(cols, 0),
                    .DocTitleType = CleanCol(cols, 1),
                    .PaymentMethodWordType = CleanCol(cols, 2),
                    .ClientCode = CleanCol(cols, 3),
                    .ContractorCode = CleanCol(cols, 4),
                    .ContractorOfficeCode = CleanCol(cols, 5),
                    .PublicationDate = If(cols.Length > 6, ParseDate(CleanCol(cols, 6)), Nothing),
                    .TargetReferenceDateType = CleanCol(cols, 7),
                    .TargetReferenceDate = CleanCol(cols, 8),
                    .SupplierCode = CleanCol(cols, 9),
                    .DeliveryFormType = CleanCol(cols, 10),
                    .CustomerItemNo = CleanCol(cols, 11),
                    .ProcessType = CleanCol(cols, 12),
                    .CustomerItemNoProcessNo1 = CleanCol(cols, 13),
                    .CustomerOrderNo = CleanCol(cols, 14),
                    .PickupInstructionsTimes = CleanCol(cols, 15),
                    .OldDeliveryDate = If(cols.Length > 16, ParseDate(CleanCol(cols, 16)), Nothing),
                    .DeliveryDate = If(cols.Length > 17, ParseDate(CleanCol(cols, 17)), Nothing),
                    .OldOrderQty = If(cols.Length > 18, ParseLong(CleanCol(cols, 18)), Nothing),
                    .OrderQty = If(cols.Length > 19, ParseLong(CleanCol(cols, 19)), Nothing),
                    .CahngeReason = CleanCol(cols, 20),
                    .Reserve = CleanCol(cols, 21),
                    .ActiveFlag = "Y",
                    .CreatedAt = DateTime.Now
                }

                list.Add(row)
            Next

            Return list
        End Function

#Region "ヘルパーメソッド"

        Private Shared Function SplitCsvLine(line As String) As String()
            Dim result As New List(Of String)()
            Dim inQuotes As Boolean = False
            Dim current As New StringBuilder()

            For Each c As Char In line
                If c = """"c Then
                    inQuotes = Not inQuotes
                ElseIf c = ","c AndAlso Not inQuotes Then
                    result.Add(current.ToString())
                    current.Clear()
                Else
                    current.Append(c)
                End If
            Next
            result.Add(current.ToString())
            Return result.ToArray()
        End Function

        Private Shared Function CleanCol(cols As String(), idx As Integer) As String
            If idx >= cols.Length Then Return String.Empty
            Return cols(idx).Trim(" "c, """"c)
        End Function

        Private Shared Function ParseDate(val As String) As Date?
            If String.IsNullOrWhiteSpace(val) Then Return Nothing
            Dim d As Date
            Dim formats As String() = {"yyyyMMdd", "yyyy/MM/dd", "yyyy-MM-dd"}
            If Date.TryParseExact(val.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
                Return d
            End If
            If Date.TryParse(val.Trim(), d) Then
                Return d
            End If
            Return Nothing
        End Function

        Private Shared Function ParseInt(val As String) As Integer?
            If String.IsNullOrWhiteSpace(val) Then Return Nothing
            Dim res As Integer
            If Integer.TryParse(val.Trim(), res) Then Return res
            Return Nothing
        End Function

        Private Shared Function ParseLong(val As String) As Long?
            If String.IsNullOrWhiteSpace(val) Then Return Nothing
            Dim res As Long
            If Long.TryParse(val.Trim(), res) Then Return res
            Return Nothing
        End Function

#End Region

    End Class

End Namespace
