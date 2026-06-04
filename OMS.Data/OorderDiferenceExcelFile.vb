Imports System.Configuration
Imports System.Data
Imports System.Text
Imports System.Web
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.SqlServer
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data
    Public Class OrderDiferenceExcelFile
        Implements IDisposable

        Public strErrMsg As String

        Public Enum DiffFileTiminge
            AfterReceivingAnOrder = 0       ' 受注取込み後
            AfterProductionPlanning = 1     ' 生産計画後
        End Enum

        ''' <summary>
        ''' IDisposable 実装
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            'Throw New NotImplementedException()
        End Sub
        '''' <summary>
        '''' Shared 受注取込後 差分Excel ファイル作成
        '''' </summary>
        '''' <param name="situation"></param>
        '''' <param name="updateDate"></param>
        '''' <param name="unNoticeDifferenceRows"></param>
        '''' <param name="instructionDifferenceRows"></param>
        '''' <returns>ファイル作成が成功したときファイル名を返す</returns>
        'Public Shared Function AfterOrderCreateOorderDiferenceExcelFile(situation As DiffFileTiminge, updateDate As Date, unNoticeDifferenceRows As IEnumerable(Of AfterOrderUnNoticeDifferenceRow), instructionDifferenceRows As IEnumerable(Of AfterOrderInstructionDifferenceRow)) As String

        '    Dim customerCode = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).CustomerCode, Nothing)
        '    Dim profitCenter = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).ProfitCenter, Nothing)
        '    Dim customerUnitName = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).CustomerUnitName, Nothing)
        '    Dim filename = GetOorderDiferenceExcelFilename(situation, updateDate, customerCode, profitCenter, customerUnitName)
        '    Dim index = 1
        '    Using excelfile As New OrderDiferenceExcelFile()
        '        If (excelfile.AfterOrderDiferenceExcelFile(filename, unNoticeDifferenceRows, instructionDifferenceRows)) Then
        '        Else
        '            filename = ""
        '        End If
        '    End Using

        '    Return filename

        'End Function

        ''' <summary>
        ''' Shared 差分Excel ファイル作成
        ''' </summary>
        ''' <param name="situation"></param>
        ''' <param name="updateDate"></param>
        ''' <param name="unNoticeDifferenceRows"></param>
        ''' <param name="instructionDifferenceRows"></param>
        ''' <returns>ファイル作成が成功したときファイル名を返す</returns>
        Public Shared Function CreateOorderDiferenceExcelFile(strPath As String, situation As DiffFileTiminge, updateDate As Date, unNoticeDifferenceRows As IEnumerable(Of UnNoticeDifferenceRow), instructionDifferenceRows As IEnumerable(Of InstructionDifferenceRow), customerSettingId As Long) As String

            Dim customerCode = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).CustomerCode, customerSettingId.ToString())
            Dim profitCenter = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).ProfitCenter, Nothing)
            Dim customerUnitName = If(unNoticeDifferenceRows.Count > 0, unNoticeDifferenceRows(0).CustomerUnitName, Nothing)
            Dim filename = GetOorderDiferenceExcelFilename(strPath, situation, updateDate, customerCode, profitCenter, customerUnitName, customerSettingId)
            Dim index = 1
            Using excelfile As New OrderDiferenceExcelFile()
                If (excelfile.OrderDiferenceExcelFile(filename, situation, unNoticeDifferenceRows, instructionDifferenceRows)) Then
                Else
                    filename = ""
                End If
            End Using

            Return filename

        End Function

        ''' <summary>
        ''' 差異リストファイル名取得
        ''' </summary>
        ''' <param name="situation"></param>
        ''' <param name="updateDate"></param>
        ''' <returns></returns>
        Public Shared Function GetOorderDiferenceExcelFilename(strPath As String, situation As DiffFileTiminge, updateDate As Date, customerCode As String, profitCenter As String, customerUnitName As String, Optional customerSettingId As Long = 0) As String

            '' [STRA納期設定ページ_受注取込後] yyyyMMdd_[取引先コード]_[PC]_[注文工場／担当者名]_受注差異リスト.xlsx
            '' [STRA納期設定ページ_生産計画後] yyyyMMdd_[取引先コード]_[PC]_[注文工場／担当者名]_生産計画差異リスト.xlsx

            ' 2026/6/4
            ' [STRA納期設定ページ_受注取込後]YYYMMDD_[取引先設定ID]_受注取込差異リスト.xlsx
            ' [STRA納期設定ページ_生産計画後]YYYMMDD_[取引先設定ID]_生産計画差異リスト.xlsx
            customerCode = If(customerCode Is Nothing, "", "_" & customerCode)
            profitCenter = If(profitCenter Is Nothing, "", "_" & profitCenter)
            customerUnitName = If(customerUnitName Is Nothing, "", "_" & customerUnitName)
            Dim tCustomerSettingId = If(customerSettingId = 0, "", customerSettingId.ToString())
            ' 2026/06/4
            customerCode = ""
            profitCenter = ""
            customerUnitName = ""

            Dim strFile = $"{updateDate.ToString("yyyyMMdd_")}{customerCode}{profitCenter}{customerUnitName}{tCustomerSettingId}{If(situation = DiffFileTiminge.AfterReceivingAnOrder, "_受注差異リスト.xlsx", "_生産計画差異リスト.xlsx")}"
            ' Server 側の File 保存Folder
            'Dim strPath = GetWorkPath()
            If (Not IO.File.Exists(strPath)) Then
                EnsureDirectory(strPath)
            End If

            Dim filename = IO.Path.Combine(strPath, strFile)
            Return filename

        End Function

        '#If True Then
        '        Private AfterOrderUnofficialNoticeTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
        '                    (1, 15, "取引先コード"),
        '                    (2, 15, "PC"),
        '                    (3, 15, "注文工場／担当者名"),
        '                    (4, 15, "品目No"),
        '                    (5, 15, "希望納期"),
        '                    (6, 15, "数量"),
        '                    (7, 15, "前回件数"),
        '                    (8, 15, "今回件数"),
        '                    (9, 15, "前回比")
        '        }
        '#Else
        '        Private AfterOrderUnofficialNoticeTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
        '                    (1, 15, "取引先設定ID"),
        '                    (2, 15, "品目No"),
        '                    (3, 15, "日割前納期"),
        '                    (4, 15, "日割前受注数"),
        '                    (5, 15, "前回件数"),
        '                    (6, 15, "今回件数"),
        '                    (7, 15, "差異")
        '        }
        '#End If
        '        ''' <summary>
        '        ''' (受注後 差異リスト:確定納入指示) 
        '        ''' Index(Column), Cell width, Title
        '        ''' </summary>
        '#If True Then
        '        Private AfterOrderDeliveryInstructionTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
        '                    (1, 15, "取引先設定ID"),
        '                    (2, 15, "品目No"),
        '                    (3, 15, "客先発注No"),
        '                    (4, 15, "希望納期"),
        '                    (5, 15, "数量"),
        '                    (6, 15, "前回件数"),
        '                    (7, 15, "今回件数"),
        '                    (8, 15, "前回比")
        '        }
        '#Else
        '        Private AfterOrderDeliveryInstructionTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
        '                    (1, 15, "取引先設定ID"),
        '                    (2, 15, "品目No"),
        '                    (3, 15, "客先発注No"),
        '                    (4, 15, "日割前納期"),
        '                    (5, 15, "日割前受注数"),
        '                    (6, 15, "前回件数"),
        '                    (7, 15, "今回件数"),
        '                    (8, 15, "差異")
        '        }
        '#End If
        '''' <summary>
        '''' 差異Excel ファイル作成
        '''' </summary>
        '''' <param name="filename"></param>
        '''' <param name="unNoticeDifferenceRows"></param>
        '''' <param name="instructionDifferenceRows"></param>
        '''' <returns></returns>
        'Public Function AfterOrderDiferenceExcelFile(filename As String, unNoticeDifferenceRows As IEnumerable(Of AfterOrderUnNoticeDifferenceRow), instructionDifferenceRows As IEnumerable(Of AfterOrderInstructionDifferenceRow)) As Boolean
        '    Dim rt = False

        '    If (IO.File.Exists(filename)) Then
        '        IO.File.Delete(filename)
        '    End If

        '    Try
        '        'ワークブックを作成
        '        Using objWBook As New XLWorkbook
        '            ' (差異リスト:内示) 
        '            Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add("差異リスト-内示")
        '            'Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add()

        '            For Each itemp In AfterOrderUnofficialNoticeTitle
        '                objSheet1.Column(itemp.index).Width = itemp.width
        '                objSheet1.Cell(1, itemp.index).Value = itemp.tString
        '            Next
        '            Dim offset = 1  ' Title 行 のかさまし分
        '            For Each row In unNoticeDifferenceRows.Select(Function(val, idx) (item:=val, index:=idx + 1 + offset))
        '                objSheet1.Cell(row.index, 1).Value = row.item.CustomerCode
        '                objSheet1.Cell(row.index, 2).Value = row.item.ProfitCenter
        '                objSheet1.Cell(row.index, 3).Value = row.item.CustomerUnitName
        '                objSheet1.Cell(row.index, 4).Value = row.item.ItemNo
        '                objSheet1.Cell(row.index, 5).Value = row.item.PreDailyDeliveryDate
        '                objSheet1.Cell(row.index, 6).Value = row.item.PreDailyOrderQty
        '                objSheet1.Cell(row.index, 7).Value = row.item.PreviousData
        '                objSheet1.Cell(row.index, 8).Value = row.item.CurrentData
        '                objSheet1.Cell(row.index, 9).Value = row.item.DiffFromPrev
        '            Next

        '            ' (差異リスト:確定納入指示) 
        '            Dim objSheet2 As IXLWorksheet = objWBook.Worksheets.Add("差異リスト-確定納入指示")

        '            For Each itemp In AfterOrderDeliveryInstructionTitle
        '                objSheet2.Column(itemp.index).Width = itemp.width
        '                objSheet2.Cell(1, itemp.index).Value = itemp.tString
        '            Next
        '            offset = 1  ' Title 行 のかさまし分
        '            For Each row In instructionDifferenceRows.Select(Function(val, idx) (item:=val, index:=idx + 1 + offset))
        '                objSheet1.Cell(row.index, 1).Value = row.item.CustomerSettingId
        '                objSheet1.Cell(row.index, 2).Value = row.item.ItemNo
        '                objSheet1.Cell(row.index, 3).Value = row.item.CustomerOrderNo
        '                objSheet1.Cell(row.index, 4).Value = row.item.PreDailyDeliveryDate
        '                objSheet1.Cell(row.index, 5).Value = row.item.PreDailyOrderQty
        '                objSheet1.Cell(row.index, 6).Value = row.item.PreviousData
        '                objSheet1.Cell(row.index, 7).Value = row.item.CurrentData
        '                objSheet1.Cell(row.index, 8).Value = row.item.DiffFromPrev
        '            Next
        '            'ファイルに保存 
        '            objWBook.SaveAs(filename)
        '        End Using
        '        rt = True
        '    Catch ex As Exception
        '        Dim m = ex.Message
        '        rt = False
        '    End Try
        '    Return rt

        'End Function
        ''' <summary>
        ''' (生産計画差異リスト:内示) 
        ''' Index(Column), Cell width, Title
        ''' </summary>
        Private UnofficialNoticeTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
                    (1, 15, "取引先コード"),
                    (2, 15, "PC"),
                    (3, 15, "注文工場/担当者名"),
                    (4, 15, "品目No"),
                    (5, 15, "希望納期"),
                    (6, 15, "数量"),
                    (7, 15, "前回_件数"),
                    (8, 15, "今回_件数"),
                    (9, 15, "前回比")
        }
        ''' <summary>
        ''' (生産計画差異リスト:確定納入指示) 
        ''' Index(Column), Cell width, Title
        ''' </summary>
        Private DeliveryInstructionTitle As New List(Of (index As Integer, width As Integer, tString As String)) From {
                    (1, 15, "取引先コード"),
                    (2, 15, "PC"),
                    (3, 15, "注文工場/担当者名"),
                    (4, 15, "品目No"),
                    (5, 15, "客先発注No"),
                    (6, 15, "希望納期"),
                    (7, 15, "数量"),
                    (8, 15, "前回_件数"),
                    (9, 15, "今回_件数"),
                    (10, 15, "前回比")
        }
        ''' <summary>
        ''' 差異Excel ファイル作成
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="unNoticeDifferenceRows"></param>
        ''' <param name="instructionDifferenceRows"></param>
        ''' <returns></returns>
        Public Function OrderDiferenceExcelFile(filename As String, situation As DiffFileTiminge, unNoticeDifferenceRows As IEnumerable(Of UnNoticeDifferenceRow), instructionDifferenceRows As IEnumerable(Of InstructionDifferenceRow)) As Boolean
            Dim rt = False

            If (IO.File.Exists(filename)) Then
                IO.File.Delete(filename)
            End If

            Try
                ' 2026/06/04
                ' 受注取り込み後
                'シート名（内示）受注取込差異リスト-内示
                'シート名（確定/納入指示）：受注取込差異リスト-確定納入指示
                ' 生産計画後
                'シート名（内示）生産計画差異リスト-内示
                'シート名（確定/納入指示）生産計画差異リスト-確定納入指示
                Dim sheetName1 = If(situation = DiffFileTiminge.AfterReceivingAnOrder, "受注取込差異リスト-内示", "生産計画差異リスト-内示")
                Dim sheetName2 = If(situation = DiffFileTiminge.AfterReceivingAnOrder, "受注取込差異リスト-確定納入指示", "生産計画差異リスト-確定納入指示")

                'ワークブックを作成
                Using objWBook As New XLWorkbook
                    ' (差異リスト:内示) 
                    Dim objSheet1 As IXLWorksheet = objWBook.Worksheets.Add(sheetName1)

                    For Each itemp In UnofficialNoticeTitle
                        objSheet1.Column(itemp.index).Width = itemp.width
                        objSheet1.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    For Each row In unNoticeDifferenceRows.Select(Function(val, idx) (item:=val, index:=idx))
                        objSheet1.Cell(row.index, 1).Value = row.item.CustomerCode
                        objSheet1.Cell(row.index, 2).Value = row.item.ProfitCenter
                        objSheet1.Cell(row.index, 3).Value = row.item.CustomerUnitName
                        objSheet1.Cell(row.index, 4).Value = row.item.ItemNo
                        objSheet1.Cell(row.index, 5).Value = row.item.DueDate
                        objSheet1.Cell(row.index, 6).Value = row.item.DemandQty
                        objSheet1.Cell(row.index, 7).Value = row.item.PreviousData
                        objSheet1.Cell(row.index, 8).Value = row.item.CurrentData
                        objSheet1.Cell(row.index, 9).Value = row.item.DiffFromPrev
                    Next
                    ' (差異リスト:確定納入指示) 
                    Dim objSheet2 As IXLWorksheet = objWBook.Worksheets.Add(sheetName2)

                    For Each itemp In DeliveryInstructionTitle
                        objSheet2.Column(itemp.index).Width = itemp.width
                        objSheet2.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    For Each row In instructionDifferenceRows.Select(Function(val, idx) (item:=val, index:=idx))
                        objSheet2.Cell(row.index, 1).Value = row.item.CustomerCode
                        objSheet2.Cell(row.index, 2).Value = row.item.ProfitCenter
                        objSheet2.Cell(row.index, 3).Value = row.item.CustomerUnitName
                        objSheet2.Cell(row.index, 4).Value = row.item.ItemNo
                        objSheet2.Cell(row.index, 5).Value = row.item.CustomerOrderNo
                        objSheet2.Cell(row.index, 6).Value = row.item.DueDate
                        objSheet2.Cell(row.index, 7).Value = row.item.DemandQty
                        objSheet2.Cell(row.index, 8).Value = row.item.PreviousData
                        objSheet2.Cell(row.index, 9).Value = row.item.CurrentData
                        objSheet2.Cell(row.index, 10).Value = row.item.DiffFromPrev
                    Next
                    'ファイルに保存 
                    objWBook.SaveAs(filename)
                End Using
                rt = True
            Catch ex As Exception
                Dim m = ex.Message
                rt = False
            End Try
            Return rt

        End Function

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

#If False Then
        ''' <summary>
        ''' 差異Excelファイル作成
        ''' </summary>
        ''' <param name="situation">false: 生産計画差異リスト true: 受注差異リスト.xls</param>
        ''' <param name="updateDate"></param>
        ''' <param name="odersRow">Order list</param>
        ''' <param name="customerSettingId">取引先コード</param>
        ''' <param name="customerCode">Profit center</param>
        ''' <param name="customerUnitName">工場／担当者名</param>
        ''' <returns></returns>
        Public Shared Function CreateOorderDiferenceExcelFile(situation As Boolean, updateDate As Date, odersRow As List(Of OrdersRow), customerSettingId As Long, customerCode As String, customerUnitName As String) As String

            Dim rt = False
            Dim filename = GetOorderDiferenceExcelFilename(situation, updateDate)
            Dim index = 1
            Using excelfile As New OorderDiferenceExcelFile()
                For Each oderRow In odersRow
                    rt = excelfile.OorderDiferenceExcelFile(filename, oderRow, index, customerSettingId, customerCode, customerUnitName)
                    If (rt = False) Then
                        Exit For
                    End If
                Next
            End Using

            Return rt

        End Function
        ''' <summary>
        ''' ExcelFile 保存 初回
        ''' </summary>
        ''' <param name="filename">ファイル名</param>
        ''' <param name="oderRow">Order class</param>
        ''' <param name="customerSettingId">取引先コード</param>
        ''' <param name="customerCode">Profit center</param>
        ''' <param name="customerUnitName">工場／担当者名</param>
        ''' <returns></returns>
        Public Function OorderDiferenceNewExcelFile(filename As String, oderRow As OrdersRow, customerSettingId As Long, customerCode As String, customerUnitName As String) As Boolean

            Dim rt = False

            Try
                'ワークブックを作成
                Dim objWBook As New XLWorkbook

                'ワークシートを定義して「sheet1」を作成し追加
                Dim objSheet As IXLWorksheet = objWBook.Worksheets.Add("sheet1")



                'セルの幅を設定
                objSheet.Column(1).Width = 10
                objSheet.Column(2).Width = 10
                objSheet.Column(3).Width = 10
                objSheet.Column(4).Width = 37
                objSheet.Column(5).Width = 37
                objSheet.Column(6).Width = 15
                objSheet.Column(7).Width = 20
                objSheet.Column(8).Width = 15
                objSheet.Column(9).Width = 15
                objSheet.Column(10).Width = 15
                objSheet.Column(11).Width = 15
                objSheet.Column(12).Width = 15
                objSheet.Column(13).Width = 15
                objSheet.Column(14).Width = 15
                objSheet.Column(15).Width = 15
                objSheet.Column(16).Width = 15
                objSheet.Column(17).Width = 15

                'ヘッダをセット
                objSheet.Cell(1, 1).Value = "取引先コード"
                objSheet.Cell(1, 2).Value = "PC"
                objSheet.Cell(1, 3).Value = "工場／担当者名"
                objSheet.Cell(1, 4).Value = "客先発注No"
                objSheet.Cell(1, 5).Value = "受注日"
                objSheet.Cell(1, 6).Value = "希望納期"
                objSheet.Cell(1, 7).Value = "客先品目No"
                objSheet.Cell(1, 8).Value = "品目No"
                objSheet.Cell(1, 9).Value = "需要数"
                objSheet.Cell(1, 10).Value = "通貨コード"
                objSheet.Cell(1, 11).Value = "製品コード"
                objSheet.Cell(1, 12).Value = "コメント"
                objSheet.Cell(1, 13).Value = "納入先コード"
                objSheet.Cell(1, 14).Value = "受注区分"
                objSheet.Cell(1, 15).Value = "分割区分"
                objSheet.Cell(1, 16).Value = "情報区分"
                objSheet.Cell(1, 17).Value = "自社予測フラグ"
        ' 行データ追加
        Dim index = 2
                OrderDiferenceExcelFileAddRow(objWBook, index, oderRow, customerSettingId, customerCode, customerUnitName)
                'ファイルに保存 
                objWBook.SaveAs(filename)
                rt = True

            Catch ex As Exception

                strErrMsg = "エクセルデータ作成中にエラーが発生しました" & ex.Message

            Finally

            End Try

            Return rt
        End Function
        ''' <summary>
        ''' function 行追加
        ''' </summary>
        ''' <param name="objWBook"></param>
        ''' <param name="i"></param>
        ''' <param name="oderRow"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="customerCode"></param>
        ''' <param name="customerUnitName"></param>
        ''' <returns></returns>
        Public Function OrderDiferenceExcelFileAddRow(objWBook As XLWorkbook, i As Integer, oderRow As OrdersRow, customerSettingId As Long, customerCode As String, customerUnitName As String) As String

            'ワークシートを定義して「sheet1」を作成し追加
            Dim objSheet As IXLWorksheet = objWBook.Worksheets("sheet1")
            'セルの文字位置を設定
            objSheet.Range(1, 1, 1, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
            '印刷用紙をA4に指定
            'objSheet.PageSetup.PaperSize = XLPaperSize.A4Paper
            '印刷向きを横に指定
            'objSheet.PageSetup.PageOrientation = XLPageOrientation.Landscape
            'Dim i = 2
            objSheet.Cell(i, 1).Value = customerSettingId
            objSheet.Cell(i, 2).Value = customerCode
            objSheet.Cell(i, 3).Value = customerUnitName
            objSheet.Cell(i, 4).Value = oderRow.CustomerOrderNo
            objSheet.Cell(i, 5).Value = oderRow.OrderDate
            objSheet.Cell(i, 6).Value = oderRow.DueDate
            objSheet.Cell(i, 7).Value = oderRow.CustomerItemNo
            objSheet.Cell(i, 8).Value = oderRow.ItemNo
            objSheet.Cell(i, 9).Value = oderRow.DemandQty
            objSheet.Cell(i, 10).Value = oderRow.CurrencyCode
            objSheet.Cell(i, 11).Value = oderRow.ProductCode
            objSheet.Cell(i, 12).Value = oderRow.Remarks
            objSheet.Cell(i, 13).Value = oderRow.DeliveryCode
            objSheet.Cell(i, 14).Value = oderRow.OrderType
            objSheet.Cell(i, 15).Value = oderRow.ProratedType
            objSheet.Cell(i, 16).Value = oderRow.InfoType
            objSheet.Cell(i, 17).Value = oderRow.SelfFcstFlag
            Return True

        End Function

        ''' <summary>
        ''' ExcelFile 保存 2行目以降
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="oderRow"></param>
        ''' <param name="index"></param>
        ''' <param name="customerSettingId"></param>
        ''' <param name="customerCode"></param>
        ''' <param name="customerUnitName"></param>
        ''' <returns></returns>
        Public Function OrderDiferenceExcelFileNext(filename As String, oderRow As OrdersRow, index As Integer, customerSettingId As Long, customerCode As String, customerUnitName As String) As Boolean

            Dim rt = False
            Try
                'ワークブックを作成
                Dim objWBook As New XLWorkbook(filename)

                ' 行データ追加
                index += 2
                OrderDiferenceExcelFileAddRow(objWBook, index, oderRow, customerSettingId, customerCode, customerUnitName)

                'ファイルに保存 
                objWBook.SaveAs(filename)
                rt = True
            Catch ex As Exception

                strErrMsg = "エクセルデータ作成中にエラーが発生しました" & ex.Message

            Finally

            End Try

        End Function
#End If

#If False Then
        '格納データと　テーブル
        '取引先コード           CUSTOMER_SETTING_MST.CUSTOMER_CODE  ORDERS.CUSTOMER_CODE
        'PC                     CUSTOMER_SETTING_MST.PROFIT_CENTER
        '工場/担当者名          CUSTOMER_UNIT_MST.CUSTOMER_UNIT_NAME (CUSTOMER_SETTING_MST.CUSTOMER_UNIT_ID-CUSTOMER_UNIT_MST.CUSTOMER_UNIT_ID)
        '客先発注No             ORDERS.CUSTOMER_ORDER_NO
        '受注日                 ORDERS.ORDER_DATE
        '希望納期               ORDERS.DUE_DATE
        '客先品目No             ORDERS.CUSTOMER_ITEM_NO
        '品目No                 ORDERS.ITEM_NO
        '需要数                 ORDERS.DEMAND_QTY
        '通貨コード             ORDERS.CURRENCY_CODE
        '製品コード             ORDERS.PRODUCT_CODE
        'コメント               ORDERS.REMARKS
        '納入先コード           ORDERS.DELIVERY_CODE
        '受注区分               ORDERS.ORDER_TYPE
        '分割区分               ORDERS.PRORATED_TYPE
        '情報区分               ORDERS.INFO_TYPE
        '自社予測フラグ         ORDERS.SELF_FCST_FLAG
#End If


        '(差異リスト:内示)
        '取引先設定ID
        '取引先コード
        'PC
        '注文工場／担当者名
        '品目No
        '希望納期
        '数量
        '前回件数
        '今回件数
        '前回比
        '
        '(差異リスト:確定納入指示)
        '取引先設定ID
        '取引先コード
        'PC
        '注文工場／担当者名
        '品目No
        '客先発注No
        '希望納期
        '数量
        '前回件数
        '今回件数
        '前回比

    End Class
End Namespace
