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
                        ' 2026/6/5 データなし チェック
                        If (row.item.CustomerCode IsNot Nothing) Then
                            objSheet1.Cell(row.index, 1).Value = row.item.CustomerCode
                        End If
                        If (row.item.ProfitCenter IsNot Nothing) Then
                            objSheet1.Cell(row.index, 2).Value = row.item.ProfitCenter
                        End If
                        If (row.item.CustomerUnitName IsNot Nothing) Then
                            objSheet1.Cell(row.index, 3).Value = row.item.CustomerUnitName
                        End If
                        If (row.item.ItemNo IsNot Nothing) Then
                            objSheet1.Cell(row.index, 4).Value = row.item.ItemNo
                        End If
                        If (row.item.DueDate.HasValue) Then
                            objSheet1.Cell(row.index, 5).Value = row.item.DueDate
                        End If
                        If (row.item.DemandQty.HasValue) Then
                            objSheet1.Cell(row.index, 6).Value = row.item.DemandQty
                        End If
                        If (row.item.PreviousData.HasValue) Then
                            objSheet1.Cell(row.index, 7).Value = row.item.PreviousData
                        End If
                        If (row.item.CurrentData.HasValue) Then
                            objSheet1.Cell(row.index, 8).Value = row.item.CurrentData
                        End If
                        If (row.item.DiffFromPrev.HasValue) Then
                            objSheet1.Cell(row.index, 9).Value = row.item.DiffFromPrev
                        End If
                    Next
                    ' (差異リスト:確定納入指示) 
                    Dim objSheet2 As IXLWorksheet = objWBook.Worksheets.Add(sheetName2)

                    For Each itemp In DeliveryInstructionTitle
                        objSheet2.Column(itemp.index).Width = itemp.width
                        objSheet2.Cell(1, itemp.index).Value = itemp.tString
                    Next
                    For Each row In instructionDifferenceRows.Select(Function(val, idx) (item:=val, index:=idx))
                        ' 2026/6/5 データなし チェック
                        If (row.item.CustomerCode IsNot Nothing) Then
                            objSheet2.Cell(row.index, 1).Value = row.item.CustomerCode
                        End If
                        If (row.item.ProfitCenter IsNot Nothing) Then
                            objSheet2.Cell(row.index, 2).Value = row.item.ProfitCenter
                        End If
                        If (row.item.CustomerUnitName IsNot Nothing) Then
                            objSheet2.Cell(row.index, 3).Value = row.item.CustomerUnitName
                        End If
                        If (row.item.ItemNo IsNot Nothing) Then
                            objSheet2.Cell(row.index, 4).Value = row.item.ItemNo
                        End If
                        If (row.item.CustomerOrderNo IsNot Nothing) Then
                            objSheet2.Cell(row.index, 5).Value = row.item.CustomerOrderNo
                        End If
                        If (row.item.DueDate.HasValue) Then
                            objSheet2.Cell(row.index, 6).Value = row.item.DueDate
                        End If
                        If (row.item.DemandQty.HasValue) Then
                            objSheet2.Cell(row.index, 7).Value = row.item.DemandQty
                        End If
                        If (row.item.PreviousData.HasValue) Then
                            objSheet2.Cell(row.index, 8).Value = row.item.PreviousData
                        End If
                        If (row.item.CurrentData.HasValue) Then
                            objSheet2.Cell(row.index, 9).Value = row.item.CurrentData
                        End If
                        If (row.item.DiffFromPrev.HasValue) Then
                            objSheet2.Cell(row.index, 10).Value = row.item.DiffFromPrev
                        End If
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

        '''' <summary>
        '''' Work path 取得 (ない場合は作成する)
        '''' </summary>
        '''' <returns></returns>
        'Public Shared Function GetWorkPath() As String

        '    Dim rawWork As String = GetWorkFolderRoot()
        '    Dim workFolder = rawWork & "FileTempFolder"
        '    If (Not IO.File.Exists(workFolder)) Then
        '        EnsureDirectory(workFolder)
        '    End If
        '    Return workFolder
        'End Function

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
