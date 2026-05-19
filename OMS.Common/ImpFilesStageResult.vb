
Namespace OMS.Common
    ''' <summary>
    ''' ファイル検索結果の1行（ファイル単位）
    ''' </summary>
    Public Class ImpFilesStageResult
        Public Property CustomerSettingId As Integer
        Public Property CustomerCode As String
        Public Property CustomerName As String
        Public Property ProfitCenter As String
        Public Property CustomerUnitId As String
        Public Property CustomerUnitName As String
        Public Property FolderType As Integer
        Public Property FolderPath As String
        Public Property FilePath As String
        Public Property FileName As String
        Public Property StagedFolderPath As String
        Public Property StagedFilePath As String
        Public Property StagedFileName As String
        Public Property Status As String
        Public Property LastWriteTime As DateTime?

        ' 画面選択に応じてセット
        Public Property ReconcileFlag As String         ' RECONCILE_FLAG
        Public Property FcstReconcileFlag As String     ' FCST_RECONCILE_FLAG
        Public Property HandFlag As String              ' HAND_FLAG

        ' エラー表示用（あれば）
        Public Property ErrorMessage As String
    End Class
    Public Class FolderPathInfo
        Public Property FolderPath As String
        Public Property FileName As String
        Public Property Staged_FolderPath As String
        Public Property Staged_FileName As String
        Public Property FolderType As Integer
    End Class

    Public Class MappingInfo
        Public Property CustomerSettingId As Long
        Public Property FileId As Long
        Public Property FolderType As Integer
        Public Property HeaderRowIndex As Integer
        Public Property DataStartRowIndex As Integer
        Public Property default_sheet_name As String
        Public Property format_type As String
        Public Property file_type As String
        Public Property delimiter As String
        Public Property enclosure As String
        Public Property header_flag As String
        Public Property line_ending As String
        Public Property charset As String
        Public Property profile_id As Integer
        Public Property target_field As String
        Public Property source_column_index As Integer
        Public Property source_header_name As String
        Public Property source_sheet_name As String
        Public Property source_cell_address As String
        Public Property row_selector_type As String
        Public Property row_selector_value As String
        Public Property data_type As String
        Public Property format_pattern As String
        Public Property trim_flag As String
        Public Property null_if As String
        Public Property default_value As String
    End Class

    Public Class OrderSummaryRow
        Public Property ItemNo As String           ' 品目No
        Public Property EarliestDueDate As DateTime  ' 集計後の希望納期（最古日）
        Public Property TotalDemandQty As Decimal   ' 品目No計需要数
    End Class

End Namespace
