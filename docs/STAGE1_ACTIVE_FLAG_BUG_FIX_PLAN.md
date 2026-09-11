# Implementation Plan: Fix `UpdateActiveFlags` Invalidating In-File Rows (`ACTIVE_FLAG = 'N'`)

## Problem Summary

When pre-importing an `0600` file (`0600-260703.csv`) containing 10 rows, all 10 rows are inserted into `SUZUKI_SPIRITS_0600AND0630`, but 9 rows are set to `ACTIVE_FLAG = 'N'` and only the last row remains `ACTIVE_FLAG = 'Y'`.

### Root Cause Analysis

In `OMS.Business/Services/SuzukiPreImportService.vb` (`UpdateActiveFlags`):
```sql
UPDATE {tableName} old_t
SET old_t.active_flag = 'N'
WHERE old_t.active_flag = 'Y'
  AND EXISTS (
    SELECT 1 FROM {tableName} new_t
    WHERE new_t.customer_item_no = old_t.customer_item_no
      AND new_t.publication_date = old_t.publication_date
      AND new_t.publication_time = old_t.publication_time
      AND (new_t.created_at > old_t.created_at OR (new_t.created_at = old_t.created_at AND new_t.info_type_id > old_t.info_type_id))
  )
```

1. **Intra-File Tie-Breaker Bug**:
   All 10 rows in `0600` share the same `CUSTOMER_ITEM_NO` (`09407-17403-000`), `PUBLICATION_DATE` (`20260702`), `PUBLICATION_TIME` (`2028`), and `CREATED_AT` (`DateTime.Now` set during the batch insert).
   Because of `OR (new_t.created_at = old_t.created_at AND new_t.info_type_id > old_t.info_type_id)`:
   - Row 1 matches Row 2 (`info_type_id` 2 > 1) -> Row 1 becomes `ACTIVE_FLAG = 'N'`.
   - Row 2 matches Row 3 (`info_type_id` 3 > 2) -> Row 2 becomes `ACTIVE_FLAG = 'N'`.
   - ...
   - Row 9 matches Row 10 (`info_type_id` 10 > 9) -> Row 9 becomes `ACTIVE_FLAG = 'N'`.
   - Only Row 10 (the last row) has no higher `info_type_id`, so only Row 10 remains `ACTIVE_FLAG = 'Y'`.

2. **Missing `IMP_FILE_ID` Scoping**:
   `UpdateActiveFlags` did not receive or filter by `impFileStageId`. Thus, records inside the current file were evaluated against each other instead of only checking against records from **previous** files.

### Official Specification Reference
According to `docs/SUZUKI_SPIRITS_DATA_PRE_IMPORT.pdf` (Page 4, 「更新チェック」 & 「データ無効処理」):
- **合致条件**: 過去の `CREATED_AT`（登録日時）で `ACTIVE_FLAG` が `[Y:有効]` になっている同一部品番号・発行日時データ
- **データ無効処理**: 追加したデータよりも **過去日時のもの** (`CREATED_AT`) の `ACTIVE_FLAG` を `'N'` に更新する。

The specification strictly dictates that invalidation targets **prior** imports (`追加したデータよりも過去日時のもの`), never records within the current file.

---

## Proposed Changes (For `feature/suzuki-pre-import` branch)

### Component: Pre-Import Service (`OMS.Business`)

#### `OMS.Business/Services/SuzukiPreImportService.vb`

1. Update the call site in `PreImportSuzukiFiles`:
   ```vb
   ' 無効化チェック (ACTIVE_FLAG = 'N')
   UpdateActiveFlags(conn, tran, infoCode, impFileStageId)
   ```

2. Update `UpdateActiveFlags` signature and query:
   ```vb
   Private Sub UpdateActiveFlags(conn As OracleConnection, tran As OracleTransaction, infoCode As String, impFileStageId As Long)
       Dim tableName = GetTargetTableName(infoCode)
       If String.IsNullOrEmpty(tableName) Then Return

       Dim sql As New StringBuilder()
       sql.AppendLine($"UPDATE {tableName} old_t")
       sql.AppendLine("SET old_t.active_flag = 'N'")
       sql.AppendLine("WHERE old_t.active_flag = 'Y'")
       If impFileStageId > 0 Then
           ' 現在取込中のファイルの行は無効化対象外とする (Do not invalidate rows within current file)
           sql.AppendLine("  AND old_t.imp_file_id <> :p_imp_file_id")
       End If
       sql.AppendLine($"  AND EXISTS (")
       sql.AppendLine($"    SELECT 1 FROM {tableName} new_t")
       sql.AppendLine($"    WHERE new_t.customer_item_no = old_t.customer_item_no")

       If impFileStageId > 0 Then
           ' 新規追加されたファイルの行とマッチング (Match against newly added file rows)
           sql.AppendLine("      AND new_t.imp_file_id = :p_imp_file_id")
       End If

       If infoCode = "0740" Then
           sql.AppendLine("      AND new_t.acceptance_date = old_t.acceptance_date")
           sql.AppendLine("      AND new_t.acceptance_time = old_t.acceptance_time")
       ElseIf {"6604", "6624", "6634", "663N", "663S", "664T"}.Contains(infoCode) Then
           sql.AppendLine("      AND new_t.publication_date = old_t.publication_date")
       Else
           sql.AppendLine("      AND new_t.publication_date = old_t.publication_date")
           sql.AppendLine("      AND new_t.publication_time = old_t.publication_time")
       End If

       If impFileStageId > 0 Then
           sql.AppendLine($"      AND new_t.created_at >= old_t.created_at")
       Else
           sql.AppendLine($"      AND new_t.created_at > old_t.created_at")
       End If
       sql.AppendLine($"  )")

       Using cmd As New OracleCommand(sql.ToString(), conn)
           cmd.Transaction = tran
           If impFileStageId > 0 Then
               cmd.Parameters.Add(":p_imp_file_id", OracleDbType.Int64).Value = impFileStageId
           End If
           cmd.ExecuteNonQuery()
       End Using
   End Sub
   ```

---

## Verification Plan

### Automated Tests
1. Rebuild solution:
   ```powershell
   & "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" OrderManagementSystem.sln /t:Build /p:Configuration=Debug /v:m
   ```
2. Run unit tests in `OMS.Tests`:
   ```powershell
   & "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "OMS.Tests\bin\Debug\net48\OMS.Tests.dll" --TestCaseFilter:"FullyQualifiedName~Suzuki"
   ```

### Manual Verification
1. Place `0600-260703.csv` (10 rows) in the Suzuki source directory.
2. Run **[取込準備]** on `OrderImportStage.aspx`.
3. Check `SUZUKI_SPIRITS_0600AND0630` in the database:
   ```sql
   SELECT info_type_id, customer_item_no, delivery_date, order_qty, active_flag, imp_file_id
   FROM suzuki_spirits_0600and0630
   WHERE imp_file_id = (SELECT MAX(imp_file_stage_id) FROM imp_files_stage);
   ```
   Verify that **all 10 rows** have `ACTIVE_FLAG = 'Y'`.
4. Re-import the file and verify that the 10 rows from the previous `imp_file_id` are updated to `ACTIVE_FLAG = 'N'`, while the 10 rows from the new import are `ACTIVE_FLAG = 'Y'`.
