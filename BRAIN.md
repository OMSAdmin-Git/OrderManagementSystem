# Project Brain & Memory

## Current Focus
- **Active Goal:** Suzuki (SPIRITS) Order Import Pipeline Implementation (`受注取込_スピリッツ`) across Web UI (`OMS.Web`) and Batch Processing (`SUZUKI_SPIRITS_DATA_IMPORT.exe`).
- **Current Status:** Stage 1 Pre-Import (`SuzukiPreImportService.vb` / `SUZUKI_SPIRITS_DATA_PRE_IMPORT`) is fully implemented, aligned with Sakai-san's 4-rule error handling flow, and verified with unit tests. Resolved the bug where `SUZUKI_SPIRITS_6624` unmapped part rows (`INFO_TYPE_ID = 3908`) failed to report errors due to column index misalignment. All tests pass and the solution builds cleanly with 0 errors.
- **Immediate Next Step:** Await confirmation from Sakai-san on testing `6624-260715.csv` with the updated column indexing and 4-rule flow, then proceed to Stage 2 (`SuzukiDataImportService.vb` / `ORDERS_STAGE` transform & reconciliation).

## Architecture & Conventions
- **Stack:** Visual Basic .NET (VB.NET), .NET Framework 4.8, ASP.NET Web Forms, Oracle Database (ODP.NET / custom DAL with parameter binding), Windows Server 2022 / IIS 10, Windows Task Scheduler.
- **Core Files / Modules:**
  - `OMS.Web`: Web presentation layer (`Pages/Orders/OrderImportStage.aspx`, `Pages/Orders/OrderImport.aspx`, `Pages/Login/Login.aspx`).
  - `OMS.Business/Services/SuzukiPreImportService.vb`: Stage 1 Pre-Import service handling file intake, WORK folder staging, dedicated table inserts, active flag invalidation, and ASTI part matching with Z-suffix logic.
  - `OMS.Business/Services/SuzukiCsvParser.vb`: CSV parsing logic handling line splitting, encoding auto-detection (UTF-8 with BOM / Shift-JIS), and column sanitization.
  - `OMS.Business/Services/SuzukiDataImportService.vb`: Stage 2 Order Import service (transform to `ORDERS_STAGE`, reconciliation, and commit).
  - `OMS.Data/Repositories/`: Staging table repositories (`Spirits0501And0502Repository.vb` through `Spirits663NAnd663SAnd66Repository.vb`, `FolderRepository.vb`, `CustomerSettingRepository.vb`).
  - `SUZUKI_SPIRITS_DATA_IMPORT`: Standalone batch console application (`SuzukiSpiritsDataImport.vb`) executing automated pipeline end-to-end.
  - `OMS.Common`: Utilities, file handling (`Utils.vb`), logging, and shared models.
- **Documentation Map:**
  - `docs/project_progress.md`: Gap analysis & 18-step requirements progress matrix between Stage 1 & Stage 2.
  - `docs/SPIRITS_IMPORT_SPEC.md`: Exhaustive 18-step specification for the complete Suzuki SPIRITS import pipeline.
  - `docs/SPIRITS_PIPELINE_SUMMARY.md`: End-to-end flowchart and stage-by-stage pipeline summary.
  - `docs/STAGE2_IMPLEMENTATION_PLAN.md`: Technical architectural plan for Stage 2 (Order Registration, reconciliation, & commit).
  - `docs/BATCH_EXECUTION_SETUP_GUIDE.md`: VM server deployment, folder structure, SQL verification queries, and Task Scheduler guide.
  - `docs/*.pdf`: Official ASTI specification sheets (`SUZUKI_SPIRITS_DATA_PRE_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_MAPPING.pdf`, `受注ファイル取込ページ.pdf`, `内示加工.pdf`, `受注残加工.pdf`, `納入指示加工.pdf`, `STRA納期設定ページ_受注取込後.pdf`).

## Decisions, Fixes & Gotchas
- [2026-09-09] **Error File Suffix Chaining (No Regex Stripping)**: In `SuzukiPreImportService.vb` (`MoveBackErrorFile` & `CopyErrorBackupFile`), removed regex stripping so that existing `_{userId}_{timestamp}` suffixes are retained and chained upon re-importing failed files (e.g. `file_astf01_20260909_092520_astf01_20260909_093045.csv`). This matches the original system (`OrderImport.aspx.vb`) and serves as an audit trail showing repeated import attempts without correction.
- [2026-09-08] **Pre-Import 4-Rule Error Flow**: Implemented exact requirements from Sakai-san:
  - **Rule 1 (0 lines read)**: Adds `「対象データが1件もありません。」` to error list, moves file back to source folder as `{originalName}_{userId}_{timestamp}`, purges from `WORK/`, and rolls back.
  - **Rule 2 (Valid records $\ge$ 1, 0 errors)**: Inserts records to Suzuki staging table, retains file in `WORK/` (`IMP_FILES_STAGE.STATUS = 'DISCOVERED'`), and commits.
  - **Rule 3 (Partial ASTI part error)**: Adds `Row(n): 品目No及び製品コードが取得できません。` for unmapped rows, deletes unmapped rows from staging table via `DeleteErrorRows(conn, tran, infoCode, impFileStageId)`, copies file to source folder with `_{userId}_{timestamp}`, retains file in `WORK/` for Stage 2, and commits valid rows.
  - **Rule 4 (All records have ASTI part error)**: Adds `Row(n): 品目No及び製品コードが取得できません。` for all rows and appends `「対象データが1件もありません。」` at the end, deletes all staged rows and rolls back, moves file back to source folder with `_{userId}_{timestamp}`, and purges file from `WORK/`.
- [2026-09-08] **SPIRITS 6624 Column Index Fix (`FindPartMatchingErrors`)**:
  - In `6624` CSV files (`6624-260715.csv`), `部品番号` (`CUSTOMER_ITEM_NO`) is located at **Column 7** (0-indexed) and Column 9 is `部品番号識別－２` (which holds `"0"`). Previously, `colIdx` was mistakenly set to `9`, comparing `"0"` against unmapped part numbers. This resulted in 0 errors detected, skipping error reporting and `DeleteErrorRows` (leaving `INFO_TYPE_ID = 3908` with `ITEM_NO IS NULL`).
  - Created `SuzukiPreImportService.GetCustomerItemNoColumnIndex(infoCode)` mapping: `6624` $\rightarrow$ 7, `663N/663S/664T` $\rightarrow$ 11, `0813/0814/6604/6634` $\rightarrow$ 9, `0602/0650/0651` $\rightarrow$ 8, `0740` $\rightarrow$ 6, `0501/0502/0600/0630` $\rightarrow$ 10 per `SUZUKI_SPIRITS_DATA_MAPPING.pdf`.
- [2026-09-08] **ASTI Part Matching Z-Suffix & Customer Code 5455**: In `UpdateAstiPartNumbers`, matching on `PRDSLSODRM` applies `UPPER(TRIM(p.FPRDCD)) LIKE '%Z'` for `6604, 6624, 6634, 663N, 663S, 664T`, and `NOT LIKE '%Z'` for others. Joined on `TRIM(p.FCUSTCD) = '5455'`.
- [2026-09-08] **File Encoding & BOM Handling**: Visual Studio and compiler Japanese comment integrity requires UTF-8 with BOM (`0xEF, 0xBB, 0xBF`). Ensured all modified `.vb` files are written with BOM.
- [2026-09-08] **UI Error Display Deduplication**: In `OrderImportStage.aspx.vb`, removed redundant concatenation where `lblError.Text` was assigned twice in the same request. Applied `.Distinct()` formatting so each unique error row appears once.
- [2026-09-08] **CSV Format Error Handling & Recovery**: Fixed issue where files with mangled internal formats were silently skipped. Enhanced `SuzukiCsvParser.PeekInfoTypeCode`, added `ReadLinesAutoEncoding` to safely handle UTF-8 BOM and Shift-JIS.
- [2026-09-07] **Suzuki Batch Master UI Visibility**: In `OrderImportStage.aspx.vb`, hid batch-only master record (Customer `5455` without Profit Center) by default via `OrderImportStage.HideBatchSuzukiCustomer`.
- [2026-09-02] **Staging File Lifecycle**: In Web UI mode, Stage 1 files stay in `C:\ASTI\DATA\WORK\[UserId]\5455\4\` with `IMP_FILES_STAGE.STATUS = 'DISCOVERED'`. Moving to `COMPLETED/` occurs strictly upon Stage 2 Save; discarding moves files back to source `FOLDER_PATH`.
- [2026-09-02] **Web UI PRG Pattern**: Solved F5 duplicate postback by implementing Post/Redirect/Get in `OrderImportStage.aspx.vb` with flash message stored in `Session`. Selection checkboxes reset on redirect.
- [2026-09-02] **Selection Validation**: If no customer checkbox is selected on `OrderImportStage.aspx`, show `「処理対象を選択してください。」` in `lblError` and abort processing.
- [2026-09-02] **Batch Packaging Target**: Added `IncludeBatchExeInPublish` MSBuild target in `OMS.Web.vbproj` and project dependency in `OrderManagementSystem.sln` so `SUZUKI_SPIRITS_DATA_IMPORT.exe` and config automatically copy to `OMS.Web\bin\` during Publish.
- [2026-09-02] **IIS Virtual Path Resolution**: Subfolders under `C:\inetpub\wwwroot\OMS` must be converted to an IIS Application; otherwise `~/Site.Master` fails with HTTP 500 parser error.
- [2026-08-20] **Database Safety**: Never modify the production database directly; test only against localhost/test environment.

## Pending Backlog
- [ ] Await feedback from Sakai-san regarding testing with `6624-260715.csv` and the 4-rule pre-import flow.
- [ ] Implement Stage 2 (`SuzukiDataImportService.vb`): Step 11 (`ORDERS_STAGE` purge) and Steps 12–13 (`ORDERS_STAGE` transform & `ADD_WORKING_DAYS2` calculation).
- [ ] Implement Stage 2 Step 14: Profit Center resolution via `USRDEFFLDF` (`ITEMM`) and special adjustments (`0630`, `663N`, `663S`).
- [ ] Implement Stage 3 Reconciliation: `内示加工` (wash replacement), `受注残加工` (backlog reconciliation), `納入指示加工` (destination conversion & forecast consumption).
- [ ] Implement Stage 4: Commit to `ORDERS` & `ORDER_HISTORY`, update `IMP_RUN`, and trigger `STRA納期設定`.
