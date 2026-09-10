# Project Brain & Memory

## Current Focus
- **Active Goal:** Suzuki (SPIRITS) Order Import Pipeline Implementation (`受注取込_スピリッツ`) across Web UI (`OMS.Web`) and Batch Processing (`SUZUKI_SPIRITS_DATA_IMPORT.exe`).
- **Current Status:** 
  - **Stage 1 (Pre-Import):** Fully implemented in `SuzukiPreImportService.vb` / `SUZUKI_SPIRITS_DATA_PRE_IMPORT`, aligned with Sakai-san's 4-rule error handling flow, 6624 column index fix (col 7), and timestamp suffix chaining. Verified with unit tests.
  - **Stage 2 (Order Registration & Reconciliation):** **100% completed and verified.**
    - `OMS.Data`: Added purge methods, destination conversion (`SUZUKI_SPIRITS_CONVERSION_MST`), wash replacement (`ReplaceNaijiRelationSuzuki` skipping `0602`), forecast consumption (`ReconcileForecastSuzuki`), and staging status updates (`UpdateSuzukiStagingStatusRegistered`).
    - `OMS.Business`: Implemented complete 8-step pipeline in `SuzukiDataImportService.vb` (Profit Center resolution via `USRDEFFLDF`, special code adjustments for `0630`, `663N`, `663S` with fallback to `ORDERS`, validations, commit, and `STRA納期設定`).
    - `OMS.Web`: Wired into `OrderImport.aspx.vb` (`btnImportFile_Click` and `btnSaveOrder_Click`).
    - `SUZUKI_SPIRITS_DATA_IMPORT`: CLI runner connects Pre-Import + Stage 2 registration.
    - `OMS.Tests`: Added `SuzukiDataImportTests.vb` and `SuzukiPreImportTests.vb`; full solution builds with 0 errors.
- **Immediate Next Step:** Await confirmation from Sakai-san on testing `6624-260715.csv` and 4-rule flow on `WS2022STD01`, and perform local end-to-end testing with sample datasets.


## Architecture & Conventions
- **Stack:** Visual Basic .NET (VB.NET), .NET Framework 4.8, ASP.NET Web Forms, Oracle Database (ODP.NET / custom DAL with parameter binding), Windows Server 2022 / IIS 10, Windows Task Scheduler.
- **Core Files / Modules:**
  - `OMS.Web`: Web presentation layer (`Pages/Orders/OrderImportStage.aspx`, `Pages/Orders/OrderImport.aspx`, `Pages/Login/Login.aspx`).
  - `OMS.Business/Services/SuzukiPreImportService.vb`: Stage 1 Pre-Import service handling file intake, WORK folder staging, dedicated table inserts, active flag invalidation, and ASTI part matching with Z-suffix logic.
  - `OMS.Business/Services/SuzukiCsvParser.vb`: CSV parsing logic handling line splitting, encoding auto-detection (UTF-8 with BOM / Shift-JIS), and column sanitization.
  - `OMS.Business/Services/SuzukiDataImportService.vb`: Stage 2 Order Import service (transform to `ORDERS_STAGE`, reconciliation, and commit).
  - `OMS.Data/Repositories/`: Staging table repositories (`Spirits0501And0502Repository.vb` through `Spirits663NAnd663SAnd66Repository.vb`, `FolderRepository.vb`, `CustomerSettingRepository.vb`, `OrderStageRepository.vb`).
  - `SUZUKI_SPIRITS_DATA_IMPORT`: Standalone batch console application (`SuzukiSpiritsDataImport.vb`) executing automated pipeline end-to-end.
  - `OMS.Common`: Utilities, file handling (`Utils.vb`), logging, and shared models.
- **Documentation Map:**
  - `docs/project_progress.md`: Gap analysis & 18-step requirements progress matrix between Stage 1 & Stage 2 (all 18 steps marked DONE).
  - `docs/SPIRITS_IMPORT_SPEC.md`: Exhaustive 18-step specification for the complete Suzuki SPIRITS import pipeline.
  - `docs/SPIRITS_PIPELINE_SUMMARY.md`: End-to-end flowchart and stage-by-stage pipeline summary.
  - `docs/STAGE2_IMPLEMENTATION_PLAN.md`: Technical architectural plan for Stage 2 (Order Registration, reconciliation, & commit).
  - `docs/BATCH_EXECUTION_SETUP_GUIDE.md`: VM server deployment, folder structure, SQL verification queries, and Task Scheduler guide.
  - `docs/STAGE2_TESTING_GUIDE.md`: Stage 2 testing guide, CSV classification (easiest to test, dependencies, pairs), and test execution methods.
  - `docs/*.pdf`: Official ASTI specification sheets (`SUZUKI_SPIRITS_DATA_PRE_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_MAPPING.pdf`, `受注ファイル取込ページ.pdf`, `内示加工.pdf`, `受注残加工.pdf`, `納入指示加工.pdf`, `STRA納期設定ページ_受注取込後.pdf`).

- [2026-09-09] **Stage 1 Parser & Pre-Import Bug Fixes**:
  - **`ParseDecimal` & 0740 `AcceptanceQty`**: Added `ParseDecimal` helper in `SuzukiCsvParser.vb` to safely handle blank values and prevent runtime `InvalidCastException`.
  - **0501/0502 `PublicationTime`**: Updated to use `ParseInt(CleanCol(cols, 7))` preventing `InvalidCastException` when publication time is blank.
  - **`ImpFilesStageRepository` Staging Sort**: Added `ORDER BY imp_file_stage_id DESC` in `GetImpFilesStage` to guarantee `GetImpFilesStageFilename` returns the newest staging row on duplicate file re-imports.
  - **Excel (`.xlsx`) Staging Error Handling**: In `SuzukiPreImportService.vb`, added `webErrors` notification and `MoveBackErrorFile` upon staging failures.
  - **Invalidation Tie-Breaker**: In `UpdateActiveFlags`, added `new_t.info_type_id > old_t.info_type_id` fallback when `created_at` timestamps match within the same clock second.
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
- [2026-09-08] **Step 1 Purge & Profit Center Timing**: In Step 1 (`DELETE ①` / `DeleteOrdersStageForSuzuki`), purge filter uses `CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID` (or `CUSTOMER_CODE = '5455'`) and `FOLDER_TYPE`. Profit Center is NOT filtered at Step 1 because Suzuki raw SPIRITS data is mixed across all Profit Centers (`FOLDER_TYPE = 4`), and Profit Center resolution occurs dynamically in Step 4 via `USRDEFFLDF`. `ORDERS_STAGE` has no `PROFIT_CENTER` column.
- [2026-09-08] **Folder Type Scope**: Batch mode is fixed to `FOLDER_TYPE = 4 (混在)`. Web UI mode defaults to `4` for SPIRITS EDI files (`spprocesstype = 1 AndAlso folderType = 4`), while hand-made files (`chkHandFlag = True`) use `MAPPING_PROFILE_MST` and can support other folder types.
- [2026-09-08] **VB.NET Keyword Conflicts**:
  - Variable named `instr` conflicts with `Microsoft.VisualBasic.Strings.InStr` (`BC30516`). Use `instRow` or `instructionRow`.
  - Loop variable `err` in `For Each err In errors` conflicts with `Microsoft.VisualBasic.Information.Err` (`BC30068`, `BC30456`). Use `errStr`.
- [2026-09-08] **Calender Type Name**: The calendar repository class in `OMS.Data` is named `CalenderRepository` (not `Calender`).
- [2026-09-08] **File Encoding & BOM Handling**: Visual Studio and compiler Japanese comment integrity requires UTF-8 with BOM (`0xEF, 0xBB, 0xBF`). Ensured all modified `.vb` files are written with BOM.
- [2026-09-08] **UI Error Display Deduplication**: In `OrderImportStage.aspx.vb`, removed redundant concatenation where `lblError.Text` was assigned twice in the same request. Applied `.Distinct()` formatting so each unique error row appears once.
- [2026-09-08] **CSV Format Error Handling & Recovery**: Fixed issue where files with mangled internal formats were silently skipped. Enhanced `SuzukiCsvParser.PeekInfoTypeCode`, added `ReadLinesAutoEncoding` to safely handle UTF-8 BOM and Shift-JIS.
- [2026-09-07] **Suzuki Batch Master UI Visibility**: In `OrderImportStage.aspx.vb`, hid batch-only master record (Customer `5455` without Profit Center) by default via `OrderImportStage.HideBatchSuzukiCustomer`.
- [2026-09-02] **Staging File Lifecycle**: In Web UI mode, Stage 1 files stay in `C:\ASTI\DATA\WORK\[UserId]\5455\4\` with `IMP_FILES_STAGE.STATUS = 'DISCOVERED'`. Moving to `COMPLETED/` occurs strictly upon Stage 2 Save ([保存]); discarding ([破棄]) moves files back to source `FOLDER_PATH`.
- [2026-09-02] **Web UI PRG Pattern**: Solved F5 duplicate postback by implementing Post/Redirect/Get in `OrderImportStage.aspx.vb` with flash message stored in `Session`. Selection checkboxes reset on redirect.
- [2026-09-02] **Selection Validation**: If no customer checkbox is selected on `OrderImportStage.aspx`, show `「処理対象を選択してください。」` in `lblError` and abort processing.
- [2026-09-02] **Batch Packaging Target**: Added `IncludeBatchExeInPublish` MSBuild target in `OMS.Web.vbproj` and project dependency in `OrderManagementSystem.sln` so `SUZUKI_SPIRITS_DATA_IMPORT.exe` and config automatically copy to `OMS.Web\bin\` during Publish.
- [2026-09-02] **IIS Virtual Path Resolution**: Subfolders under `C:\inetpub\wwwroot\OMS` must be converted to an IIS Application; otherwise `~/Site.Master` fails with HTTP 500 parser error.
- [2026-08-20] **Database Safety**: Never modify the production database directly; test only against localhost/test environment.

## Pending Backlog
- [ ] Await feedback from Sakai-san regarding testing with `6624-260715.csv` and the 4-rule pre-import flow on `WS2022STD01`.
- [ ] Execute local end-to-end tests (Pre-Import $\rightarrow$ Stage 2 Import $\rightarrow$ Reconciliation $\rightarrow$ Master Orders).
- [ ] Coordinate with Sakai-san for VM server deployment and batch Task Scheduler verification.

