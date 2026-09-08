# Project Brain & Memory

## Current Focus
- **Active Goal:** Suzuki (SPIRITS) Order Import Pipeline Implementation (`受注取込_スピリッツ`) across Web UI (`OMS.Web`) and Batch Processing (`SUZUKI_SPIRITS_DATA_IMPORT.exe`).
- **Current Status:** Stage 1 Pre-Import (`SuzukiPreImportService.vb` / `SUZUKI_SPIRITS_DATA_PRE_IMPORT`) is fully implemented and tested. Files remain in `WORK/` (`C:\ASTI\DATA\WORK\[UserId]\5455\4\`) with status `DISCOVERED` awaiting Stage 2 Save or Discard. Web UI PRG (Post/Redirect/Get) implemented to prevent duplicate submissions on F5; customer selection validation added; batch packaging target added to `OMS.Web.vbproj` to copy `SUZUKI_SPIRITS_DATA_IMPORT.exe` to `bin\`; IIS deployment instructions prepared for Sakai-san.
- **Immediate Next Step:** Verify test execution of `SUZUKI_SPIRITS_DATA_IMPORT.exe` and IIS Web UI with Sakai-san, then begin Stage 2 (`SuzukiDataImportService.vb` / `ORDERS_STAGE` transform & reconciliation).

## Architecture & Conventions
- **Stack:** Visual Basic .NET (VB.NET), .NET Framework 4.8, ASP.NET Web Forms, Oracle Database (ODP.NET / custom DAL with parameter binding), Windows Server 2022 / IIS 10, Windows Task Scheduler.
- **Core Files / Modules:**
  - `OMS.Web`: Web presentation layer (`Pages/Orders/OrderImportStage.aspx`, `Pages/Orders/OrderImport.aspx`, `Pages/Login/Login.aspx`).
  - `OMS.Business/Services/SuzukiPreImportService.vb`: Stage 1 Pre-Import service handling file intake, WORK folder staging, dedicated table inserts, active flag invalidation, and ASTI part matching with Z-suffix logic.
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
- [2026-09-08] **UI Error Display Deduplication**: In `OrderImportStage.aspx.vb`, removed redundant concatenation where `lblError.Text` was assigned twice in the same request (lines 452 & 536). Applied `.Distinct()` formatting to guarantee each unique error row appears exactly once on screen.
- [2026-09-08] **CSV Format Error Handling & Recovery**: Fixed issue where files with mangled internal formats (e.g. missing leading zeros `"0600"` -> `600` due to Excel editing) were silently skipped. Enhanced `SuzukiCsvParser.PeekInfoTypeCode` with diagnostics, added `ReadLinesAutoEncoding` to safely handle UTF-8 BOM and Shift-JIS. In `SuzukiPreImportService.vb`, format errors now increment `errorFilesCount`, output `ErrorList_*.csv`, display clear errors on `lblError`, and return the file to the intake folder while deleting it from `WORK/` to eliminate stranded files. Added transaction rollback on any file error. Added loading overlay and double-click prevention on `OrderImportStage.aspx`.
- [2026-09-07] **Suzuki Batch Master UI Visibility**: In `OrderImportStage.aspx.vb`, hid the batch-only master record (Customer `5455` with no Profit Center) by default from the customer selection grid (`gvSelectCustomers`). Added a hybrid toggle via `ConfigurationManager.AppSettings("OrderImportStage.HideBatchSuzukiCustomer")`, allowing instant re-enabling if business specifications change, without recompilation.
- [2026-09-02] **Staging File Lifecycle**: In Web UI mode, Stage 1 files stay in `C:\ASTI\DATA\WORK\[UserId]\5455\4\` with `IMP_FILES_STAGE.STATUS = 'DISCOVERED'`. Moving to `COMPLETED/` occurs strictly upon Stage 2 Save ([保存]); discarding ([破棄]) moves files back to source `FOLDER_PATH`.
- [2026-09-02] **Web UI PRG Pattern**: Solved F5 duplicate postback by implementing Post/Redirect/Get in `OrderImportStage.aspx.vb` with flash message stored in `Session`. Selection checkboxes reset on redirect.
- [2026-09-02] **Selection Validation**: If no customer checkbox is selected on `OrderImportStage.aspx`, show `「処理対象を選択してください。」` in `lblError` and abort processing.
- [2026-09-02] **Batch Packaging Target**: Added `IncludeBatchExeInPublish` MSBuild target in `OMS.Web.vbproj` and project dependency in `OrderManagementSystem.sln` so `SUZUKI_SPIRITS_DATA_IMPORT.exe` and config automatically copy to `OMS.Web\bin\` during Publish.
- [2026-09-02] **IIS Virtual Path Resolution**: Subfolders under `C:\inetpub\wwwroot\OMS` must be converted to an IIS Application; otherwise `~/Site.Master` fails with HTTP 500 parser error.
- [2026-09-01] **Part Matching Z-Suffix Rule**: Part numbers ending with `'Z'` apply strictly to EDI codes `6604`, `6624`, `6634`, `663N`, `663S`, `664T` in `PRDSLSODRM`; other codes match without 'Z'. Unmatched errors generate Excel error reports in `[FOLDER_PATH]\エラーリスト\`.
- [2026-08-20] **Database Safety**: Never modify the production database directly; test only against localhost/test environment.

## Pending Backlog
- [ ] Assist user with Sakai-san's test execution of `SUZUKI_SPIRITS_DATA_IMPORT.exe` and IIS Web UI.
- [ ] Implement Stage 2 (`SuzukiDataImportService.vb`): Step 11 (`ORDERS_STAGE` purge) and Steps 12–13 (`ORDERS_STAGE` transform & `ADD_WORKING_DAYS2` calculation).
- [ ] Implement Stage 2 Step 14: Profit Center resolution via `USRDEFFLDF` (`ITEMM`) and special adjustments (`0630`, `663N`, `663S`).
- [ ] Implement Stage 3 Reconciliation: `内示加工` (wash replacement), `受注残加工` (backlog reconciliation), `納入指示加工` (destination conversion & forecast consumption).
- [ ] Implement Stage 4: Commit to `ORDERS` & `ORDER_HISTORY`, update `IMP_RUN`, and trigger `STRA納期設定`.
