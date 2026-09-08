# Project Brain & Memory

## Current Focus
- **Active Goal:** Suzuki (SPIRITS) Order Import Pipeline Implementation (`受注取込_スピリッツ`) across Web UI (`OMS.Web`) and Batch Processing (`SUZUKI_SPIRITS_DATA_IMPORT.exe`).
- **Current Status:** 
  - **Stage 1 (Pre-Import):** Fully implemented in `SuzukiPreImportService.vb` / `SUZUKI_SPIRITS_DATA_PRE_IMPORT`. (Note: user will adjust Stage 1 UI on the pre-import branch later).
  - **Stage 2 (Order Registration & Reconciliation):** **100% completed and verified.**
    - `OMS.Data`: Added purge methods, destination conversion (`SUZUKI_SPIRITS_CONVERSION_MST`), wash replacement (`ReplaceNaijiRelationSuzuki` skipping `0602`), forecast consumption (`ReconcileForecastSuzuki`), and staging status updates (`UpdateSuzukiStagingStatusRegistered`).
    - `OMS.Business`: Implemented complete 8-step pipeline in `SuzukiDataImportService.vb` (Profit Center resolution via `USRDEFFLDF`, special code adjustments for `0630`, `663N`, `663S` with fallback to `ORDERS`, validations, commit, and `STRA納期設定`).
    - `OMS.Web`: Wired into `OrderImport.aspx.vb` (`btnImportFile_Click` and `btnSaveOrder_Click`).
    - `SUZUKI_SPIRITS_DATA_IMPORT`: CLI runner connects Pre-Import + Stage 2 registration.
    - `OMS.Tests`: Added `SuzukiDataImportTests.vb`; all 5 tests passed. Full solution builds with 0 errors.
- **Immediate Next Step:** Switch to pre-import branch when ready to refine Stage 1 page/flow, or perform local end-to-end testing of Stage 2 with test data.

## Architecture & Conventions
- **Stack:** Visual Basic .NET (VB.NET), .NET Framework 4.8, ASP.NET Web Forms, Oracle Database (ODP.NET / custom DAL with parameter binding), Windows Server 2022 / IIS 10, Windows Task Scheduler.
- **Core Files / Modules:**
  - `OMS.Web`: Web presentation layer (`Pages/Orders/OrderImportStage.aspx`, `Pages/Orders/OrderImport.aspx`, `Pages/Login/Login.aspx`).
  - `OMS.Business/Services/SuzukiPreImportService.vb`: Stage 1 Pre-Import service handling file intake, WORK folder staging, dedicated table inserts, active flag invalidation, and ASTI part matching with Z-suffix logic.
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

## Decisions, Fixes & Gotchas
- [2026-09-08] **Stage 1 UI Separation**: On branch `feature/suzuki-import`, `OrderImportStage.aspx` and `OrderImportStage.aspx.vb` are left untouched as requested by user, to avoid branch merge conflicts. Any Stage 1 UI refinements will be done on the Stage 1 branch.
- [2026-09-08] **Step 1 Purge & Profit Center Timing**: In Step 1 (`DELETE ①` / `DeleteOrdersStageForSuzuki`), purge filter uses `CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID` (or `CUSTOMER_CODE = '5455'`) and `FOLDER_TYPE`. Profit Center is NOT filtered at Step 1 because Suzuki raw SPIRITS data is mixed across all Profit Centers (`FOLDER_TYPE = 4`), and Profit Center resolution occurs dynamically in Step 4 via `USRDEFFLDF`. `ORDERS_STAGE` has no `PROFIT_CENTER` column.
- [2026-09-08] **Folder Type Scope**: Batch mode is fixed to `FOLDER_TYPE = 4 (混在)`. Web UI mode defaults to `4` for SPIRITS EDI files (`spprocesstype = 1 AndAlso folderType = 4`), while hand-made files (`chkHandFlag = True`) use `MAPPING_PROFILE_MST` and can support other folder types.
- [2026-09-08] **VB.NET Keyword Conflicts**:
  - Variable named `instr` conflicts with `Microsoft.VisualBasic.Strings.InStr` (`BC30516`). Use `instRow` or `instructionRow`.
  - Loop variable `err` in `For Each err In errors` conflicts with `Microsoft.VisualBasic.Information.Err` (`BC30068`, `BC30456`). Use `errStr`.
- [2026-09-08] **Calender Type Name**: The calendar repository class in `OMS.Data` is named `CalenderRepository` (not `Calender`).
- [2026-09-08] **UTF-8 with BOM Requirement**: All `.vb` source files must strictly preserve UTF-8 BOM (`EF BB BF`). Use BOM verification script (`scratch/check_bom.ps1`) after file modifications.
- [2026-09-07] **Suzuki Batch Master UI Visibility**: In `OrderImportStage.aspx.vb`, hid the batch-only master record (Customer `5455` with no Profit Center) by default from the customer selection grid (`gvSelectCustomers`). Added a hybrid toggle via `ConfigurationManager.AppSettings("OrderImportStage.HideBatchSuzukiCustomer")`, allowing instant re-enabling if business specifications change, without recompilation.
- [2026-09-02] **Staging File Lifecycle**: In Web UI mode, Stage 1 files stay in `C:\ASTI\DATA\WORK\[UserId]\5455\4\` with `IMP_FILES_STAGE.STATUS = 'DISCOVERED'`. Moving to `COMPLETED/` occurs strictly upon Stage 2 Save ([保存]); discarding ([破棄]) moves files back to source `FOLDER_PATH`.
- [2026-09-02] **Web UI PRG Pattern**: Solved F5 duplicate postback by implementing Post/Redirect/Get in `OrderImportStage.aspx.vb` with flash message stored in `Session`. Selection checkboxes reset on redirect.
- [2026-09-02] **Selection Validation**: If no customer checkbox is selected on `OrderImportStage.aspx`, show `「処理対象を選択してください。」` in `lblError` and abort processing.
- [2026-09-02] **Batch Packaging Target**: Added `IncludeBatchExeInPublish` MSBuild target in `OMS.Web.vbproj` and project dependency in `OrderManagementSystem.sln` so `SUZUKI_SPIRITS_DATA_IMPORT.exe` and config automatically copy to `OMS.Web\bin\` during Publish.
- [2026-09-02] **IIS Virtual Path Resolution**: Subfolders under `C:\inetpub\wwwroot\OMS` must be converted to an IIS Application; otherwise `~/Site.Master` fails with HTTP 500 parser error.
- [2026-09-01] **Part Matching Z-Suffix Rule**: Part numbers ending with `'Z'` apply strictly to EDI codes `6604`, `6624`, `6634`, `663N`, `663S`, `664T` in `PRDSLSODRM`; other codes match without 'Z'. Unmatched errors generate Excel error reports in `[FOLDER_PATH]\エラーリスト\`.
- [2026-08-20] **Database Safety**: Never modify the production database directly; test only against localhost/test environment.

## Pending Backlog
- [ ] Switch to pre-import branch when ready to adjust Stage 1 logic/UI on `OrderImportStage.aspx` / `OrderImportStage.aspx.vb`.
- [ ] Execute local end-to-end tests (Pre-Import $\rightarrow$ Stage 2 Import $\rightarrow$ Reconciliation $\rightarrow$ Master Orders).
- [ ] Coordinate with Sakai-san for VM server deployment and batch Task Scheduler verification.
