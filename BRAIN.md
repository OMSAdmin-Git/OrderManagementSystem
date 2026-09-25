# Project Brain & Memory

## Current Focus
- **Active Goal:** Suzuki (SPIRITS) Order Import Pipeline Implementation (`受注取込_スピリッツ`) across Web UI (`OMS.Web`), Console Batch (`SUZUKI_SPIRITS_DATA_IMPORT.exe`), and downstream STRA scheduling (`STRA納期設定`).
- **Current Status:**
  - **Stage 1 (Pre-Import):** Completed and verified. In-file `UpdateActiveFlags` bug fixed (multi-schedule lines in the same file all maintain `ACTIVE_FLAG = 'Y'`). 6-digit date (`202609`) monthly forecast handling identified for future spec alignment.
  - **Stage 2 (Order Registration & Reconciliation):** Completed and verified. Fixed `DBNull` cast crash across all 8 loaders via `GetSafeDate`/`GetSafeInt64` in `SuzukiDataImportService.vb`. Updated `ValidateStagedRows` to accept `DEMAND_QTY = 0` per EDI specifications. Successfully imported and committed rows into master `ORDERS` and `ORDER_HISTORY`.
  - **Downstream STRA Delivery Date Setting (`STRA納期設定`):** Resolved batch mode omission where parent Suzuki customer setting (`CUSTOMER_SETTING_ID = 85`, where `PROFIT_CENTER IS NULL`) was skipped. Updated `ExecuteStraDeliveryDateSetting` in `SuzukiDataImportService.vb` to receive `committedSettings` and include all active 5455 settings. Eliminated duplicate Stage 2 call in `SuzukiSpiritsDataImport.vb`.
  - **Order Difference Reporting (`受注差異リスト`):** Investigated empty Excel export. Identified that it is a difference report comparing current vs previous snapshot, plus identified query parameter mismatch (`:customer_setting_id` vs `:p_customer_setting_id`) and trailing semicolon syntax error in `OrderHistoryRepository.vb`.
- **Immediate Next Step:** Run `SUZUKI_SPIRITS_DATA_IMPORT.exe` to verify that `SHIP_PLAN_DATE`, `SHIP_DATE`, and `SHIP_SCHEDULED_DATE` are calculated and updated to `STATUS = 'DUE_SET'` in `ORDERS`.

## Architecture & Conventions
- **Stack:** Visual Basic .NET (VB.NET), .NET Framework 4.8, ASP.NET Web Forms, Oracle Database 19c/21c (ODP.NET Managed Driver with parameter binding), Windows Server 2022 / IIS 10, Windows Task Scheduler.
- **Core Files / Modules:**
  - `OMS.Web`: Web presentation layer (`Pages/Orders/OrderImportStage.aspx` [Stage 1], `Pages/Orders/OrderImport.aspx` [Stage 2], `Pages/Orders/DueDateSettingFromImport.aspx` [STRA納期設定]).
  - `OMS.Business/Services/SuzukiPreImportService.vb`: Stage 1 Pre-Import service (file intake, WORK folder staging, dedicated table inserts, active flag invalidation, ASTI part matching with Z-suffix logic).
  - `OMS.Business/Services/SuzukiCsvParser.vb`: CSV parsing logic with line splitting, auto-encoding detection (UTF-8 BOM / Shift-JIS), and column sanitization.
  - `OMS.Business/Services/SuzukiDataImportService.vb`: Stage 2 Order Import service (transform to `ORDERS_STAGE`, Profit Center resolution, reconciliation, master commit, and batch STRA invocation).
  - `OMS.Data/OrderStageRepository.vb`: Data access for staging, reconciliation, backlog forecast, and `DeliveryDateSetting` lead-time calculation.
  - `OMS.Data/OrderRepository.vb`: Master `ORDERS` CRUD operations and `OrderUpdate` (`MERGE INTO ORDERS`).
  - `OMS.Data/OrderHistoryRepository.vb`: `ORDER_HISTORY` audit tracking and difference extraction queries (`AfterOrderUnofficialNoticeDifference`).
  - `OMS.Data/OorderDiferenceExcelFile.vb`: ClosedXML Excel export for `受注取込差異リスト` and `生産計画差異リスト`.
  - `SUZUKI_SPIRITS_DATA_IMPORT`: Standalone batch console application (`SuzukiSpiritsDataImport.vb`) executing full pipeline end-to-end (`/batch`).
  - `OMS.Common`: Utilities (`Utils.vb`), working calendar calculations, file helpers, and logging.
  - `OMS.Tests`: MSTest suite (`SuzukiPreImportTests.vb`, `SuzukiDataImportTests.vb`).
- **Documentation Map:**
  - `docs/project_progress.md`: 18-step requirements progress matrix between Stage 1 & Stage 2.
  - `docs/SPIRITS_IMPORT_SPEC.md`: Exhaustive 18-step specification for the complete Suzuki SPIRITS import pipeline.
  - `docs/SPIRITS_ORDER_IMPORT_SPEC.md`: Detailed Japanese specification of the order import pipeline and screen/batch differences.
  - `docs/SPIRITS_PIPELINE_SUMMARY.md`: High-level end-to-end flowchart and stage-by-stage pipeline summary.
  - `docs/STAGE1_ACTIVE_FLAG_BUG_FIX_PLAN.md`: Bug analysis and fix documentation for Stage 1 multi-delivery schedule active flag handling.
  - `docs/STAGE2_IMPLEMENTATION_PLAN.md`: Technical architectural design for Stage 2 (registration, reconciliation, and master commit).
  - `docs/STAGE2_TESTING_GUIDE.md`: Guide for CSV test files, data pairings, and test execution procedures.
  - `docs/STAGE2_TEST_PLAN.md` & `docs/STAGE2_TEST_PLAN.html`: Interactive testing matrices and scenario validation plans.
  - `docs/BATCH_EXECUTION_SETUP_GUIDE.md`: Windows Server 2022 deployment, folder layout, Task Scheduler, and SQL checks.
  - `docs/*.pdf`: Official ASTI specification sheets (`SUZUKI_SPIRITS_DATA_PRE_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_IMPORT.pdf`, `SUZUKI_SPIRITS_DATA_MAPPING.pdf`, `受注ファイル取込ページ.pdf`, `内示加工.pdf`, `受注残加工.pdf`, `納入指示加工.pdf`, `STRA納期設定ページ_受注取込後.pdf`).
  - `docs/ORDERSTable/` & `docs/SuzukiTable/`: Table schema reference dumps.

## Decisions, Fixes & Gotchas
- [2026-09-15] **Batch STRA Delivery Date Setting Scope Fix**:
  - `ExecuteStraDeliveryDateSetting` previously filtered `WHERE profit_center IS NOT NULL`, excluding parent customer setting `85` (`profit_center IS NULL`). Updated to pass `committedSettings` directly and query all active `5455` settings so `DeliveryDateSetting` executes for all newly committed orders.
  - In `SuzukiSpiritsDataImport.vb`, eliminated duplicate second call to `ExecuteOrderRegistration` when files have already been processed in `ExecuteImport`.
- [2026-09-11] **Downstream STRA Delivery Date Setting (`STRA納期設定`)**:
  - **Automated Workflow**: In Web UI mode, `DueDateSettingFromImport.aspx` calculates lead times (輸送L/T from `SHPROUTM` + 品揃L/T from `USRDEFFLDF`), shifts dates via ASTI work calendar `00001`, updates `ORDERS` (`SHIP_DATE`, `SHIP_SCHEDULED_DATE`, `SHIP_PLAN_DATE`, `STATUS = 'DUE_SET'`), and automatically downloads the difference Excel report. In Batch mode (`/batch`), Step 8 calls `DeliveryDateSetting` automatically without user intervention.
  - **Difference Report Query Bug**: In `OrderHistoryRepository.vb`, `AfterOrderUnofficialNoticeDifference` and `AfterOrderDeliveryInstructionDifference` use `:customer_setting_id` in SQL but bind `:p_customer_setting_id` with `BindByName = True`. Additionally, line 644 contains a trailing `;` inside the SQL string. Both errors are caught and swallowed, resulting in an empty difference dataset.
- [2026-09-11] **Stage 2 Validation & DBNull Defensive Handling**:
  - **`DEMAND_QTY = 0` Allowed**: Changed `DemandQty.Value <= 0` to `< 0` in `SuzukiDataImportService.ValidateStagedRows`. Forecast EDI files routinely contain 0-quantity schedule lines representing zero demand or cancellations.
  - **Null-Safe Reading**: Added `GetSafeDate` and `GetSafeInt64` across all 8 SPIRITS loaders and 3 adjustment routines to prevent unhandled `InvalidCastException` crashes when nullable date/qty columns are read from database readers.
- [2026-09-10] **Stage 1 Multi-Schedule Line Active Flag Invalidation Fix**:
  - `UpdateActiveFlags` previously set `ACTIVE_FLAG = 'N'` for all rows except the last row in the same file. Fixed SQL/logic so in-file rows for identical part/order numbers across different delivery dates retain `ACTIVE_FLAG = 'Y'`. Only records from *older* import runs are invalidated.
- [2026-09-09] **Stage 1 Parser & Pre-Import Bug Fixes**:
  - Added `ParseDecimal` helper in `SuzukiCsvParser.vb` to prevent `InvalidCastException` on blank values.
  - Used `ParseInt(CleanCol(cols, 7))` for 0501/0502 publication time.
  - Added `ORDER BY imp_file_stage_id DESC` in `GetImpFilesStage` to return the newest staging row on duplicate file re-imports.
  - Invalidation tie-breaker added (`new_t.info_type_id > old_t.info_type_id`) for matching timestamp seconds.
- [2026-09-09] **Error File Suffix Chaining**: Removed regex stripping from `MoveBackErrorFile` so repeated import attempts chain suffixes (e.g. `file_astf01_20260909_092520_astf01_20260909_093045.csv`) for audit traceability.
- [2026-09-08] **Pre-Import 4-Rule Error Flow**: Implemented exact requirements from Sakai-san (Rule 1: 0 lines read $\rightarrow$ error; Rule 2: valid $\ge$ 1, 0 errors $\rightarrow$ staged; Rule 3: partial ASTI part error $\rightarrow$ strip error rows, stage valid; Rule 4: all part errors $\rightarrow$ reject all).
- [2026-09-08] **SPIRITS 6624 Column Index**: `CUSTOMER_ITEM_NO` for `6624` is at Column 7 (0-indexed). Centralized lookup via `SuzukiPreImportService.GetCustomerItemNoColumnIndex`.
- [2026-09-08] **ASTI Part Matching Z-Suffix**: Joined on `TRIM(p.FCUSTCD) = '5455'`, matching `UPPER(TRIM(p.FPRDCD)) LIKE '%Z'` for `6604, 6624, 6634, 663N, 663S, 664T`, and `NOT LIKE '%Z'` for other info codes.
- [2026-09-08] **File Encoding & BOM Handling**: Visual Studio compiler and Japanese comments require UTF-8 with BOM (`0xEF, 0xBB, 0xBF`). Must restore BOM whenever files are edited.
- [2026-08-20] **Database Safety**: Never modify the production database directly; test only against localhost/test environment.

## Pending Backlog
- [ ] Fix query parameter name mismatch (`:customer_setting_id` vs `:p_customer_setting_id`) and trailing semicolon in `OrderHistoryRepository.vb` for `受注差異リスト` reporting.
- [ ] Clarify business requirement on 6-digit forecast delivery dates (`202609` / `yyyyMM`) for Stage 1 parser.
- [ ] Await feedback from Sakai-san on testing `6624-260715.csv` and the 4-rule pre-import flow on VM server `WS2022STD01`.
