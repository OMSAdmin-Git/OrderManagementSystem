# 🚀 Implementation Plan: Stage 2 (Order Registration & Reconciliation / SUZUKI_SPIRITS_DATA_IMPORT)

## 📋 Overview & Scope

In Stage 1 (`SUZUKI_SPIRITS_DATA_PRE_IMPORT`), incoming EDI/text files from Suzuki SPIRITS were validated, parsed, part-matched against `PRDSLSODRM`, and registered into dedicated Suzuki staging tables (`SUZUKI_SPIRITS_*`).

**Stage 2 (`SUZUKI_SPIRITS_DATA_IMPORT`)** transforms these staged Suzuki records into ASTI's standard order staging table (`ORDERS_STAGE`), applies code-specific adjustments (`0630`, `663N`, `663S`), resolves Profit Centers (`USRDEFFLDF`), performs stage error checks, executes destination conversions (`SUZUKI_SPIRITS_CONVERSION_MST`), applies full reconciliation processing (`内示加工`, `受注残加工`, `納入指示加工`), commits valid rows into master production tables (`ORDERS` and `ORDER_HISTORY`), and executes downstream delivery scheduling (`STRA納期設定`).

---

## 🎛️ Dual Execution Architecture & Lifecycle

```mermaid
flowchart TD
    Start([Start Stage 2]) --> ModeCheck{Execution Mode?}
    
    %% Batch Mode Flow
    ModeCheck -- Batch / CLI<br/>SUZUKI_SPIRITS_DATA_IMPORT.exe /batch --> B1[Step 1: Purge ORDERS_STAGE]
    B1 --> B2[Step 2-3: Filter & Transform to ORDERS_STAGE]
    B2 --> B3[Step 4: Profit Center Resolution & Special Adjustments<br/>0630, 663N, 663S]
    B3 --> B4[Step 5: Post-Stage Validation & Error Check]
    B4 -- Error --> B_Err[Generate CSV Error Report & Rollback]
    B4 -- Valid --> B5[Step 6: Destination Conversion & Forecast Reconciliation<br/>内示加工 / 受注残加工 / 納入指示加工]
    B5 --> B6[Step 7: Commit to ORDERS & ORDER_HISTORY<br/>Set Suzuki Staging STATUS='REGISTERED']
    B6 --> B7[Step 8: Execute STRA納期設定 & Move files to COMPLETED<br/>Update IMP_RUN to COMPLETED]
    B7 --> B_End([Batch Completed])

    %% UI Mode Flow
    ModeCheck -- Web UI: 取込実行<br/>OrderImport.aspx --> U1[Step 1: Purge ORDERS_STAGE]
    U1 --> U2[Check ASTI追加内示 HAND_FLAG]
    U2 -- Hand='Y' --> U_Hand[Parse manual file via MAPPING_PROFILE_MST]
    U2 -- Hand='N' --> U3[Step 2-3: Transform Suzuki Staging to ORDERS_STAGE]
    U_Hand --> U4[Step 4: Special Adjustments]
    U3 --> U4
    U4 --> U5[Step 5: Validate & Render errors to lblImportError]
    U5 --> U6[Step 6: Destination Conversion & Forecast Reconciliation]
    U6 --> U_Review([Display in gvImportOrder for User Review])
    
    U_Review -- User clicks 保存 (Save) --> U_Save[Step 7: Commit to ORDERS & ORDER_HISTORY<br/>Move WORK files to COMPLETED<br/>Update IMP_FILES / IMP_RUN]
    U_Review -- User clicks 破棄 (Discard) --> U_Discard[Rollback ORDERS_STAGE & Move WORK files back to IN]
    U_Save --> U_End([UI Save Completed])
    U_Discard --> U_End2([UI Discard Completed])
```

---

## 🔄 Detailed 8-Step Technical Implementation Plan

### Step 1: Stage Purge (`DELETE ① / 受注ワーク削除`)
* Purges any existing uncommitted records in `ORDERS_STAGE` where `CUSTOMER_SETTING_ID = Target` and `FOLDER_TYPE = 4`.

### Step 2 & 3: Order Data Filtering & `ORDERS_STAGE` Insert (`受注データチェック` / `INSERT ORDERS_STAGE`)
* Target Suzuki tables (active records `ACTIVE_FLAG = 'Y'`):
  * `SUZUKI_SPIRITS_0600AND0630` (`0600`) - 確定受注（月次内示）
  * `SUZUKI_SPIRITS_0602` (`0602`) - 確定受注（追加内示）
  * `SUZUKI_SPIRITS_0650` (`0650`) - 納入指示（通常）
  * `SUZUKI_SPIRITS_0651` (`0651`) - 納入指示（臨時）
  * `SUZUKI_SPIRITS_0813` (`0813`) - 支給品情報データ
  * `SUZUKI_SPIRITS_0814` (`0814`) - 支給品納入指示データ
  * `SUZUKI_SPIRITS_6604AND6634` (`6604`) - 注文データ（内示）
  * `SUZUKI_SPIRITS_6624` (`6624`) - 注文データ（確定/指示）
  *(Note: `0630`, `663N`, `663S` are excluded from initial insert; handled in Step 4 as updates)*.
* Transforms and inserts into `ORDERS_STAGE`:
  * `COMPANY_ID`: Fixed `'1000'`, `BILLING_STANDARD`: Fixed `'S'`, `TRANSPORT_METHOD`: Fixed `2`
  * `STATUS`: Fixed `'IMPORTED'`, `ACTIVE_FLAG`: Fixed `'Y'`
  * `ORDER_TYPE`:
    * `1` (内示 / Forecast) for `0600`, `0602`, `0813`, `6604`
    * `3` (納入指示 / Delivery Instruction) for `0650`, `0651`, `0814`, `6624`
  * `DEMAND_STATUS`: `'F'` (for `ORDER_TYPE = 1`), `'O'` (for `ORDER_TYPE = 3`)
  * `SHIP_PROCESS_TYPE`: `'O'` (for `ORDER_TYPE = 1`), `'K'` (for `ORDER_TYPE = 3`)
  * `DELIVERY_INSTR_FLAG`: `'N'` (for `ORDER_TYPE = 1`), `'Y'` (for `ORDER_TYPE = 3`)
  * `TOTAL_SHIP_QTY`: `NULL` (for `ORDER_TYPE = 1`), `0` (for `ORDER_TYPE = 3`)
  * `DUE_DATE`: Oracle working calendar `ADD_WORKING_DAYS2('00001', [DELIVERY_DATE], 0)` formatted as `yyyy/MM/dd`
  * `PRE_DAILY_DELIVERY_DATE`: Original unadjusted `DELIVERY_DATE`
  * `DEMAND_QTY` & `PRE_DAILY_ORDER_QTY`: Staging `ORDER_QTY`
  * `DEMAND_UNIT`: Joined from `ITEMM.FUNIT` where `ITEMM.FITEMNO = PRDSLSODRM.FPRDCD`
  * `SHIP_STOCK_LOCATION`: Joined from `ITEMM.FSHPWHCD` where `ITEMM.FITEMNO = ORDERS_STAGE.ITEM_NO`
  * `BILLING_TO`: Looked up from `SECTM.FBILLTOCD` via `GetBillingTo(CUSTOMER_CODE)`
  * `CURRENCY_CODE`: Looked up from `SECTM.FCURR` where `SECTM.FSECTCD = CUSTOMER_CODE`
  * `PRORATED_TYPE`:
    * If `DELIVERY_DATE_TYPE` in (`'1'`, `'4'`) $\rightarrow$ `2` (日割以外 / Non-prorated)
    * If `DELIVERY_DATE_TYPE` in (`'5'`, `'6'`) $\rightarrow$ `1` (日割り / Daily prorated)
  * Extended Phase 2 Columns:
    * `DELIVERY_TIME`, `CONTAINER_CAPACITY`, `INITIAL_FLAG`, `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE`.

### Step 4: Special Demand & Delivery Adjustments (`UPDATE ① / 受注ワーク更新`)
1. **Profit Center Resolution**:
   * Retrieves `FUSRSTR1` from `USRDEFFLDF` where `FTABLEID='ITEMM'` and `FRECKEY = PRODUCT_CODE`.
   * Queries `CUSTOMER_SETTING_MST` where `CUSTOMER_CODE = '5455'` and `PROFIT_CENTER = [FUSRSTR1]` and `ACTIVE_FLAG = 'Y'`.
   * If found, updates `ORDERS_STAGE.CUSTOMER_SETTING_ID`.
2. **Code `0630` (増減 / Increment-Decrement)**:
   * Matches `ORDERS_STAGE` on `ITEM_NO`, `PRE_DAILY_DELIVERY_DATE`, `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE = '0600'`.
   * Updates `DEMAND_QTY = DEMAND_QTY + 0630.ORDER_QTY` and `PRE_DAILY_ORDER_QTY`.
   * Fallback: If not found in `ORDERS_STAGE`, fetches active `0600` from master `ORDERS`, copies to `ORDERS_STAGE`, and adds quantity.
3. **Code `663N` (納期変更 / Due Date Change)**:
   * Matches `ORDERS_STAGE` on `CUSTOMER_ORDER_NO` and `INFO_TYPE_CODE = '6624'`.
   * Updates `DUE_DATE = ADD_WORKING_DAYS2('00001', 663N.DELIVERY_DATE, 0)`.
   * Fallback: If not found in `ORDERS_STAGE`, fetches active `6624` from master `ORDERS`, copies to `ORDERS_STAGE`, and updates `DUE_DATE`.
4. **Code `663S` (数量変更 / Quantity Change)**:
   * Matches `ORDERS_STAGE` on `CUSTOMER_ORDER_NO` and `INFO_TYPE_CODE = '6624'`.
   * Updates `DEMAND_QTY = 663S.ORDER_QTY` and `PRE_DAILY_ORDER_QTY`.
   * Fallback: If not found in `ORDERS_STAGE`, fetches active `6624` from master `ORDERS`, copies to `ORDERS_STAGE`, and updates `DEMAND_QTY`.

### Step 5: Post-Stage Validation & Error Handling (`エラーチェック`)
* Validates staged data in `ORDERS_STAGE`:
  * Flags error if `CUSTOMER_SETTING_ID` has generic customer `5455` with blank `PROFIT_CENTER` / `CUSTOMER_UNIT_ID` (unresolved Profit Center).
  * Validates standard required fields (`ITEM_NO`, `DUE_DATE`, `DEMAND_QTY > 0`, `SHIP_TO`).
* If errors exist:
  * Batch Mode: Generates CSV error report in `[FOLDER_PATH]\エラーリスト\`.
  * UI Mode: Renders error messages in `lblImportError`.
  * Rollback / physical `DELETE` of uncommitted `ORDERS_STAGE` rows.
  * Backs up error file as `[OriginalFileName]_[UserID]_[Timestamp]`.

### Step 6: Reconciliation & Destination Processing (`UPDATE ②`)

#### 1. 内示加工 (Forecast Wash Replacement per `内示加工.pdf`)
* Counts customer naiji ($x$: `SELF_FCST_FLAG = 'N'`) and ASTI manual naiji ($y$: `SELF_FCST_FLAG = 'Y'`):
  - **Case ① ($x \ge 1, y = 0$)**: Invalidate prior customer forecasts (`SELF_FCST_FLAG = 'N'`) and auto-delete ASTI naiji (`SELF_FCST_DELETE_FLAG = 'Y'`).
  - **Case ② ($x = 0, y \ge 1$)**: Invalidate prior ASTI manual forecasts (`SELF_FCST_FLAG = 'Y'`).
  - **Case ③ ($x \ge 1, y \ge 1$)**: Invalidate all prior forecasts.
* **Special Rule `0602`**: Has no reference date fields $\rightarrow$ treated as pure new data (skips wash replacement).
* Sets newly processed naiji status to `STATUS = 'PROCESSED'`.

#### 2. 受注残加工 (Backlog Reconciliation per `受注残加工.pdf`)
* When `RECONCILE_FLAG = 'Y'`:
  - Queries `STRA_ORDER_BACKLOG` and reconciles demand quantities against orders based on `RECONCILE_TYPE` (`1: 順次`, `2: 同月まで`, `3: 同月のみ`).
  - When remaining `DEMAND_QTY = 0`: sets `STATUS = 'RECONCILED'` and `ACTIVE_FLAG = 'N'`.

#### 3. 納入指示加工 (Delivery Instruction Processing per `納入指示加工.pdf`)
1. **打切 (Termination)**: Records where `ORDER_TYPE = 3` and `INFO_TYPE = 'N'` $\rightarrow$ `STATUS = 'PROCESSED'`, `ACTIVE_FLAG = 'N'`.
2. **取消 (Cancellation)**: Records where `ORDER_TYPE = 3` and `INFO_TYPE = 'D'` $\rightarrow$ `DEMAND_QTY = 0`, `STATUS = 'PROCESSED'`.
3. **出荷先変換 (`SUZUKI_SPIRITS_CONVERSION_MST`)**:
   - Matches `DELIVERY_CODE_ORDER = SHIP_TO` and replaces `SHIP_TO = DELIVERY_CODE_PLAN`.
4. **内示消込 (Forecast Consumption)**:
   - When `FCST_RECONCILE_FLAG = 'Y'`:
     - `0650` / `0651` consumes active `0600` forecasts.
     - `0814` consumes active `0813` forecasts.
     - `6624` consumes active `6604` forecasts.
5. **新規納入指示確定**: Marks remaining new delivery instructions as `STATUS = 'PROCESSED'`.

### Step 7: Master Production Table Commitment (`INSERT ORDERS` / `INSERT ORDER_HISTORY`)
* Commits valid rows from `ORDERS_STAGE` into master production table `ORDERS` (including extended attributes).
* Appends historical audit records into `ORDER_HISTORY`.
* Updates Suzuki staging table records (`STATUS = 'REGISTERED'`).
* Inserts records into `IMP_FILES` (`STATUS = 'PARSED'`) and purges processed entries from `IMP_FILES_STAGE`.
* Moves processed files from `WORK/` to `COMPLETED/[UserId]/5455/4/` as `[FileName]_[User]_[Timestamp].[ext]`.

### Step 8: Completion Logging & STRA Integration (`UPDATE IMP_RUN` / `STRA納期設定`)
* Updates `IMP_RUN`: `ENDED_AT = SYSDATE`, `STATUS = 'COMPLETED'`, `FILE_COUNT`, `ROW_COUNT`, `ERROR_COUNT`.
* **Batch Mode Only**: Automatically triggers downstream `OrderStageRepository.DeliveryDateSetting` (`STRA納期設定`) for customer `5455`:
  - Kit Lead Time (`USRDEFFLDF.FUSRDEC1`)
  - Transport Lead Time (`SHPROUTM.FTRANLT`)
  - Calculates `SHIP_SCHEDULED_DATE`, `SHIP_DATE`, `SHIP_PLAN_DATE` via `ADD_WORKING_DAYS`.
  - Sets `ORDERS.STATUS = 'DUE_SET'` and logs to `ORDER_HISTORY`.

---

## 📂 Proposed File Changes

### 1. `OMS.Business`
#### [MODIFY] [SuzukiDataImportService.vb](file:///c:/Dev/Projects/ASTI/OrderManagementSystem-dev/OMS.Business/Services/SuzukiDataImportService.vb)
* Implement full Stage 2 workflow:
  * `ExecuteOrderRegistration`: Batch Mode entry point (orchestrates Steps 1–8).
  * `ExecuteStageImportToOrdersStage`: Web UI Mode entry point (orchestrates Steps 1–6).
  * `ExecuteSaveOrders`: Commits `ORDERS_STAGE` to `ORDERS` and `ORDER_HISTORY` (Step 7).
  * `ExecuteDiscardOrders`: Discards uncommitted stage records and moves files back to `IN`.
  * Helper methods for Profit Center resolution, code adjustments (`0630`, `663N`, `663S`), ship-to conversion, forecast reconciliation, and forecast consumption.

### 2. `OMS.Data`
#### [MODIFY] [Spirits repositories under `OMS.Data\SUZUKI\`](file:///c:/Dev/Projects/ASTI/OrderManagementSystem-dev/OMS.Data/SUZUKI/)
* Add query and status update methods (`GetActiveRows`, `UpdateStatusRegistered`) across all Suzuki repositories:
  * `Spirits0600And0630Repository.vb`
  * `Spirits0602Repository.vb`
  * `Spirits0650Repository.vb`
  * `Spirits0651Repository.vb`
  * `Spirits0813Repository.vb`
  * `Spirits0814Repository.vb`
  * `Spirits6604And6634Repository.vb`
  * `Spirits6624Repository.vb`
  * `Spirits663NAnd663SAnd66Repository.vb`

#### [MODIFY] [OrderStageRepository.vb](file:///c:/Dev/Projects/ASTI/OrderManagementSystem-dev/OMS.Data/OrderStageRepository.vb)
* Add support for Suzuki-specific consumption pairs (`0650/0651 -> 0600`, `0814 -> 0813`, `6624 -> 6604`) and destination conversion.

### 3. `OMS.Web`
#### [MODIFY] [OrderImport.aspx.vb](file:///c:/Dev/Projects/ASTI/OrderManagementSystem-dev/OMS.Web/Pages/Orders/OrderImport.aspx.vb)
* Hook `SuzukiDataImportService.ExecuteStageImportToOrdersStage` in `btnImportFile_Click` when `SpProcessType = 1`.
* Hook `SuzukiDataImportService.ExecuteSaveOrders` in `btnSaveOrder_Click`.
* Hook `SuzukiDataImportService.ExecuteDiscardOrders` in `btnCancelOrder_Click`.

### 4. `SUZUKI_SPIRITS_DATA_IMPORT`
#### [MODIFY] [SuzukiSpiritsDataImport.vb](file:///c:/Dev/Projects/ASTI/OrderManagementSystem-dev/SUZUKI_SPIRITS_DATA_IMPORT/SuzukiSpiritsDataImport.vb)
* Connect CLI runner to `SuzukiDataImportService.ExecuteOrderRegistration` following Stage 1 pre-import completion.

---

## 🧪 Verification Plan

### Automated Verification
- Verify UTF-8 BOM (`EF BB BF`) on all modified and created files.
- Compile entire solution with `MSBuild` to confirm 0 Errors and 0 Warnings.

### Functional Verification
- Test CLI batch runner with sample EDI CSV files end-to-end against local Oracle DB:
  - Working calendar due date calculation (`ADD_WORKING_DAYS2`).
  - `0630` quantity addition to `0600`.
  - `663N` due date update and `663S` quantity update to `6624`.
  - Ship-to conversion via `SUZUKI_SPIRITS_CONVERSION_MST`.
  - Forecast consumption (`0650/0651` consumes `0600`, `0814` consumes `0813`, `6624` consumes `6604`).
  - Commitment to `ORDERS` and `ORDER_HISTORY`.
  - `STRA納期設定` execution.
