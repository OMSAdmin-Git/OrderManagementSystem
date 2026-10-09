---
pdf_options:
  format: a4
  margin: 15mm 15mm 15mm 15mm
  printBackground: true
---


# 📄 Suzuki (SPIRITS) Data Import End-to-End Test Procedure
**Document Version:** 1.0  
**Target Git Branch:** `feature/suzuki-import`  
**Target Customer:** Suzuki Motor Corporation (`CUSTOMER_CODE = '5455'`)  
**Source CSV Directory:** `C:\Dev\Projects\ASTI\OrderManagementSystem-dev\CSV_Example`  

---

## 📌 1. Test Overview & Prerequisites

### 1.1 Target Branch
- **Branch**: `feature/suzuki-import`
  - **Stage 1 (Pre-Import)**: Auto-invalidation of previous month Nagare orders, info code whitelist filtering (skips deprecated `0813`).
  - **Stage 2 (Order Registration)**: `0814 -> 0600` forecast reconciliation, `0813` exclusion, `0602` supplemental forecast retention.
  - **STRA Delivery Settings**: Automated calculation and reflection of shipping schedules and delivery dates for firm orders.

### 1.2 Prerequisites Checklist
- [ ] Connected to the local development Oracle database (**NEVER** connect to or modify production DB).
- [ ] In `CUSTOMER_SETTING_MST` / `IMP_RULE_SETTING`, Suzuki (`customer_code = '5455'`) has `SpProcessType = 1` (Suzuki).
- [ ] In `FOLDER_MST`, Suzuki has an active folder registered with `FOLDER_TYPE = 4` (Mixed / 混在).
- [ ] Execution mode:
  - **Automated Batch Mode (Recommended)**:
    ```cmd
    SUZUKI_SPIRITS_DATA_IMPORT.exe /batch
    ```
    Executes Stage 1 → Stage 2 → STRA Delivery Settings end-to-end automatically.
  - **Web UI Mode**:
    1. Navigate to `OrderImportStage.aspx` (取込準備), select Suzuki, and click **取込準備**.
    2. Navigate to `OrderImport.aspx` (受注ファイル取込) and execute order registration.

---

## 🗂️ 2. Suzuki Data File Types Reference (15 Codes)

| Info Code | Data Name | Type / Pipeline Function |
|:---:|:---|:---|
| **`0600`** | Monthly Forecast (月次内示) | **[Baseline Forecast]** Multi-delivery schedule matrix. In Stage 2, subject to washing and consumed/reconciled by delivery instructions (`0650`, `0651`, `0814`). |
| **`0602`** | Supplemental Forecast (追加内示) | **[Additional Forecast]** Excluded from forecast washing (洗い替え除外). Appended directly without wiping `0600`. |
| **`0630`** | Forecast Adjustment (内示増減) | **[Quantity Adjustment]** Adjusts quantity (`DEMAND_QTY`) of matching `0600` rows. |
| **`0650`** | Regular Delivery Instruction (通常指示) | **[Firm Instruction]** Reconciles against `0600` (`0650 -> 0600`). |
| **`0651`** | Temporary Delivery Instruction (臨時指示) | **[Firm Instruction]** Reconciles against `0600` (`0651 -> 0600`). |
| **`0814`** | Supplied Parts Instruction (支給品納入指示) | **[Firm Instruction]** Per 2026/09/25 update, reconciles against **`0600`** (`0814 -> 0600`). |
| **`6604`** | PO Forecast (注文内示) | **[EDI PO Forecast]** Subject to forecast washing; reconciled by `6624`. |
| **`6624`** | Firm Purchase Order (注文確定) | **[Firm PO]** Reconciles against `6604` (`6624 -> 6604`). |
| **`6634`** | PO Forecast Adjustment (注文内示増減) | **[Quantity Adjustment]** Adjusts quantity of matching `6604` rows. |
| **`663N`** | PO Due Date Change (納期変更) | **[PO Modification]** Updates `DUE_DATE` on matching `6624` firm PO. |
| **`663S`** | PO Quantity Change (数量変更) | **[PO Modification]** Updates `DEMAND_QTY` on matching `6624` firm PO. |
| **`664T`** | PO Cancellation (注文取消) | **[PO Cancellation]** Cancels matching `6624` order (`active_flag = 'N'`). |
| **`0501`** | PO Receipt (受領書) | **[Reference Notice]** Ingested into staging table `SUZUKI_SPIRITS_0501`. |
| **`0502`** | Inspection Notice (検収書) | **[Reference Notice]** Ingested into staging table `SUZUKI_SPIRITS_0502`. |
| **`0740`** | Shipping Request (出荷依頼) | **[Logistics Notice]** Ingested into staging table `SUZUKI_SPIRITS_0740`. |

<div class="page-break"></div>

---

## 🚀 3. Step-by-Step Phased Test Execution

<div class="test-phase">

### Phase 1: Baseline Forecast Ingestion (`0600`, `6604`)
**Objective**: Verify that baseline monthly forecasts are ingested and registered into `ORDERS` as `ORDER_TYPE = 1`.

1. **Place Files in Input Folder**:
   - `0600-260819.csv`
   - `6604-260827.csv`
2. **Execute**:
   - Run batch command: `SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`
3. **Verification Items**:
   - [ ] Rows exist in `SUZUKI_SPIRITS_0600AND0630` with `STATUS = 'REGISTERED'`.
   - [ ] Rows exist in `SUZUKI_SPIRITS_6604AND6634`.
   - [ ] In `ORDERS` table: Records exist with `CUSTOMER_CODE = '5455'`, `ORDER_TYPE = 1` (Forecast), and `ACTIVE_FLAG = 'Y'`.
</div>

<div class="test-phase">

### Phase 2: Supplemental Forecast & Quantity Adjustment (`0602`, `6634`, `0630`)
**Objective**: Verify that supplemental forecasts do NOT wipe baseline `0600` forecasts, and quantity adjustments correctly alter target forecast rows.

1. **Place Files in Input Folder**:
   - `0602-260817.csv`
   - `6634-260827.csv`
   - `0630-260904.csv`
2. **Execute**:
   - Run batch command: `SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`
3. **Verification Items**:
   - [ ] `0602` records are added to `ORDERS` as `ORDER_TYPE = 1`.
   - [ ] Existing `0600` records from Phase 1 remain `ACTIVE_FLAG = 'Y'` (confirms washing exclusion).
   - [ ] Target `0600` row quantities are adjusted by `0630`.
</div>

<div class="test-phase">

### Phase 3: Delivery Instructions & Forecast Reconciliation (`0650`, `0651`, `0814`, `6624`)
**Objective**: Verify that delivery instructions consume/reduce the forecast quantities of `0600` and `6604`.

1. **Place Files in Input Folder**:
   - `0650-260817.csv` (Regular Instruction)
   - `0651-260817.csv` (Temporary Instruction)
   - `0814-260817.csv` (Supplied Parts Instruction)
   - `6624-260817.csv` (Firm Purchase Order)
2. **Execute**:
   - Run batch command: `SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`
3. **Verification Items**:
   - [ ] **0814 Target Check**: `0814` reconciles against **`0600`** (NOT `0813`), reducing the `0600` forecast quantity.
   - [ ] `0650` and `0651` reconcile against `0600`.
   - [ ] `6624` reconciles against `6604`.
   - [ ] Firm instruction records are inserted into `ORDERS` with `ORDER_TYPE = 3` and `ACTIVE_FLAG = 'Y'`.
</div>

<div class="page-break"></div>

<div class="test-phase">

### Phase 4: PO Modifications & Cancellations (`663N`, `663S`, `664T`)
**Objective**: Verify that PO changes and cancellations update existing firm orders.

1. **Place Files in Input Folder**:
   - `663N-260817.csv` (Due Date Change)
   - `663S-260911.csv` (Quantity Change)
   - `664T-260911.csv` (Order Cancellation)
2. **Execute**:
   - Run batch command: `SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`
3. **Verification Items**:
   - [ ] Matching `6624` row's `DUE_DATE` is updated by `663N`.
   - [ ] Matching `6624` row's `DEMAND_QTY` is updated by `663S`.
   - [ ] Matching `6624` row is cancelled (`ACTIVE_FLAG = 'N'`) by `664T`.
</div>

<div class="test-phase">

### Phase 5: Reference / Inspection & Shipping Data (`0501`, `0502`, `0740`)
**Objective**: Verify error-free staging of reference notices.

1. **Place Files in Input Folder**:
   - `0501-260911.csv`
   - `0502-260911.csv`
   - `0740-260817.csv`
2. **Execute**:
   - Run batch command: `SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`
3. **Verification Items**:
   - [ ] Rows successfully inserted into `SUZUKI_SPIRITS_0501`.
   - [ ] Rows successfully inserted into `SUZUKI_SPIRITS_0502`.
   - [ ] Rows successfully inserted into `SUZUKI_SPIRITS_0740`.
</div>

<div class="test-phase">

### Phase 6: STRA Delivery Setting Execution Check
**Objective**: Verify downstream STRA delivery schedule calculation.

1. **Verification Items**:
   - [ ] Console / `IMP_RUN` log displays `[STRA納期設定処理開始]` and `[STRA納期設定処理完了]`.
   - [ ] Firm order rows in `ORDERS` have populated `SHIP_SCHEDULED_DATE` values.
</div>

---

## 🔍 4. Verification SQL Queries

```sql
-- 1. Check orders created for Suzuki
SELECT order_id, customer_code, customer_order_no, order_type, info_type_code, 
       item_no, demand_qty, due_date, ship_scheduled_date, active_flag, updated_at
FROM orders
WHERE customer_code = '5455'
ORDER BY updated_at DESC, order_id DESC;

-- 2. Verify 0814 and 0650 reconciliation against 0600
SELECT info_type_code, order_type, item_no, demand_qty, due_date, active_flag
FROM orders
WHERE customer_code = '5455'
  AND info_type_code IN ('0600', '0814', '0650', '0651')
ORDER BY item_no, due_date;

-- 3. Staging table ingestion counts and status
SELECT '0600' AS code, status, COUNT(*) FROM suzuki_spirits_0600and0630 GROUP BY status
UNION ALL
SELECT '0602', status, COUNT(*) FROM suzuki_spirits_0602 GROUP BY status
UNION ALL
SELECT '0650', status, COUNT(*) FROM suzuki_spirits_0650 GROUP BY status
UNION ALL
SELECT '0814', status, COUNT(*) FROM suzuki_spirits_0814 GROUP BY status
UNION ALL
SELECT '6624', status, COUNT(*) FROM suzuki_spirits_6624 GROUP BY status;
```

<div class="page-break"></div>

---

## 📝 5. Test Sign-off Sheet (Print & Fill In)

| Test Phase | Execution Date | Tester | Status | Notes / Observations |
|---|:---:|:---:|:---:|---|
| **Phase 1: Baseline Forecast (`0600`, `6604`)** | `____/__/__` | | [ ] Pass<br>[ ] Fail | |
| **Phase 2: Supplemental & Adjustment (`0602`, `0630`)** | `____/__/__` | | [ ] Pass<br>[ ] Fail | `0602` retained without wipe |
| **Phase 3: Delivery Instructions (`0650`, `0814`)** | `____/__/__` | | [ ] Pass<br>[ ] Fail | `0814 -> 0600` consumed |
| **Phase 4: Order Changes & Cancel (`663N`, `664T`)** | `____/__/__` | | [ ] Pass<br>[ ] Fail | Due date & cancel verified |
| **Phase 5: Reference Data (`0501`, `0740`)** | `____/__/__` | | [ ] Pass<br>[ ] Fail | Staged cleanly |
| **Phase 6: STRA Delivery Settings Integration** | `____/__/__` | | [ ] Pass<br>[ ] Fail | Schedules calculated |

---
**Lead QA / Developer Signature:** ___________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **Date:** _______________
