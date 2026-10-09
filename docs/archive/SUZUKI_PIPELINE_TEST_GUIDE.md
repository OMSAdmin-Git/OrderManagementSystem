# Suzuki SPIRITS EDI Pipeline Testing & Operation Guide

**Document Version:** 1.0  
**Target Customer:** スズキ (5455)  
**Target Branch:** `feature/suzuki-import`  
**Date:** 2026-10-06  

---

## 1. Overview of Suzuki EDI Codes

Suzuki SPIRITS uses specialized 4-character information type codes (`INFO_TYPE_CODE`). Understanding their business roles is essential for folder distribution and test execution:

| Code | Classification | Category Name (JP) | Description & System Behavior |
| :---: | :---: | :--- | :--- |
| **`0600`** | Forecast | 月次内示 (Base Monthly Forecast) | Ingests the baseline monthly forecast orders (`ORDER_TYPE = 1`). Staged into `ORDERS_STAGE` and committed to `ORDERS`. |
| **`0602`** | Forecast | 追加内示 (Supplemental Forecast) | Ingests supplemental forecast orders (`ORDER_TYPE = 1`). **Never washed out** by regular forecast reconciliation. |
| **`0630`** | Adjustment | 内示増減 (Forecast Qty Adjustment) | **Not an independent order**. Modifies the quantity of a matching `0600` forecast row: `DEMAND_QTY = DEMAND_QTY + 0630.ORDER_QTY`. If `0600` is already saved in `ORDERS`, it retrieves it via fallback and updates it. |
| **`0650`** | Delivery | 納入指示・通常 (Firm Delivery) | Firm delivery instructions (`ORDER_TYPE = 3`). Consumes and reconciles existing `0600` forecast lines. |
| **`0651`** | Delivery | 納入指示・臨時 (Emergency Delivery)| Emergency firm delivery instructions (`ORDER_TYPE = 3`). Consumes existing `0600` forecast lines. |
| **`0814`** | Delivery | 支給品納入指示 (Supplied Parts) | Firm delivery instructions for supplied parts (`ORDER_TYPE = 3`). Consumes matching `0600` forecast lines. |
| **`6604`** | Forecast | 注文内示 (PO Forecast) | Purchase order forecast (`ORDER_TYPE = 2`). |
| **`6624`** | Delivery | 注文確定 (PO Firm Order) | Purchase order firm delivery instructions (`ORDER_TYPE = 3`). |
| **`663N`** | Adjustment | 注文納期変更 (**N** = 納期 Due Date) | Modifies the **Due Date (`DUE_DATE`)** of a matching `6624` order based on customer PO number (`客先発注No`). |
| **`663S`** | Adjustment | 注文数量変更 (**S** = 数量 Quantity) | Modifies the **Quantity (`DEMAND_QTY`)** of a matching `6624` order based on customer PO number (`客先発注No`). |

---

## 2. Directory Architecture & Folder Setup

According to ASTI's operation specifications, files must be distributed based on folder type:

```
C:\ASTI\Order\Fukuroi\
├── 5455_E1_N\    [FolderType = 1 : 内示フォルダ (Forecast & Adjustments)]
│   ├── 0600-test.csv
│   ├── 0602-test.csv
│   └── 0630-test.csv
│
├── 5455_E1_S\    [FolderType = 3 : 納入指示フォルダ (Firm Deliveries & Adjustments)]
│   ├── 0650-test.csv
│   └── 0814-test.csv
│
└── 5455_E1_M\    [FolderType = 4 : 混在フォルダ (Automated Batch Only)]
    └── (Used exclusively by SUZUKI_SPIRITS_DATA_IMPORT.exe /batch)
```

### Folder Mapping Reference Table

| Folder Path | Folder Type | Purpose | Expected Files |
| :--- | :---: | :--- | :--- |
| `C:\ASTI\Order\Fukuroi\5455_E1_N` | **1 (内示)** | Manual Web UI Import (Forecasts) | `0600`, `0602`, `0630`, `6604`, ASTI追加内示 Excel |
| `C:\ASTI\Order\Fukuroi\5455_E1_S` | **3 (納入指示)** | Manual Web UI Import (Firm Orders) | `0650`, `0651`, `0814`, `6624`, `663N`, `663S` |
| `C:\ASTI\Order\Fukuroi\5455_E1_M` | **4 (混在)** | Automated Batch Process | All files combined (EDI automated dump) |

> **Note:** Folders marked `ACTIVE_FLAG = 'N'` in `FOLDER_MST` are automatically ignored during Stage 1 staging preparation.

---

## 3. Test Suite Data Details

The test files located in `CSV_Example/TEST_SUITE_SMALL/` contain the following test cases:

| File Name | Test Target Part No | Action / Test Purpose | Expected Value |
| :--- | :--- | :--- | :--- |
| `0600-test.csv` | `36890-85030-000` | Ingests base forecast for 2026-10-02 | Initial `DEMAND_QTY = 144` |
| `0600-test.csv` | `09407-17403-000` | Ingests base forecast for 2026-09-14 | Initial `DEMAND_QTY = 348` |
| `0602-test.csv` | `36890-85030-000` | Supplemental forecast for 2026-09-15 | Preserved, `DEMAND_QTY = 114` |
| `0630-test.csv` | `36890-85030-000` | Quantity adjustment (+100) on 2026-10-02 | Updates `DEMAND_QTY`: `144 + 100 = 244` |
| `0650-test.csv` | Multiple parts | Firm delivery instructions (4 rows) | `STATUS = 'PROCESSED'` |
| `0814-test.csv` | `09407-17403-000` | Supplied parts deliveries (10,000 total) | Consumes `0600` row $\rightarrow$ `DEMAND_QTY = 0, RECONCILED` |

---

## 4. Step-by-Step Test Execution Procedure

### Step 0: Clean Test Database (Localhost)
Run the following SQL in your database client (Oracle):
```sql
DELETE FROM ORDERS_STAGE WHERE CUSTOMER_CODE = '5455';
DELETE FROM ORDERS WHERE CUSTOMER_CODE = '5455';
COMMIT;
```

---

### Step 1: Stage 1 Staging Preparation (`OrderImportStage.aspx`)
1. Place the test files in their physical directories:
   - Copy `0600-test.csv`, `0602-test.csv`, and `0630-test.csv` into `C:\ASTI\Order\Fukuroi\5455_E1_N`.
   - Copy `0650-test.csv` and `0814-test.csv` into `C:\ASTI\Order\Fukuroi\5455_E1_S`.
2. Open the web browser and navigate to `OrderImportStage.aspx`.
3. Locate Customer **5455** (PC: **E1**) in the grid.
4. Check the selection checkbox and click **取込準備 (Prepare Import)**.
5. **Verification**: 
   - A success message indicates 5 files were prepared into `IMP_FILES_STAGE`.
   - Inactive folders (e.g. `5455_E1_M`) are not scanned.

---

### Step 2: Ingest Base Forecast (`0600-test.csv`)
1. Open `OrderImport.aspx`.
2. In the top table (`gvImpFilesStage`), check **`0600-test.csv`**.
3. Click **取込実行 (Execute Import)**.
4. Verify 15 rows appear in the bottom grid (`gvImportOrder`).
5. Click **保存 (Save)**.
   - *Result*: Base forecast records are committed to `ORDERS`.

---

### Step 3: Test Standalone Adjustment (`0630-test.csv`)
1. On `OrderImport.aspx`, check **`0630-test.csv`** alone in the top table.
2. Click **取込実行 (Execute Import)**.
3. **Observation**:
   - The system executes the fallback query against `ORDERS`, finds the matching `0600` forecast row (`36890-85030-000` on `2026-10-02`), and applies `+100`.
   - The adjusted row appears in the bottom table with `DEMAND_QTY = 244`.
4. Click **保存 (Save)**.
   - *Result*: The system executes a `MERGE` into `ORDERS`, updating the quantity from 144 to 244.

---

### Step 4: Ingest Supplemental Forecast (`0602-test.csv`)
1. Check **`0602-test.csv`** in the top table.
2. Click **取込実行 (Execute Import)** $\rightarrow$ verify bottom grid $\rightarrow$ click **保存 (Save)**.
   - *Result*: Supplemental forecast is committed and marked exempt from wash.

---

### Step 5: Ingest Delivery Instructions (`0650-test.csv` & `0814-test.csv`)
1. Check **`0650-test.csv`** $\rightarrow$ click **取込実行** $\rightarrow$ click **保存**.
2. Check **`0814-test.csv`** $\rightarrow$ click **取込実行** $\rightarrow$ click **保存**.
   - *Result*: Firm orders are saved, and the matching forecast for `09407-17403-000` is consumed and reconciled.

---

## 5. SQL Verification Queries

Run the queries below to verify all requirements:

### Check A: Summary of All Active Orders
```sql
SELECT 
    INFO_TYPE_CODE,
    CASE ORDER_TYPE 
        WHEN 1 THEN '1: Forecast (内示)'
        WHEN 2 THEN '2: PO Forecast (注文内示)'
        WHEN 3 THEN '3: Delivery Instruction (納入指示)'
        ELSE TO_CHAR(ORDER_TYPE)
    END AS ORDER_TYPE_DESC,
    STATUS,
    ACTIVE_FLAG,
    COUNT(*) AS ROW_COUNT,
    SUM(DEMAND_QTY) AS TOTAL_QTY
FROM ORDERS
WHERE CUSTOMER_CODE = '5455'
GROUP BY INFO_TYPE_CODE, ORDER_TYPE, STATUS, ACTIVE_FLAG
ORDER BY INFO_TYPE_CODE, ORDER_TYPE;
```

### Check B: Verify `0630` Adjustment (+100 Addition)
```sql
SELECT 
    ORDER_ID,
    ITEM_NO,
    CUSTOMER_ITEM_NO,
    INFO_TYPE_CODE,
    DEMAND_QTY,
    TO_CHAR(DUE_DATE, 'YYYY-MM-DD') AS DUE_DATE,
    STATUS,
    ACTIVE_FLAG
FROM ORDERS
WHERE CUSTOMER_CODE = '5455'
  AND CUSTOMER_ITEM_NO = '36890-85030-000'
ORDER BY DUE_DATE;
```
*Expected Result*: Row with `DUE_DATE = '2026-10-02'` shows **`DEMAND_QTY = 244`**.

### Check C: Verify `0814` Delivery & `0600` Consumption
```sql
SELECT 
    ORDER_ID,
    ITEM_NO,
    CUSTOMER_ITEM_NO,
    INFO_TYPE_CODE,
    ORDER_TYPE,
    DEMAND_QTY,
    TO_CHAR(DUE_DATE, 'YYYY-MM-DD') AS DUE_DATE,
    STATUS,
    ACTIVE_FLAG
FROM ORDERS
WHERE CUSTOMER_CODE = '5455'
  AND CUSTOMER_ITEM_NO = '09407-17403-000'
ORDER BY DUE_DATE, ORDER_TYPE;
```
*Expected Result*: Firm delivery orders (`0814`) are present; the matching `0600` row on `2026-09-14` has **`DEMAND_QTY = 0`** and **`STATUS = 'RECONCILED'`** (`ACTIVE_FLAG = 'N'`).

### Check D: Verify `0602` Supplemental Retention
```sql
SELECT 
    ORDER_ID,
    ITEM_NO,
    CUSTOMER_ITEM_NO,
    INFO_TYPE_CODE,
    DEMAND_QTY,
    STATUS,
    ACTIVE_FLAG
FROM ORDERS
WHERE CUSTOMER_CODE = '5455'
  AND INFO_TYPE_CODE = '0602'
ORDER BY ORDER_ID;
```
*Expected Result*: 1 row for `36890-85030-000` with `DEMAND_QTY = 114` and `ACTIVE_FLAG = 'Y'` is preserved.
