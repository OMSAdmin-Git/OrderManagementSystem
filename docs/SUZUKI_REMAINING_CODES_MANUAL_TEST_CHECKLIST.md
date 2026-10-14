# 📝 Suzuki SPIRITS Remaining Codes Manual Test Checklist

> **Target Application:** `SUZUKI_SPIRITS_DATA_IMPORT.exe` (or Web UI Order Import)  
> **Target Environment:** Localhost Oracle Database (`localhost:1521/OMSDB`)  
> **Customer:** Suzuki (`5455`) | **Folder Type:** `4 (混在)`  
> **Intake Folder:** `C:\ASTI\Order\Suzuki\`  
> **Source Test Files Folder:** `c:\Dev\Projects\ASTI\OrderManagementSystem-dev\CSV_Example\TEST_SUITE_SMALL\`  

---

## 🧹 Step 1: Pre-Execution Cleanup (Localhost DB)

Run this SQL query in your Oracle SQL client (PL/SQL Developer, SQL Developer, or DBeaver) to clear old test data on localhost:

```sql
DELETE FROM ORDERS_STAGE WHERE customer_code = '5455';
DELETE FROM ORDERS WHERE customer_code = '5455';
DELETE FROM SUZUKI_SPIRITS_0600AND0630;
DELETE FROM SUZUKI_SPIRITS_0602;
DELETE FROM SUZUKI_SPIRITS_0650;
DELETE FROM SUZUKI_SPIRITS_0651;
DELETE FROM SUZUKI_SPIRITS_0814;
DELETE FROM SUZUKI_SPIRITS_6604AND6634;
DELETE FROM SUZUKI_SPIRITS_6624;
DELETE FROM SUZUKI_SPIRITS_663NAND663SAND664T;
DELETE FROM SUZUKI_SPIRITS_0501AND0502;
DELETE FROM SUZUKI_SPIRITS_0740;
DELETE FROM IMP_FILES_STAGE WHERE customer_setting_id IN (
    SELECT customer_setting_id FROM customer_setting_mst WHERE customer_code = '5455'
);
DELETE FROM IMP_FILES WHERE customer_setting_id IN (
    SELECT customer_setting_id FROM customer_setting_mst WHERE customer_code = '5455'
);
COMMIT;
```

---

## 🏍️ Step 2: Test Parent Orders & Forecast (`6624-test.csv`, `6604-test.csv`, `0651-test.csv`)

### 2.1 File Placement
1. Ensure `C:\ASTI\Order\Suzuki\` is empty.
2. Copy the following 3 files from `CSV_Example\TEST_SUITE_SMALL\` to `C:\ASTI\Order\Suzuki\`:
   - `6624-test.csv` (Two-wheeler fixed orders: includes `D25293H` and `F99854G`)
   - `6604-test.csv` (Two-wheeler forecast orders: Type 1)
   - `0651-test.csv` (Four-wheeler spot delivery orders: Type 3)

### 2.2 Execution
Run the batch executable:
```powershell
cd c:\Dev\Projects\ASTI\OrderManagementSystem-dev\SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug
.\SUZUKI_SPIRITS_DATA_IMPORT.exe
```

### 2.3 Verification SQL Queries
```sql
-- A. Verify Staging Status
SELECT '6624' AS CODE, COUNT(*) AS CNT, status FROM SUZUKI_SPIRITS_6624 GROUP BY status
UNION ALL
SELECT '6604', COUNT(*), status FROM SUZUKI_SPIRITS_6604AND6634 GROUP BY status
UNION ALL
SELECT '0651', COUNT(*), status FROM SUZUKI_SPIRITS_0651 GROUP BY status;
-- Expected: All rows show STATUS = 'REGISTERED'

-- B. Verify ORDERS Registration
SELECT customer_order_no, customer_info_type, order_type, demand_status, demand_qty, due_date, pre_daily_delivery_date
FROM orders
WHERE customer_code = '5455' AND active_flag = 'Y'
ORDER BY customer_info_type, customer_order_no;

-- Expected:
-- Order D25293H (6624): DEMAND_QTY = 10, PRE_DAILY_DELIVERY_DATE = 2026-12-02
-- Order F99854G (6624): DEMAND_QTY = 30, PRE_DAILY_DELIVERY_DATE = 2026-09-14
```

---

## 🔄 Step 3: Test Adjustments (`663N-test.csv` & `663S-test.csv`)

### 3.1 File Placement
1. Ensure `C:\ASTI\Order\Suzuki\` is empty.
2. Copy the following 2 modifier files from `CSV_Example\TEST_SUITE_SMALL\` to `C:\ASTI\Order\Suzuki\`:
   - `663N-test.csv` (Due Date Modifier targeting `D25293H`)
   - `663S-test.csv` (Quantity Modifier targeting `F99854G`)

### 3.2 Execution
Run the batch executable:
```powershell
cd c:\Dev\Projects\ASTI\OrderManagementSystem-dev\SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug
.\SUZUKI_SPIRITS_DATA_IMPORT.exe
```

### 3.3 Verification SQL Queries
```sql
-- A. Verify 663N/663S Staging Status
SELECT info_type_code, customer_order_no, status, updated_at 
FROM SUZUKI_SPIRITS_663NAND663SAND664T;
-- Expected: STATUS = 'REGISTERED'

-- B. Verify ORDERS Adjustments on Parent Orders
SELECT 
    customer_order_no, 
    customer_info_type, 
    order_type, 
    demand_qty, 
    pre_daily_order_qty,
    due_date, 
    pre_daily_delivery_date, 
    updated_at
FROM orders
WHERE customer_order_no IN ('D25293H', 'F99854G')
  AND active_flag = 'Y';

-- Expected Results:
-- 1. D25293H: PRE_DAILY_DELIVERY_DATE was 2026-12-02 -> NOW UPDATED to 2026-10-23!
-- 2. F99854G: DEMAND_QTY was 30 -> NOW UPDATED to 0 (or new qty per 663S)!
```

---

## 📑 Step 4: Test Staging-Only Reference Data (`0501-test.csv` & `0740-test.csv`)

### 4.1 File Placement
1. Ensure `C:\ASTI\Order\Suzuki\` is empty.
2. Copy the following 2 reference files from `CSV_Example\TEST_SUITE_SMALL\` to `C:\ASTI\Order\Suzuki\`:
   - `0501-test.csv` (Confirmed forecast - staging only)
   - `0740-test.csv` (Inspection/acceptance data - staging only)

### 4.2 Execution
Run the batch executable:
```powershell
cd c:\Dev\Projects\ASTI\OrderManagementSystem-dev\SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug
.\SUZUKI_SPIRITS_DATA_IMPORT.exe
```

### 4.3 Verification SQL Queries
```sql
-- A. Verify Staging Insertion & Settlement
SELECT '0501' AS CODE, COUNT(*) AS CNT, status FROM SUZUKI_SPIRITS_0501AND0502 GROUP BY status
UNION ALL
SELECT '0740', COUNT(*), status FROM SUZUKI_SPIRITS_0740 GROUP BY status;
-- Expected: 4 rows each, STATUS = 'REGISTERED'

-- B. Verify NO ORDERS were created
SELECT COUNT(*) AS should_be_zero 
FROM orders 
WHERE customer_code = '5455' 
  AND customer_info_type IN ('0501', '0740');
-- Expected: 0 rows (confirms these codes remain staging-only per specification)
```

---

## 📁 Step 5: File Settlement Verification

After all tests:
- Check `C:\ASTI\Order\Suzuki\`: Must be empty.
- Check `C:\ASTI\Order\Suzuki\エラーリスト\`: No error files.
- Check `C:\ASTI\DATA\COMPLETED\SUZUKI-AT\5455\4\`: All processed `.csv` files are archived with user/timestamp suffixes.
