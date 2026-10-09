# Suzuki SPIRITS Batch Import (`.exe`) Test Checklist

> **Target Application:** `SUZUKI_SPIRITS_DATA_IMPORT.exe`  
> **Target Environment:** Localhost Oracle Database (`localhost:1521/OMSDB`)  
> **Target Customer:** Suzuki (`5455`) | **Folder Type:** `4 (混在)`  
> **Intake Folder:** `C:\ASTI\Order\Suzuki`

---

## [ ] Step 1: Pre-Execution Setup & Safety Check

- [ ] **1.1 Check App.config Connection String**
      Open `SUZUKI_SPIRITS_DATA_IMPORT\App.config` and ensure the connection string points to **localhost**:

  ```xml
  <connectionStrings>
    <!-- Localhost environment -->
    <add name="OMSConnection" connectionString="User Id=OMSDB;Password=Amagata001;Data Source=//localhost:1521/OMSDB;" providerName="Oracle.ManagedDataAccess.Client" />
    <!-- Aska OracleDB environment (MUST BE COMMENTED OUT) -->
    <!-- <add name="OMSConnection" connectionString="User Id=OMSDB;Password=Amagata001;Data Source=//192.168.99.126:1521/orcl;" providerName="Oracle.ManagedDataAccess.Client" /> -->
  </connectionStrings>
  ```

- [x] **1.2 Clean Localhost Database**
      Run the cleanup SQL in your Oracle client:
  ```sql
  DELETE FROM ORDERS_STAGE WHERE CUSTOMER_CODE = '5455';
  DELETE FROM ORDERS WHERE CUSTOMER_CODE = '5455';
  DELETE FROM SUZUKI_SPIRITS_0600AND0630;
  DELETE FROM SUZUKI_SPIRITS_0602;
  DELETE FROM SUZUKI_SPIRITS_0650;
  DELETE FROM SUZUKI_SPIRITS_0814;
  DELETE FROM SUZUKI_SPIRITS_6604AND6634;
  DELETE FROM SUZUKI_SPIRITS_6624;
  DELETE FROM SUZUKI_SPIRITS_663NAND663SAND664T;
  DELETE FROM IMP_FILES_STAGE WHERE CUSTOMER_SETTING_ID IN (
      SELECT CUSTOMER_SETTING_ID FROM CUSTOMER_SETTING_MST WHERE CUSTOMER_CODE = '5455'
  );
  DELETE FROM IMP_FILES WHERE CUSTOMER_SETTING_ID IN (
      SELECT CUSTOMER_SETTING_ID FROM CUSTOMER_SETTING_MST WHERE CUSTOMER_CODE = '5455'
  );
  COMMIT;
  ```

---

## [ ] Step 2: Prepare Intake Files

- [x] **2.1 Clear Intake Directory**
      In `C:\ASTI\Order\Suzuki\`, delete any existing `.csv` or `.txt` files (keep `bk` and `エラーリスト` folders).

- [x] **2.2 Copy Test Files**
      Copy the following 5 files from `CSV_Example\TEST_SUITE_SMALL\` into `C:\ASTI\Order\Suzuki\`:
  - [x] `0600-test.csv` (Monthly Base Forecast)
  - [x] `0630-test.csv` (Quantity Adjustment +100)
  - [x] `0602-test.csv` (Supplemental Forecast)
  - [x] `0650-test.csv` (Firm Delivery Instructions)
  - [x] `0814-test.csv` (Supplied Parts Firm Delivery)

---

## [ ] Step 3: Run the Batch Executable

- [x] **3.1 Execute `SUZUKI_SPIRITS_DATA_IMPORT.exe`**
      Choose either method:
  - **Method A:** Double-click `SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug\SUZUKI_SPIRITS_DATA_IMPORT.exe`
  - **Method B:** Run in PowerShell to see log output:
    ```powershell
    cd c:\Dev\Projects\ASTI\OrderManagementSystem-dev\SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug
    .\SUZUKI_SPIRITS_DATA_IMPORT.exe
    ```

- [x] **3.2 Verify Console Finished with Exit Code 0**
  - [x] No exception popup or error message.
  - [x] Message displayed: `Pre-Import processed 5 file(s)` / `Order Registration completed successfully`.

---

## [ ] Step 4: Verify File Movement

- [x] **4.1 Intake Folder Emptied**
      Check `C:\ASTI\Order\Suzuki\`:
  - [x] The 5 test `.csv` files have been moved (no longer in `C:\ASTI\Order\Suzuki\`).

- [x] **4.2 No Error List Generated**
      Check `C:\ASTI\Order\Suzuki\エラーリスト\`:
  - [x] No Excel error file was created.

- [ ] **4.3 Files Archived to COMPLETED**
      Check `C:\ASTI\DATA\COMPLETED\SUZUKI-AT\5455\4\`:
  - [ ] The 5 test files are archived with timestamp suffix (e.g., `0600-test_SUZUKI-AT_*.csv`).
  - [ ] `C:\ASTI\DATA\WORK\SUZUKI-AT\5455\4\` is empty.

---

## [ ] Step 5: Verify Database Records in `ORDERS`

> **Note on Status in Batch Mode:**  
> Because the batch executable automatically runs Step 8 (**STRA 納期設定**) at the end of the run, all newly registered active orders transition directly from `'PROCESSED'` to **`'DUE_SET'`** with `HAND_FLAG = 'N'`.

Run this verification SQL in your Oracle client:

```sql
SELECT
    o.ORDER_ID,
    o.INFO_TYPE_CODE,
    o.HAND_FLAG,
    o.CUSTOMER_ITEM_NO,
    o.CUSTOMER_ORDER_NO,
    o.DUE_DATE,
    o.DEMAND_QTY,
    o.ORDER_TYPE,
    o.STATUS,
    o.ACTIVE_FLAG,
    o.CREATED_USER_ID,
    o.CREATED_PG_ID
FROM ORDERS o
WHERE o.CUSTOMER_CODE = '5455'
ORDER BY o.ORDER_ID;
```

Check the results against these criteria:

- [ ] **5.1 Created User Check**
  - [ ] `CREATED_USER_ID` is `'SUZUKI-AT'` across all records.

- [ ] **5.2 Quantity Adjustment Check (`0630`)**
  - [ ] Item `36890-85030-000` on `2026-10-02` has `DEMAND_QTY = 244` (144 base + 100 adjustment), `STATUS = 'DUE_SET'`, `HAND_FLAG = 'N'`, `ACTIVE_FLAG = 'Y'`.

- [ ] **5.3 Supplemental Forecast Check (`0602`)**
  - [ ] Item `36890-85030-000` on `2026-09-15` has `DEMAND_QTY = 114`, `STATUS = 'DUE_SET'`, `HAND_FLAG = 'N'`, `ACTIVE_FLAG = 'Y'`.

- [ ] **5.4 Delivery Instructions Check (`0650`)**
  - [ ] Firm orders (`ORDER_TYPE = 3`) created for `09403-13402-000` and `09407-17403-000` with `STATUS = 'DUE_SET'`, `ACTIVE_FLAG = 'Y'`.

- [ ] **5.5 Supplied Parts Check (`0814`)**
  - [ ] Firm orders (`ORDER_TYPE = 3`) created with quantities `2000`, `6000`, `2000`, `STATUS = 'DUE_SET'`, `ACTIVE_FLAG = 'Y'`.

- [ ] **5.6 Forecast Reconciliation Check (消し込み)**
  - [ ] Base forecast for `09407-17403-000` on `2026-09-14` has `DEMAND_QTY = 0`, `STATUS = 'RECONCILED'`, `ACTIVE_FLAG = 'N'`.

- [ ] **5.7 Future Month Bucket Check**
  - [ ] Item `36890-85030-000` on `2026-11-02` has `STATUS = 'REPLACED'`, `ACTIVE_FLAG = 'N'`.

---

## [ ] Step 6: Verify `ORDERS_STAGE` Table is Clean

Run this SQL in your Oracle client:

```sql
SELECT COUNT(*) FROM ORDERS_STAGE WHERE CUSTOMER_CODE = '5455';
```

- [ ] **6.1 Zero Leftover Rows in `ORDERS_STAGE`**
  - [ ] Count is `0`.
- [ ] **6.2 Clean Web UI Screen**
  - [ ] Open `OrderImport.aspx` in browser: the bottom list (`gvImportOrder`) is completely empty. No need to click [保存].
