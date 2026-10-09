# SPIRITS Data Import Process Specification Document (Suzuki Module)

---

## 1. Overview & System Scope

* **Target System**: ASTI Order Data Extraction System (受注データ抽出システム)
* **Project Names**:
  * **Stage 1 (Pre-Import)**: `SUZUKI_SPIRITS_DATA_PRE_IMPORT`
  * **Stage 2 (Order Registration)**: `SUZUKI_SPIRITS_DATA_IMPORT`
* **Target Database**: Oracle Database
* **Scope**: Dedicated Suzuki (`CUSTOMER_CODE = '5455'`) processing workflow. Non-Suzuki workflows (e.g., Yamaha special processing) are excluded.
* **Execution Modes**:
  * **Batch Execution**: Automated process scheduled via Windows Task Scheduler executing as a console application under the OS identity `SUZUKI-AT`.
  * **Web UI Execution**: Manual process triggered by operators via `gvSelectCustomers` on the management screen (`OrderImport.aspx`).

---

## 2. Execution Modes: Batch vs. Web UI Comparison

| Dimension | Batch Mode (`SUZUKI-AT`) | Web UI Mode (Browser) |
| :--- | :--- | :--- |
| **Trigger Mechanism** | Windows Task Scheduler | `OrderImport.aspx` (`btnImportFile`, `btnSaveOrder`) |
| **Executing User ID** | Hardcoded `'SUZUKI-AT'` | Current logged-in user ID |
| **Target Parameter Retrieval** | Queries `CUSTOMER_SETTING_MST` (`CUSTOMER_CODE = '5455'`, blank PC/Unit ID) and `FOLDER_MST` (`FOLDER_TYPE = 4`) | Selected row parameters in `gvSelectCustomers` |
| **Stage 1 & Stage 2 Transition** | Automatic immediate pipeline transition | Halts after Stage 1; operator reviews data before executing Stage 2 |
| **Reconcile Flags** | Defaults `RECONCILE_FLAG = 'Y'`, `FCST_RECONCILE_FLAG = 'Y'` | Reads screen dropdowns (`ddlReconcileFlag`, `ddlFcstReconcileFlag`) |
| **Error Handling / Output** | Exports Excel report to `[FOLDER_PATH]\エラーリスト\` | Displays error messages directly on the Web UI |
| **Post-Import Trigger** | Invokes `STRA納期設定` (STRA Delivery Date Setting) | Halts after log updates; skips downstream trigger |

---

## 3. End-to-End Workflow & Step-by-Step Execution Sequence

### Phase 1: Data Pre-Import Stage (`SUZUKI_SPIRITS_DATA_PRE_IMPORT`)

#### Step 1: Parameter & Value Retrieval (値取得)
* **Batch Mode**:
  * Queries `CUSTOMER_SETTING_ID` from `CUSTOMER_SETTING_MST` where `CUSTOMER_CODE = '5455'`, `PROFIT_CENTER IS NULL`, and `CUSTOMER_UNIT_ID IS NULL`.
  * Queries `FOLDER_PATH` from `FOLDER_MST` where `FOLDER_TYPE = 4` (混在).
* **Web UI Mode**:
  * Reads `CUSTOMER_SETTING_ID`, `CUSTOMER_CODE`, `PROFIT_CENTER`, `CUSTOMER_UNIT_ID`, `FOLDER_TYPE`, and `FOLDER_PATH` from selected rows in `gvSelectCustomers`.

#### Step 2: Source File Discovery (ファイル検索)
* Checks physical existence of `FOLDER_PATH`. If the folder does not exist, terminates immediately.
* Checks if files exist in `FOLDER_PATH`. If no files exist, execution terminates normally without error.

#### Step 3: Work Directory Management (フォルダ作成 / ファイル移動)
* Invokes `Utils.EnsureDirectory` to confirm or create the user-specific WORK folder (batch uses folder path for user `SUZUKI-AT`).
* Moves discovered files from `FOLDER_PATH` into the user WORK directory.

#### Step 4: Import Run Tracking Initialization (実行管理 / INSERT IMP_RUN)
* Queries database server timestamp: `SELECT SYSDATE FROM DUAL`.
* Inserts a record into `IMP_RUN`:
  * `STARTED_AT`: `SYSDATE`
  * `STATUS`: `'RUNNING'`
  * `STARTED_USER_ID`: `'SUZUKI-AT'` (Batch) or `[Login User ID]` (Web UI)
  * `STARTED_PG_ID`: `'OrderImport(Execute)'`

#### Step 5: Work File Staging Registration (INSERT IMP_FILES_STAGE)
* Inserts discovered files into `IMP_FILES_STAGE`:
  * `STATUS`: `'DISCOVERED'`
  * `RECONCILE_FLAG`: `'Y'` (Batch) or value from `ddlReconcileFlag` (Web UI)
  * `FCST_RECONCILE_FLAG`: `'Y'` (Batch) or value from `ddlFcstReconcileFlag` (Web UI)
  * `HAND_FLAG`: `'N'`
  * Audit columns: `CREATED_AT` (OS timestamp), `CREATED_USER_ID`, `CREATED_PG_ID = 'OrderImport(Stage)'`

#### Step 6: Specialized Suzuki Staging Insertion (INSERT スズキ専用テーブル)
* Inspects Column 1 (`INFO_TYPE_CODE`) of each line in the text file to route data to the corresponding staging table:
  * `0501`, `0502` -> `SUZUKI_SPIRITS_0501AND0502`
  * `0600`, `0630` -> `SUZUKI_SPIRITS_0600AND0630`
  * `0602` -> `SUZUKI_SPIRITS_0602`
  * `0650` -> `SUZUKI_SPIRITS_0650`
  * `0651` -> `SUZUKI_SPIRITS_0651`
  * `0740` -> `SUZUKI_SPIRITS_0740`
  * `0813` -> `SUZUKI_SPIRITS_0813` *(Excluded / Deprecated)*
  * `0814` -> `SUZUKI_SPIRITS_0814`
  * `6604`, `6634` -> `SUZUKI_SPIRITS_6604AND6634`
  * `6624` -> `SUZUKI_SPIRITS_6624`
  * `663N`, `663S`, `664T` -> `SUZUKI_SPIRITS_663NAND663SAND664T`
* Appends metadata: `IMP_FILE_ID` (`IMP_FILE_STAGE_ID`), `IMP_RUN_ID`, `CREATED_AT` (OS timestamp), `CREATED_USER_ID`, `CREATED_PG_ID = 'OrderImport(Stage)'`.
* If `INFO_TYPE_CODE` is empty or not 4 digits, increments `ERROR_COUNT` and advances to the next file.

#### Step 7: Existing Record Invalidation (UPDATE 無効 / ACTIVE_FLAG = 'N')
* Checks for prior active records (`ACTIVE_FLAG = 'Y'`) with matching keys:
  * **Standard Tables**: Matches `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, and `PUBLICATION_TIME`.
  * **Exceptions without PUBLICATION_TIME** (`SUZUKI_SPIRITS_6624`, `SUZUKI_SPIRITS_6604AND6634`, `SUZUKI_SPIRITS_663NAND663SAND664T`): Matches `CUSTOMER_ITEM_NO` and `PUBLICATION_DATE` only.
  * **Exception for 0740** (`SUZUKI_SPIRITS_0740`): Matches `CUSTOMER_ITEM_NO`, `ACCEPTANCE_DATE` (検収日), and `ACCEPTANCE_TIME` (検収時刻).
* If matching records with an older `CREATED_AT` exist, updates their `ACTIVE_FLAG` to `'N'`.

#### Step 8: ASTI Part Number Lookup (UPDATE 品番)
* Matches `CUSTOMER_ITEM_NO` against `PRDSLSODRM` to assign ASTI item number `ITEM_NO` (`FPRDCD`):
  ```sql
  WHERE PRDSLSODRM.FCUSTCD = '5455'
    AND PRDSLSODRM.FCUSTITEMNO = [CUSTOMER_ITEM_NO]
  ```
* **"Z" Suffix Rule**: If `INFO_TYPE_CODE` is `6604`, `6624`, `6634`, `663N`, `663S`, or `664T`, matches `FCUSTITEMNO` ending with `'Z'`. For all other codes, matches exact part numbers without `'Z'`.

#### Step 9: Pre-Import Error Handling & Purge
* Verifies if any newly inserted rows have `ITEM_NO IS NULL`.
* If unmapped records exist:
  * Generates error: `"Row(n):品目No及び製品コードが取得できません。"`.
  * **Batch**: Exports error list to Excel file at `[FOLDER_PATH]\エラーリスト\` (creates folder if missing).
  * **Web UI**: Displays error directly on screen.
  * Copies source file back to source folder renamed as: `[FileName]_[UserID]_[yyyyMMddhhmmss]`.
  * Physically executes `DELETE` on unmapped rows from the Suzuki staging tables.

#### Step 10: Transition Branch
* **Batch**: Automatically continues to Phase 2 (`SUZUKI_SPIRITS_DATA_IMPORT`).
* **Web UI**: Halts execution, allowing user review before manual triggering of Phase 2.

---

### Phase 2: Order Registration Stage (`SUZUKI_SPIRITS_DATA_IMPORT`)

#### Step 11: Staging Cleanup (受注ワーク削除)
* Physically deletes records from `ORDERS_STAGE` where `CUSTOMER_SETTING_ID = [Current ID]` and `FOLDER_TYPE = [Current Folder Type]`.

#### Step 12: Target Filter & Base Insertion (受注ワーク挿入)
* Filters valid order information types: `0600`, `0602`, `0630`, `0650`, `0651`, `0814`, `6604`, `6624`, `663N`, `663S`.
* **Exclusion**: Codes `0630`, `663N`, and `663S` are not inserted as base rows in `ORDERS_STAGE`; they act as update modifiers in Step 14.
* Base field mappings from Suzuki staging tables to `ORDERS_STAGE`:
  * `CUSTOMER_SETTING_ID`: `IMP_FILES_STAGE.CUSTOMER_SETTING_ID`
  * `CUSTOMER_CODE`: `CUSTOMER_SETTING_MST.CUSTOMER_CODE`
  * `BILLING_TO`: `SECTM.FBILLTOCD` (where `SECTD.FSECTTYP = 'CU'`)
  * `CUSTOMER_ORDER_NO`: `[客先発注No]`
  * `DEMAND_STATUS`: `'F'` (for `0600`, `0602`, `6604`); `'O'` (for `0650`, `0651`, `0814`, `6624`)
  * `SHIP_TO`: `[出荷先]`
  * `ORDER_DATE`: `[受注日]`
  * `DUE_DATE`: `SELECT ADD_WORKING_DAYS2('00001', [希望納期], 0) FROM DUAL`
  * `CUSTOMER_ITEM_NO`: `[客先品目No]`
  * `ITEM_NO` / `PRODUCT_CODE`: `[品目No]`
  * `DEMAND_QTY` / `PRE_DAILY_ORDER_QTY`: `[需要数]`
  * `PRE_DAILY_DELIVERY_DATE`: `[希望納期]` (`yyyy/MM/dd`)
  * `DEMAND_UNIT`: `ITEMM.FUNIT` (where `ITEMM.FITEMNO = PRDSLSODRM.FPRDCD`)
  * `CURRENCY_CODE`: `SECTM.FCURR` (where `SECTM.FSECTCD = CUSTOMER_SETTING_MST.CUSTOMER_CODE`)
  * `SHIP_STOCK_LOCATION`: `ITEMM.FSHPWHCD` (where `ITEMM.FITEMNO = ORDERS_STAGE.ITEM_NO`)
  * `COMPANY_ID`: `'1000'`
  * `BILLING_STANDARD`: `'S'`
  * `SHIP_PROCESS_TYPE`: `'O'` (for `0600`, `0602`, `6604`); `'K'` (for `0650`, `0651`, `0814`, `6624`)
  * `DELIVERY_INSTR_FLAG`: `'N'` (for `0600`, `0602`, `6604`); `'Y'` (for `0650`, `0651`, `0814`, `6624`)
  * `TOTAL_SHIP_QTY`: `NULL` (for `0600`, `0602`, `6604`); `0` (for `0650`, `0651`, `0814`, `6624`)
  * `TRANSPORT_METHOD`: `2`
  * `ORDER_TYPE`: `1` (内示 for `0600`, `0602`, `6604`); `3` (納入指示 for `0650`, `0651`, `0814`, `6624`)
  * `STATUS`: `'IMPORTED'`
  * `ACTIVE_FLAG`: `'Y'`
  * Metadata: `IMP_RUN_ID`, `CREATED_AT = STARTED_AT`, `CREATED_USER_ID`, `CREATED_PG_ID = 'OrderImport(Execute)'`

#### Step 13: Profit Center Dynamic Setting (受注ワーク更新: CUSTOMER_SETTING_ID)
* Retrieves Profit Center (`PC`):
  ```sql
  SELECT FUSRSTR1 FROM USRDEFFLDF 
  WHERE FTABLEID = 'ITEMM' AND FRECKEY = ORDERS_STAGE.PRODUCT_CODE;
  ```
* Updates `ORDERS_STAGE.CUSTOMER_SETTING_ID`:
  ```sql
  SELECT CUSTOMER_SETTING_ID FROM CUSTOMER_SETTING_MST 
  WHERE CUSTOMER_CODE = '5455' AND PROFIT_CENTER = [取得したPC];
  ```
* If no matching Profit Center is returned, leaves `CUSTOMER_SETTING_ID` unadjusted.

#### Step 14: Special Demand & Delivery Adjustments (受注ワーク更新: 0630 / 663N / 663S)
* **Code 0630 (Quantity Modifier)**:
  * Updates `ORDERS_STAGE.DEMAND_QTY` where `ITEM_NO`, `PRE_DAILY_DELIVERY_DATE`, `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, and `INFO_TYPE_CODE = '0600'` match:
    ```
    ORDERS_STAGE.DEMAND_QTY = ORDERS_STAGE.DEMAND_QTY + SUZUKI_SPIRITS_0600AND0630.DEMAND_QTY
    ```
  * Fallback: If not found in `ORDERS_STAGE`, retrieves from `ORDERS` master and reflects to `ORDERS_STAGE`.
* **Code 663N (Due Date Modifier)**:
  * Updates `ORDERS_STAGE.DUE_DATE` matching `CUSTOMER_ORDER_NO` and `INFO_TYPE_CODE = '6624'`.
  * Fallback: If not found in `ORDERS_STAGE`, retrieves from `ORDERS` master and reflects to `ORDERS_STAGE`.
* **Code 663S (Quantity Modifier)**:
  * Updates `ORDERS_STAGE.DEMAND_QTY` matching `CUSTOMER_ORDER_NO` and `INFO_TYPE_CODE = '6624'`.
  * Fallback: If not found in `ORDERS_STAGE`, retrieves from `ORDERS` master and reflects to `ORDERS_STAGE`.

#### Step 15: Shipping Location Mapping & Forecast Adjustments (内示加工)
* Maps ship-to code via `SUZUKI_SPIRITS_CONVERSION_MST` where `DELIVERY_CODE_ORDER = ORDERS_STAGE.SHIP_TO` and `ACTIVE_FLAG = 'Y'`:
  * If found, sets `ORDERS_STAGE.SHIP_TO = DELIVERY_CODE_PLAN`.
  * If not found, retains original `ORDERS_STAGE.SHIP_TO`.
* **Rule for 0602**: Does not contain reference dates (`TARGET_REFERENCE_DATE`); treated as independent new demand with no forecast invalidation.

#### Step 16: Reconciliation & Offset Processing (消込処理)
* **Firm Order Offset (確定消込)**:
  * When `RECONCILE_FLAG = 'Y'`, matches incoming delivery instructions (`ORDER_TYPE = 3`) against firm orders in `ORDERS_STAGE` (`ORDER_TYPE = 2`) on `CUSTOMER_SETTING_ID` and `CUSTOMER_ORDER_NO`:
    ```sql
    MERGE INTO orders_stage tgt
    USING (
      SELECT p.rowid AS rid, p.demand_qty AS past_qty, c.demand_qty AS cur_qty
      FROM past p JOIN cur c
        ON c.customer_setting_id = p.customer_setting_id
       AND c.customer_order_no = p.customer_order_no
    ) m ON (tgt.rowid = m.rid)
    WHEN MATCHED THEN
      UPDATE SET
        tgt.demand_qty = NVL(m.past_qty, 0) - NVL(m.cur_qty, 0),
        tgt.status = CASE WHEN NVL(m.past_qty, 0) - NVL(m.cur_qty, 0) = 0 THEN 'RECONCILED' ELSE tgt.status END,
        tgt.active_flag = CASE WHEN NVL(m.past_qty, 0) - NVL(m.cur_qty, 0) = 0 THEN 'N' ELSE tgt.active_flag END,
        tgt.updated_at = :startedAt,
        tgt.updated_user_id = :loginUserId,
        tgt.updated_pg_id = :updatedPgId;
    ```
* **Forecast Offset (内示消込)**:
  * When `FCST_RECONCILE_FLAG = 'Y'`, evaluates `RECONCILE_TYPE` (`1`: 順次, `2`: 同月まで, `3`: 同月のみ) from `IMP_RULE_MST`:
    * Reduces quantity of active forecast records (`ORDER_TYPE = 1`).
    * If remaining `DEMAND_QTY = 0`: sets `STATUS = 'RECONCILED'`, `ACTIVE_FLAG = 'N'`.
    * If remaining `DEMAND_QTY > 0`: retains `STATUS`, keeps `ACTIVE_FLAG = 'Y'`.
* **Delivery Instruction Update**:
  * Sets processed delivery instruction records to `STATUS = 'PROCESSED'`, `ACTIVE_FLAG = 'Y'`.

#### Step 17: Production Order Registration (正規データ追加 / 更新)
* **Commit to Production Tables**:
  * Updates existing rows in `ORDERS` matching `ORDERS_ID`:
    * Fields: `CUSTOMER_ORDER_NO`, `DUE_DATE`, `CUSTOMER_ITEM_NO`, `DEMAND_QTY`, `STATUS`, `ACTIVE_FLAG`, `UPDATED_AT`, `UPDATED_USER_ID`, `UPDATED_PG_ID = 'OrderImport(Save)'`.
  * If `ORDERS_ID IS NULL`, inserts new record into `ORDERS` using current `IMP_FILE_STAGE_ID`.
  * Inserts historical audit record into `ORDER_HISTORY`.
* **Staging Status Update**:
  * Checks `IMP_FILES.HAND_FLAG = 'N'` and `CUSTOMER_CODE = '5455'`.
  * Updates Suzuki staging tables: sets `IMP_FILE_ID` and `STATUS = 'REGISTERED'`.
* **Staging Cleanup**:
  * Deletes processed records from `ORDERS_STAGE` where `CUSTOMER_SETTING_ID = [Current ID]`, `STATUS = 'PROCESSED'`, and user matches `PROD_MGMT_USER_ID` or admin.

#### Step 18: File Archival & Downstream Dispatch
* Moves processed physical files from user WORK folder to completed archive folder:
  * Naming pattern: `[OriginalFileName]_[LoginUserID]_[yyyyMMddhhmmss].[ext]`.
* Inserts row into `IMP_FILES` (`STATUS = 'PARSED'`) and deletes from `IMP_FILES_STAGE`.
* Updates `IMP_RUN`:
  * `ENDED_AT`: `SYSDATE`
  * `STATUS`: `'COMPLETED'`
  * `FILE_COUNT`: Total processed files
  * `ROW_COUNT`: Total processed records
  * `ERROR_COUNT`: Total errors
* **Batch Downstream Trigger**: Invokes `STRA納期設定` (STRA Delivery Date Setting). Web UI execution terminates without invoking this step.

---

## 4. Complete Suzuki Table & Code Mapping Matrix

| INFO_TYPE_CODE | Japanese Description | Target Suzuki Staging Table | Target Production Table | Register to Master (`ORDERS`)? | ASTI Part Number Z-Suffix Lookup? |
| :---: | :--- | :--- | :--- | :---: | :---: |
| **0501** | 内示データ（確定） | `SUZUKI_SPIRITS_0501AND0502` | — | No | No (Exact match) |
| **0502** | 内示データ（内示） | `SUZUKI_SPIRITS_0501AND0502` | — | No | No (Exact match) |
| **0600** | 確定受注データ（月次内示） | `SUZUKI_SPIRITS_0600AND0630` | `ORDERS` | **Yes** (ORDER_TYPE: 1) | No (Exact match) |
| **0602** | 確定受注データ（追加内示） | `SUZUKI_SPIRITS_0602` | `ORDERS` | **Yes** (ORDER_TYPE: 1) | No (Exact match) |
| **0630** | 確定受注データ（内示増減） | `SUZUKI_SPIRITS_0600AND0630` | `ORDERS` (Updates 0600) | **Yes** (Modifier) | No (Exact match) |
| **0650** | 納入指示データ（通常） | `SUZUKI_SPIRITS_0650` | `ORDERS` | **Yes** (ORDER_TYPE: 3) | No (Exact match) |
| **0651** | 納入指示データ（臨時） | `SUZUKI_SPIRITS_0651` | `ORDERS` | **Yes** (ORDER_TYPE: 3) | No (Exact match) |
| **0740** | 検収データ | `SUZUKI_SPIRITS_0740` | — | No | No (Exact match) |
| **0813** | 支給品情報データ（内示） | `SUZUKI_SPIRITS_0813` | — | **No** (Excluded/Deprecated) | No (Exact match) |
| **0814** | 支給品納入指示データ（確定） | `SUZUKI_SPIRITS_0814` | `ORDERS` | **Yes** (ORDER_TYPE: 3) | No (Exact match) |
| **6604** | 注文データ（内示） | `SUZUKI_SPIRITS_6604AND6634` | `ORDERS` | **Yes** (ORDER_TYPE: 1) | **Yes (Ends with 'Z')** |
| **6624** | 注文データ（確定/指示） | `SUZUKI_SPIRITS_6624` | `ORDERS` | **Yes** (ORDER_TYPE: 3) | **Yes (Ends with 'Z')** |
| **6634** | 注文データ（変更） | `SUZUKI_SPIRITS_6604AND6634` | — | No | **Yes (Ends with 'Z')** |
| **663N** | 注文データ（納期変更） | `SUZUKI_SPIRITS_663NAND663SAND664T` | `ORDERS` (Updates 6624) | **Yes** (Modifier) | **Yes (Ends with 'Z')** |
| **663S** | 注文データ（数量変更） | `SUZUKI_SPIRITS_663NAND663SAND664T` | `ORDERS` (Updates 6624) | **Yes** (Modifier) | **Yes (Ends with 'Z')** |
| **664T** | 注文データ（打切/取消） | `SUZUKI_SPIRITS_663NAND663SAND664T` | — | No | **Yes (Ends with 'Z')** |

---

## 5. EDI Column Transformations & Business Conversion Rules

### Conversion Rule A: Information Type (`CUSTOMER_INFO_TYPE`)
* **Codes 0600 & 6604**:
  * `ORDER_DATA_TYPE = '1'` (内示) -> `'新規'`
  * `ORDER_DATA_TYPE = '2'` (確定) -> `'新規'`
  * `ORDER_DATA_TYPE = '3'` (変更) -> `'変更'`
* **Code 0602**:
  * `ORDER_DATA_TYPE = '2'` (確定) -> `'新規'`
  * `ORDER_DATA_TYPE = '3'` (変更) -> `'変更'`
* **Codes 0651, 6634, 663N, 663S**:
  * Evaluates position 17 of `ORDER_NOTES` (発行注釈). If position 17 is non-zero (`<> '0'`), sets `CUSTOMER_INFO_TYPE = '打切'`.

### Conversion Rule B: Prorated Type (`PRORATED_TYPE` / 分割区分)
* **Mapped from `DELIVERY_DATE_TYPE` (納入日区分)**:
  * Value `'1'` (年月 / 大日程) -> `'2'` (日割以外)
  * Value `'4'` (年週) -> `'2'` (日割以外)
  * Value `'5'` (年月日) -> `'1'` (日割り)
  * Value `'6'` (年月日時分) -> `'1'` (日割り)

### Conversion Rule C: Date & Delivery Timing
* All incoming dates (`ORDER_DATE`, `DUE_DATE`, `TARGET_REFERENCE_DATE`) must follow format `yyyy/MM/dd`.
* `ORDERS_STAGE.DUE_DATE` calculates the ASTI factory calendar using:
  `ADD_WORKING_DAYS2('00001', [希望納期], 0)`.
* `PRE_DAILY_DELIVERY_DATE` and `PRE_DAILY_ORDER_QTY` preserve raw unadjusted pre-prorated incoming values.

---

## 6. Antigravity IDE Verification Checklist

1. [ ] **Identity & Role**: Does batch mode use `'SUZUKI-AT'` as `CREATED_USER_ID` and `STARTED_USER_ID`?
2. [ ] **Default Flags**: Are `RECONCILE_FLAG` and `FCST_RECONCILE_FLAG` hardcoded to `'Y'` in batch mode, while UI reads dropdown controls?
3. [ ] **Part Number Matching**: Does the lookup in `PRDSLSODRM` filter for parts ending in `'Z'` exclusively for codes `6604, 6624, 6634, 663N, 663S, 664T`?
4. [ ] **Step 7 Invalidation Keys**: Does Step 7 skip `PUBLICATION_TIME` on tables `SUZUKI_SPIRITS_6624`, `SUZUKI_SPIRITS_6604AND6634`, and `SUZUKI_SPIRITS_663NAND663SAND664T`, and match `ACCEPTANCE_DATE`/`ACCEPTANCE_TIME` on `0740`?
5. [ ] **Stage 2 Modifiers**: Are codes `0630`, `663N`, and `663S` excluded from Step 12 initial insertion and executed strictly as quantity/due-date updates in Step 14?
6. [ ] **Profit Center Assignment**: Does Step 13 look up `FUSRSTR1` from `USRDEFFLDF` (`FTABLEID = 'ITEMM'`) and set `CUSTOMER_SETTING_ID` based on `PROFIT_CENTER`?
7. [ ] **Forecast Reconciliation**: Does Step 16 execute sequential/monthly forecast offsets and update `STATUS = 'RECONCILED'` and `ACTIVE_FLAG = 'N'` when remaining demand equals zero?
8. [ ] **Downstream Dispatch**: Does batch mode trigger `STRA納期設定` upon completing `IMP_RUN`, while Web UI bypasses it?