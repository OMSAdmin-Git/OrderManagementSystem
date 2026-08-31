# 📄 SPIRITS Data Import Process Specification Document

## 🏗️ 1. Overview & System Scope

* **Project Names**: 
  * `SUZUKI_SPIRITS_DATA_PRE_IMPORT` (Stage 1: Pre-Import & Staging)
  * `SUZUKI_SPIRITS_DATA_IMPORT` (Stage 2: Order Validation & Registration)
* **Module Names**: SPIRITS Data Import & Order Registration (`SPIRITSデータ取込` & `SPIRITSデータ受注登録`)
* **Purpose**: Automatically extract text/EDI files downloaded from Suzuki's SPIRITS system, insert them into dedicated Suzuki staging tables, cross-reference external customer item numbers with internal ASTI part numbers via `PRDSLSODRM`, validate demand parameters, and commit valid orders into production tables (`ORDERS`, `ORDER_HISTORY`).
* **Target Database System**: Oracle Database

---

## 🎛️ 1.1 Execution Modes & Runtime Behavioral Matrix

The system can be triggered in two primary environments, which exhibit key behavioral differences across parameter retrieval, user contextual auditing, execution scope, error reporting, and post-processing steps:

| Execution Feature / Aspect | 🤖 Task Scheduler (Batch Mode) | 💻 Web Screen (Web UI Mode) |
| :--- | :--- | :--- |
| **Trigger Mechanism** | Scheduled Windows Task / Console Application | User button click on Web Interface |
| **OS / User Identity (`CREATED_USER_ID`)** | Hardcoded OS user `'SUZUKI-AT'` | Authenticated Login User ID (e.g., `EMP1040`) |
| **Parameter Lookup Source** | `CUSTOMER_SETTING_MST` (where `CUSTOMER_CODE = '5455'` and `PROFIT_CENTER` / `CUSTOMER_UNIT_ID` are `NULL`) | Selected row grid controls (`gvSelectCustomers`) |
| **Folder Type Scope** | Fixed to `FOLDER_TYPE = 4` (`4: 混在`) | User-selected `FOLDER_TYPE` |
| **Reconcile Flag Defaults** | Hardcoded `'Y'` for `RECONCILE_FLAG` & `FCST_RECONCILE_FLAG` | Read from UI drop-down choices (`ddlReconcileFlag`, `ddlFcstReconcileFlag`) |
| **Stage 1 to Stage 2 Pipeline** | Automatically seamlessly proceeds to Stage 2 (`SUZUKI_SPIRITS_DATA_IMPORT`) | Halts after Stage 1, or proceeds manually based on `chkStageImport` check box |
| **Error Reporting Destination** | Generates aggregated Excel report at `[FOLDER_PATH]\エラーリスト\` | Directly renders validation errors on the browser screen (`lblError.Text`) |
| **Downstream Process Integration** | Automatically executes `STRA納期設定` (STRA Delivery Date Setting) on completion | Bypasses `STRA納期設定` process entirely |

---

## 🔄 2. Workflow & Step-by-Step Execution Sequence

The complete process is executed across 18 distinct processing steps divided into **Data Pre-Import Stage** (`SUZUKI_SPIRITS_DATA_PRE_IMPORT`) and **Order Registration Stage** (`SUZUKI_SPIRITS_DATA_IMPORT`):

### Step 1: Parameter & Value Retrieval (`値取得`)
* **Batch Mode**:
  * Retrieves `CUSTOMER_SETTING_ID` from `CUSTOMER_SETTING_MST` where `CUSTOMER_CODE = '5455'` and both `PROFIT_CENTER` and `CUSTOMER_UNIT_ID` are `NULL`/blank.
  * Fetches `FOLDER_TYPE = 4` (`4: 混在`) and `FOLDER_PATH` from `FOLDER_MST`.
* **UI Mode**:
  * Retrieves `CUSTOMER_SETTING_ID`, `CUSTOMER_CODE`, `PC`, `CUSTOMER_UNIT_ID`, `FOLDER_TYPE`, and `FOLDER_PATH` directly from UI grid controls (`gvSelectCustomers`).

### Step 2: Source File Discovery (`ファイル検索`)
* Validates whether `FOLDER_PATH` exists on the file system; if missing, terminates execution.
* Searches for import target files inside `FOLDER_PATH`. If no files exist, execution terminates normally.

### Step 3: Work Directory Management (`フォルダ作成` / `ファイル移動`)
* Calls helper `Utils.EnsureDirectory` to verify/create user-specific WORK directories: `C:\ASTI\DATA\WORK\[UserId]\[CustomerCode]\[FolderType]\` (uses directory `SUZUKI-AT\5455\4` for batch execution).
* Moves discovered files from `FOLDER_PATH` into the user-specific WORK directory.

### Step 4: Import Run Tracking Initialization (`実行管理` / `INSERT IMP_RUN`)
* Creates an execution tracking log entry in `IMP_RUN`:
  * `STARTED_AT = SYSDATE`
  * `STATUS = 'RUNNING'`
  * `STARTED_USER_ID = [Login User ID]` (or `'SUZUKI-AT'`)
  * `STARTED_PG_ID = 'OrderImport(Execute)'`

### Step 5: Work File Staging Registration (`INSERT IMP_FILES_STAGE`)
* Registers each file into `IMP_FILES_STAGE`:
  * `STATUS = 'DISCOVERED'`
  * `STAGED_FOLDER_PATH`: `C:\ASTI\DATA\WORK\[UserId]\[CustomerCode]\[FolderType]\`
  * `STAGED_FILE_NAME`: `[FileName]`
  * `RECONCILE_FLAG` & `FCST_RECONCILE_FLAG`: Explicitly set to `'Y'` during batch execution, or populated from UI selection controls (`ddlReconcileFlag`, `ddlFcstReconcileFlag`) during Web UI execution.
  * `HAND_FLAG`: Hardcoded default `'N'`.
  * Audit fields: `CREATED_AT = OS日時`, `CREATED_USER_ID`, `CREATED_PG_ID = 'OrderImport(Stage)'`.

### Step 6: Specialized Suzuki Staging Insert (`INSERT (スズキ専用テーブル)`)
* Inspects Column 1 (`INFO_TYPE_CODE`) of each text row to direct data into dedicated Suzuki staging tables:
  * `0501`, `0502` $\rightarrow$ `SUZUKI_SPIRITS_0501AND0502`
  * `0600`, `0630` $\rightarrow$ `SUZUKI_SPIRITS_0600AND0630`
  * `0602` $\rightarrow$ `SUZUKI_SPIRITS_0602`
  * `0650` $\rightarrow$ `SUZUKI_SPIRITS_0650`
  * `0651` $\rightarrow$ `SUZUKI_SPIRITS_0651`
  * `0740` $\rightarrow$ `SUZUKI_SPIRITS_0740`
  * `0813` $\rightarrow$ `SUZUKI_SPIRITS_0813` *(Note: Table physically dropped/excluded from import scope per 2026/8/6 system revision)*
  * `0814` $\rightarrow$ `SUZUKI_SPIRITS_0814`
  * `6604`, `6634` $\rightarrow$ `SUZUKI_SPIRITS_6604AND6634`
  * `6624` $\rightarrow$ `SUZUKI_SPIRITS_6624`
  * `663N`, `663S`, `664T` $\rightarrow$ `SUZUKI_SPIRITS_663NAND663SAND664T`
* Attaches runtime lineage metadata: `IMP_FILE_ID` (`IMP_FILE_STAGE_ID`), `IMP_RUN_ID`, `CREATED_AT = OS日時`, `CREATED_USER_ID`, and `CREATED_PG_ID = 'OrderImport(Stage)'`.
* If `INFO_TYPE_CODE` is invalid or not 4 digits, increments `ERROR_COUNT` and skips to the next file.

### Step 7: Existing Record Invalidation (`UPDATE（無効）/ ACTIVE_FLAG = 'N'`)
* Checks target Suzuki staging tables for prior active records (`ACTIVE_FLAG = 'Y'`) matching `CUSTOMER_ITEM_NO` and publication timestamps.
* **Key Mechanics & Exceptions**:
  * Tables `SUZUKI_SPIRITS_6624`, `SUZUKI_SPIRITS_6604AND6634`, and `SUZUKI_SPIRITS_663NAND663SAND664T` lack `PUBLICATION_TIME`; invalidation checks match `PUBLICATION_DATE` + `CUSTOMER_ITEM_NO` only.
  * Table `SUZUKI_SPIRITS_0740` lacks publication fields entirely; invalidation relies on `ACCEPTANCE_DATE` (`検収日`) + `ACCEPTANCE_TIME` (`検収時刻`) + `CUSTOMER_ITEM_NO`.
* Sets matched prior records to `ACTIVE_FLAG = 'N'` to preserve historical data lineage.

### Step 8: Part Number Matching & Suffix Business Logic (`UPDATE（品番）`)
* Maps `CUSTOMER_ITEM_NO` against `PRDSLSODRM` where `FCUSTCD = '5455'` and `FCUSTITEMNO = CUSTOMER_ITEM_NO` to obtain `ITEM_NO` (`FPRDCD`).
* **"Z" Suffix Rules**: If `INFO_TYPE_CODE` is `6604`, `6624`, `6634`, `663N`, `663S`, or `664T`, matches against `FCUSTITEMNO` ending with **'Z'**. All other codes match exact part numbers without 'Z'.

### Step 9: Pre-Import Error Handling & File Recovery (`エラーチェック` / `ファイルコピー` / `DELETE` / `ファイル移動`)
* Scans inserted staging rows for missing `ITEM_NO` values.
* If unmapped rows exist:
  * Outputs error message: `"Row(n):品目No及び製品コードが取得できません。"`.
  * **Batch Mode**: Generates Excel error report at `[FOLDER_PATH]\エラーリスト\`.
  * **UI Mode**: Renders error directly on the browser screen in `lblError.Text` (modal popup code is preserved in comments for future use).
  * Copies source file back to original location as `[OriginalFileName]_[UserID]_[ProcessingTimestamp]`.
  * Physically deletes (`DELETE`) the unmapped rows from Suzuki staging tables.
* If **NO errors exist**:
  * Successfully parsed and staged files are moved immediately from the `WORK` directory to the final `COMPLETED` directory (`CompletedFolderRoot`). This early cleanup is possible because Stage 2 will read from the `SUZUKI_SPIRITS_...` database tables rather than the physical files.

### Step 10: Transition to Order Registration (`受注登録`)
* Batch execution automatically moves to Order Registration Stage (`SUZUKI_SPIRITS_DATA_IMPORT`). Manual UI execution halts after staging or proceeds based on user triggering.

### Step 11: Stage Cleaning (`受注ワーク削除`)
* Physically purges `ORDERS_STAGE` where `CUSTOMER_SETTING_ID = Current Target` and `FOLDER_TYPE = Current Target`.

### Step 12: Order Data Verification (`受注データチェック`)
* Filters Suzuki staging records for valid order `INFO_TYPE_CODE` values: `0600`, `0602`, `0630`, `0650`, `0651`, `0814`, `6604`, `6624`, `663N`, `663S`.

### Step 13: Order Stage Staging Insert (`INSERT ORDERS_STAGE`)
* Inserts filtered records into `ORDERS_STAGE` (skips `0630`, `663N`, `663S` during initial insert; updated in Step 14):
  * `COMPANY_ID`: Fixed value `'1000'`
  * `BILLING_STANDARD`: Fixed value `'S'`
  * `TRANSPORT_METHOD`: Fixed value `2`
  * `ORDER_TYPE`: Set to `1` (`1:内示`) for codes `0600`, `0602`, `0630`, `6604`, `663N`, `663S`; set to `3` (`3:納入指示`) for `0650`, `0651`, `0814`, `6624`.
  * `DEMAND_STATUS`: `'F'` (when `ORDER_TYPE = 1`), `'O'` (otherwise).
  * `SHIP_PROCESS_TYPE`: `'O'` (when `ORDER_TYPE = 1`), `'E'` (when `ORDER_TYPE = 2`), `'K'` (when `ORDER_TYPE = 3`).
  * `DELIVERY_INSTR_FLAG`: `'N'` (when `ORDER_TYPE ≠ 3`), `'Y'` (when `ORDER_TYPE = 3`).
  * `DUE_DATE = ADD_WORKING_DAYS2('00001', [DUE_DATE], 0)` formatted as `yyyy/MM/dd`.
  * `DEMAND_UNIT`: Joined from `ITEMM.FUNIT` where `ITEMM.FITEMNO = PRDSLSODRM.FPRDCD`.
  * `SHIP_STOCK_LOCATION`: Joined from `ITEMM.FSHPWHCD` where `ITEMM.FITEMNO = ORDERS_STAGE.ITEM_NO`.
  * **Phase 2 Extended Attributes** (Added per system revisions): `DELIVERY_TIME`, `CONTAINER_CAPACITY`, `INITIAL_FLAG`, `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE`.

### Step 14: Special Demand & Delivery Adjustments (`受注ワーク更新`)
* Updates `CUSTOMER_SETTING_ID` in `ORDERS_STAGE` by fetching Profit Center (`FUSRSTR1`) from `USRDEFFLDF` where `FTABLEID='ITEMM'` and `FRECKEY = PRODUCT_CODE`, then joining `CUSTOMER_SETTING_MST` on `CUSTOMER_CODE = '5455'` and `PROFIT_CENTER`.
* Code-specific adjustments:
  * **Code `0630`**: `DEMAND_QTY = ORDERS_STAGE.DEMAND_QTY + SUZUKI_SPIRITS_0600AND0630.DEMAND_QTY` (matching `ITEM_NO`, `PRE_DAILY_DELIVERY_DATE`, `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE = '0600'`). Fallback to `ORDERS` master table if staging record is missing.
  * **Code `663N`**: Updates `DUE_DATE` from `SUZUKI_SPIRITS_663NAND663SAND664T` (matching `CUSTOMER_ORDER_NO`, `INFO_TYPE_CODE = '6624'`). Fallback to `ORDERS` master table if missing.
  * **Code `663S`**: Updates `DEMAND_QTY` from `SUZUKI_SPIRITS_663NAND663SAND664T` (matching `CUSTOMER_ORDER_NO`, `INFO_TYPE_CODE = '6624'`). Fallback to `ORDERS` master table if missing.

### Step 15: Post-Stage Validation & File Settlement (`エラーチェック` / `IMP_FILES`)
* Performs validation. If errors exist: outputs error log and executes `DELETE FROM ORDERS_STAGE`.
* If valid:
  * Updates `IMP_FILES_STAGE` with `STAGED_FOLDER_PATH` and `STAGED_FILE_NAME`.
  * Inserts records into `IMP_FILES` with `STATUS = 'PARSED'`.
  * Deletes work record from `IMP_FILES_STAGE`.
  * *(Note: Physical files were already moved to the `Completed` directory during Stage 1).*

### Step 16: Forecast Invalidation & Conversion Processing (`内示加工` / `受注残加工` / `納入指示加工`)
* Maps destination codes using `SUZUKI_SPIRITS_CONVERSION_MST` (`DELIVERY_CODE_PLAN` $\rightarrow$ `SHIP_TO` where `DELIVERY_CODE_ORDER = SHIP_TO`).
* **Special Rule `0602`**: Contains no reference date fields; always processed as pure new data without forecast invalidation/replacement.
* **Forecast Consumption**: `0650`/`0651` consumes active `0600` forecasts; `6624` consumes active `6604` forecasts.

### Step 17: Production Order Table Commitment (`INSERT ORDERS` / `INSERT ORDER_HISTORY`)
* Updates Suzuki staging table status to `'REGISTERED'` and populates `IMP_FILE_ID`.
* Copies valid rows from `ORDERS_STAGE` into master production table `ORDERS` (including Phase 2 extended attributes: `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE`).
* Appends historical audit records into `ORDER_HISTORY` (including Phase 2 extended attributes: `TARGET_REFERENCE_DATE_TYPE`, `TARGET_REFERENCE_DATE`, `INFO_TYPE_CODE`, `DELIVERY_TIME`, `CONTAINER_CAPACITY`, `INITIAL_FLAG`).

### Step 18: Completion & Execution Logging (`UPDATE IMP_RUN` / `STRA納期設定`)
* Updates `IMP_RUN`: `ENDED_AT = SYSDATE`, `STATUS = 'COMPLETED'`, `FILE_COUNT`, `ROW_COUNT`, `ERROR_COUNT`.
* **Batch Mode Only**: Invokes downstream application `STRA納期設定` (STRA Delivery Date Setting) before finishing execution.

---

## 🗂️ 3. Data Mapping & Table Matrix

The following table details all 4-digit `INFO_TYPE_CODE` values processed by the SPIRITS Data Import System:

| INFO_TYPE_CODE | Japanese Description | Target Oracle Table Name | Update Active Flag ('N') Match Condition | Register to Master? | Z-Suffix Applicable? |
| :---: | :--- | :--- | :--- | :---: | :---: |
| **0501** | 内示データ（確定） | `SUZUKI_SPIRITS_0501AND0502` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | No | No |
| **0502** | 内示データ（内示） | `SUZUKI_SPIRITS_0501AND0502` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | No | No |
| **0600** | 確定受注データ（月次内示） | `SUZUKI_SPIRITS_0600AND0630` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **0602** | 確定受注データ（追加内示） | `SUZUKI_SPIRITS_0602` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **0630** | 確定受注データ（内示増減） | `SUZUKI_SPIRITS_0600AND0630` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **0650** | 納入指示データ（通常） | `SUZUKI_SPIRITS_0650` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **0651** | 納入指示データ（臨時） | `SUZUKI_SPIRITS_0651` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **0740** | 検収データ | `SUZUKI_SPIRITS_0740` | `CUSTOMER_ITEM_NO`, `ACCEPTANCE_DATE`, `ACCEPTANCE_TIME` | No | No |
| **0813** | 支給品情報データ | `SUZUKI_SPIRITS_0813` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | No | No |
| **0814** | 支給品納入指示データ | `SUZUKI_SPIRITS_0814` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE`, `PUBLICATION_TIME` | Yes (`ORDERS`) | No |
| **6604** | 注文データ（内示） | `SUZUKI_SPIRITS_6604AND6634` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | Yes (`ORDERS`) | **Yes** |
| **6624** | 注文データ（確定/指示） | `SUZUKI_SPIRITS_6624` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | Yes (`ORDERS`) | **Yes** |
| **6634** | 注文データ（変更） | `SUZUKI_SPIRITS_6604AND6634` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | No | **Yes** |
| **663N** | 注文データ（納期変更） | `SUZUKI_SPIRITS_663NAND663SAND664T` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | Yes (`ORDERS`) | **Yes** |
| **663S** | 注文データ（数量変更） | `SUZUKI_SPIRITS_663NAND663SAND664T` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | Yes (`ORDERS`) | **Yes** |
| **664T** | 注文データ（打切/取消） | `SUZUKI_SPIRITS_663NAND663SAND664T` | `CUSTOMER_ITEM_NO`, `PUBLICATION_DATE` *(No Time)* | No | **Yes** |

---

## 🛡️ 4. Key Technical Considerations & Exception Handling

### 1. Dual Execution UI vs. Batch Error Notification & Value Acquisition
* **Error File Generation (Both Modes)**: Errors are ALWAYS compiled into a Shift-JIS CSV file saved under `[FOLDER_MST.FOLDER_PATH]\エラーリスト\`. The directory is created automatically if missing.
* **Web UI Error Display**: When executed from the Web UI, error messages are *additionally* passed back to the caller to render directly on screen via user interface message dialogs/grids.

**Value Acquisition Rules (値取得):**
* **Batch Mode**:
  - `Customer Setting ID`: Fetched using hardcoded condition `CUSTOMER_CODE = '5455'` AND `PROFIT_CENTER` / `CUSTOMER_UNIT_ID` are blank.
  - `Folder Type`: Hardcoded to `4` (混在 / Mixed).
* **Web UI Mode**:
  - `Customer Setting ID`, `Profit Center`, `Customer Unit ID` are supplied dynamically from the Web Grid (`gvSelectCustomers`).
  - `Folder Type`: Retrieved dynamically based on the selected setting.

### 2. Strict Date & Calendar Conversion Rules
* Incoming dates (`DUE_DATE`, `PRE_DAILY_DELIVERY_DATE`) must be explicitly formatted as string `yyyy/MM/dd`.
* Calendar due dates calculate factory working days via Oracle function:
  ```sql
  SELECT ADD_WORKING_DAYS2('00001', [希望納期], 0) FROM DUAL;
  ```

* **Prorated Conversion Rules (`DELIVERY_DATE_TYPE` / 分割区分)**:
* Codes `1: 年月` / `4: 年週` $\rightarrow$ `2: 日割以外`.
* Codes `5: 年月日` / `6: 年月日時分` $\rightarrow$ `1: 日割り`.



### 3. File Cleanup, Rollback & Reversion Rules

* **Error File Handling**: Files failing validation/item mapping are returned to `FOLDER_PATH` renamed as `[OriginalName]_[UserID]_[Timestamp]`.
* **Transactional Cleanup**: Validation failures trigger physical `DELETE` queries on uncommitted staging rows in Suzuki tables and `ORDERS_STAGE` to maintain data integrity.
* **Completed File Archiving**: Successfully imported files are moved to the `COMPLETED` directory immediately during Stage 1, because all data is securely buffered in the `SUZUKI_SPIRITS_...` staging tables. In Stage 2, logical tracking is updated (recorded in `IMP_FILES` with status `'PARSED'`, and purged from `IMP_FILES_STAGE`).