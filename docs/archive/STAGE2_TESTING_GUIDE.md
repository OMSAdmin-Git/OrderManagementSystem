# 🧪 Suzuki SPIRITS Stage 2 Testing Guide & CSV Reference

This document provides a comprehensive reference for testing the Suzuki Stage 2 Order Registration and Reconciliation pipeline across Unit Tests, Batch CLI, and the Web UI.

---

## 📋 1. CSV Code Classification & Testability

### Quick Terminology
* **内示 (Naiji / Forecast - `ORDER_TYPE = 1`)**: Non-firm planning orders. Can be wash-replaced by newer forecasts or consumed by subsequent delivery instructions.
* **納入指示 (Nounyu Shiji / Delivery Instruction - `ORDER_TYPE = 3`)**: Firm shipment instructions. Converts destination and consumes matching prior forecasts.
* **特殊変更 (Special Modifier Adjustments)**: Never create independent order rows. They find an existing order and adjust its quantity or due date.

---

### Group A: 🟢 Easiest for First-Time Testing (Standalone / Direct to DB)
These files do not depend on other files to exist first.

| Code | Order Type | Staging Table | Description | Why it is easy to test |
| :--- | :--- | :--- | :--- | :--- |
| **`0602`** | `1` (内示) | `SUZUKI_SPIRITS_0602` | Supplemental / Special Forecast | **Easiest #1**: Has no target reference date; specification explicitly states it **skips wash replacement**. Moves cleanly: CSV $\rightarrow$ staging $\rightarrow$ `ORDERS_STAGE` $\rightarrow$ `ORDERS`. |
| **`0813`** | `1` (内示) | `SUZUKI_SPIRITS_0813` | Monthly Forecast | Straightforward monthly forecast. Enters `ORDERS` cleanly without pre-existing records. |
| **`0600`** | `1` (内示) | `SUZUKI_SPIRITS_0600AND0630` | Main Monthly Forecast | The primary monthly forecast from Suzuki. Standard working-day delivery date calculation (`ADD_WORKING_DAYS2`). |
| **`6604`** | `1` (内示) | `SUZUKI_SPIRITS_6604AND6634` | Overseas / Spare Parts Forecast | Clean forecast insert. *(Requires part number with `'Z'` suffix in `PRDSLSODRM`)*. |

> [!TIP]
> **Recommended First Test:** Use a single **`0602`** or **`0600`** CSV file. Verify that records successfully transition into `ORDERS` with `STATUS = 'DUE_SET'` or `'CONFIRMED'`.

---

### Group B: 🟡 Firm Delivery Instructions (Converts & Consumes Forecasts)
Can be tested alone to create firm orders, but will consume quantities if a matching forecast already exists in `ORDERS` or `ORDERS_STAGE`.

| Code | Order Type | Staging Table | Description | Behavior |
| :--- | :--- | :--- | :--- | :--- |
| **`0650`** | `3` (納入指示) | `SUZUKI_SPIRITS_0650` | Daily Delivery Instruction | Converts `SHIP_TO` via `SUZUKI_SPIRITS_CONVERSION_MST`. Consumes matching active `0600` forecast. |
| **`0651`** | `3` (納入指示) | `SUZUKI_SPIRITS_0651` | Daily Delivery Instruction (with Notes) | Same as `0650`. If 17th character of notes is not `'0'`, marked as termination (打切). |
| **`0814`** | `3` (納入指示) | `SUZUKI_SPIRITS_0814` | Delivery Instruction | Consumes matching active `0813` forecast. |
| **`6624`** | `3` (納入指示) | `SUZUKI_SPIRITS_6624` | Overseas / Spare Parts Firm Order | Consumes matching active `6604` forecast. |

---

### Group C: 🔴 Modifier Adjustments (Cannot Be Tested Alone!)
These files **never** create new order records in `ORDERS`. They find a target order and modify its attributes.

| Code | Target File Code | Staging Table | Target Field | Behavior |
| :--- | :--- | :--- | :--- | :--- |
| **`0630`** | **`0600`** | `SUZUKI_SPIRITS_0600AND0630` | `DEMAND_QTY` | Matches `0600` row by part & date; **adds quantity** (`DEMAND_QTY = 0600 + 0630`). |
| **`663N`** | **`6624`** | `SUZUKI_SPIRITS_663NAND663SAND664T` | `DUE_DATE` | Matches `6624` row by customer PO number; **updates delivery due date**. |
| **`663S`** | **`6624`** | `SUZUKI_SPIRITS_663NAND663SAND664T` | `DEMAND_QTY` | Matches `6624` row by customer PO number; **updates order demand quantity**. |

---

### Group D: 🔗 Natural Testing Pairs (Multi-File Testing)
1. **`0600` + `0630`**: Test that the `0630` adjustment correctly increases the demand quantity of the `0600` forecast.
2. **`0600` + `0650`**: Test that the `0650` delivery instruction consumes the corresponding `0600` forecast quantity.
3. **`6624` + `663N` / `663S`**: Test that `663N` shifts the due date and `663S` modifies the quantity of the `6624` firm order.

---

## 🛠️ 2. How to Test the Stage 2 Pipeline

### Method 1: Automated Unit Tests (Visual Studio Test Explorer)
* **Location**: `OMS.Tests\SuzukiDataImportTests.vb`
* **Coverage**: Prorated date calculation, `0630` quantity addition, `663N`/`663S` PO matching and modification, `SUZUKI_SPIRITS_CONVERSION_MST` ship-to conversion, and forecast consumption pairs.
* **Run**: Visual Studio $\rightarrow$ **Test Explorer** $\rightarrow$ Run All.

---

### Method 2: Batch CLI Runner (`SUZUKI_SPIRITS_DATA_IMPORT.exe /batch`)
Runs the full pipeline end-to-end against the local database:
1. Ensure folders exist:
   ```powershell
   New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\SUZUKI\IN"
   New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\WORK"
   New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\COMPLETED"
   ```
2. Place test CSV(s) into `C:\ASTI\DATA\SUZUKI\IN\`.
3. Run:
   ```powershell
   & ".\SUZUKI_SPIRITS_DATA_IMPORT\bin\Debug\SUZUKI_SPIRITS_DATA_IMPORT.exe" /batch
   ```
4. Verify:
   * Files moved to `C:\ASTI\DATA\COMPLETED\SUZUKI-AT\5455\4\`.
   * Staging table (`SUZUKI_SPIRITS_*`) records marked `STATUS = 'REGISTERED'`.
   * New orders inserted into `ORDERS` and `ORDER_HISTORY`.
   * Downstream `STRA納期設定` lead times and dates calculated in `ORDERS`.

---

### Method 3: Web UI Flow (`OrderImport.aspx`)
1. Run `OMS.Web` locally.
2. Navigate to `Pages/Orders/OrderImportStage.aspx` (Stage 1 Pre-Import):
   * Select Customer `5455` and click **取込**. Files move to `WORK` and populate `SUZUKI_SPIRITS_*`.
3. Navigate to `Pages/Orders/OrderImport.aspx` (Stage 2 Order Import):
   * Select Customer `5455`, Folder Type `4 (混在)`.
   * Click **取込** (`btnImportFile`): Triggers `ExecuteStageImportToOrdersStage`. Validated and reconciled orders load into the grid.
   * Click **保存** (`btnSaveOrder`): Triggers `ExecuteSaveOrders`. Commits to `ORDERS`, updates staging records to `'REGISTERED'`, and archives WORK files to `COMPLETED/`.
   * *(Optional)* Click **破棄** (`btnCancelOrder`): Purges `ORDERS_STAGE` and returns WORK files back to `IN/`.
