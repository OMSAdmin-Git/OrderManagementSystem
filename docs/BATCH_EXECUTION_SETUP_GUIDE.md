# 🚀 Batch Execution Setup & Task Scheduler Testing Guide

This document outlines the complete setup, directory structure, deployment steps, SQL verification queries, and Windows Task Scheduler configuration for running **`SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe`** on the VM Server.

---

## 📂 1. Directory Structure & Path Specifications

According to the official specifications (`SUZUKI_SPIRITS_DATA_PRE_IMPORT.pdf` & `SPIRITS_IMPORT_SPEC.md`):

| Purpose / Description | Path in Specification / Configuration | Notes |
| :--- | :--- | :--- |
| **Source Input Folder**<br/>*(取込元フォルダ)* | **`FOLDER_MST.FOLDER_PATH`**<br/>*(Default / Standard: `C:\ASTI\DATA\SUZUKI\IN`)* | Queried dynamically from `FOLDER_MST` where `CUSTOMER_CODE = '5455'` and `FOLDER_TYPE = 4` (`4: 混在`). |
| **Work Folder**<br/>*(作業用WORKフォルダ)* | **`C:\ASTI\DATA\WORK\SUZUKI-AT\5455\4\`**<br/>*(Format: `[WorkFolderRoot]\[UserId]\[CustomerCode]\[FolderType]\`)* | Created automatically by `Utils.EnsureDirectory`. Discovered files are moved here during processing. |
| **Completed Folder**<br/>*(取込完了フォルダ)* | **`C:\ASTI\DATA\COMPLETED\SUZUKI-AT\5455\4\`**<br/>*(Format: `[CompletedFolderRoot]\[UserId]\[CustomerCode]\[FolderType]\`)* | Successfully parsed CSV files are moved here after Stage 1 staging. |
| **Error List Folder**<br/>*(エラーリスト出力先)* | **`[FOLDER_PATH]\エラーリスト\`**<br/>*(e.g., `C:\ASTI\DATA\SUZUKI\IN\エラーリスト\`)* | If part numbers cannot be matched in `PRDSLSODRM`, `ErrorList_*.csv` is exported here. |
| **Log File**<br/>*(実行ログ)* | **`[BatchFolder]\app.log`** | Writes start, progress, and error logs for each run. |

---

## 🔍 2. SQL Queries to Retrieve & Verify Folder Path

### **Query 1: Retrieve Customer Setting ID for Suzuki Batch**
```sql
SELECT 
    customer_setting_id,
    customer_code,
    profit_center,
    customer_unit_id,
    active_flag
FROM customer_setting_mst
WHERE customer_code = '5455'
  AND profit_center IS NULL
  AND customer_unit_id IS NULL
  AND active_flag = 'Y';
```

### **Query 2: Direct Single Query to Retrieve Batch `FOLDER_PATH`**
```sql
SELECT 
    f.folder_id,
    f.customer_setting_id,
    c.customer_code,
    c.profit_center,
    f.folder_type,
    f.folder_path,
    f.active_flag
FROM folder_mst f
INNER JOIN customer_setting_mst c ON f.customer_setting_id = c.customer_setting_id
WHERE c.customer_code = '5455'
  AND c.profit_center IS NULL
  AND c.customer_unit_id IS NULL
  AND f.folder_type = 4
  AND f.active_flag = 'Y'
  AND c.active_flag = 'Y';
```

---

## 🏗️ 3. Deployment & Folder Layout on the VM Server

```text
C:\ (or D:\)
 ├── ASTI\
 │    ├── DATA\
 │    │    ├── SUZUKI\IN\             <-- Source EDI/CSV/Excel input folder
 │    │    ├── WORK\                  <-- WorkFolderRoot (C:\ASTI\DATA\WORK)
 │    │    └── COMPLETED\             <-- CompletedFolderRoot (C:\ASTI\DATA\COMPLETED)
 │    │
 │    └── OMS\
 │         ├── Web\                   <-- IIS Published Web App (C:\inetpub\wwwroot\OMS or D:\ASTI\OMS\Web)
 │         │    ├── bin\
 │         │    ├── Pages\
 │         │    └── Web.config
 │         │
 │         └── Batch\                 <-- Standalone Batch Console Binaries
 │              ├── SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe
 │              ├── SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe.config
 │              ├── OMS.Business.dll
 │              ├── OMS.Data.dll
 │              ├── OMS.Common.dll
 │              ├── Oracle.ManagedDataAccess.dll
 │              ├── CsvHelper.dll
 │              └── app.log
```

---

## 🧪 4. Testing & Execution Steps on the VM Server

### **Step 1: Create Base Folders via PowerShell**
```powershell
New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\SUZUKI\IN"
New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\WORK"
New-Item -ItemType Directory -Force -Path "C:\ASTI\DATA\COMPLETED"
```

### **Step 2: Deploy Batch Binaries**
1. In Visual Studio, set configuration to **Release**.
2. Build the `SUZUKI_SPIRITS_DATA_PRE_IMPORT` project.
3. Copy the output files from `bin/Release/` to `C:\ASTI\OMS\Batch\`.
4. Open `SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe.config` and verify the `OMSConnection` connection string points to your target Oracle database.

### **Step 3: Manual Execution Test**
1. Place a sample Suzuki CSV file into `C:\ASTI\DATA\SUZUKI\IN\`.
2. Open PowerShell / Command Prompt on the VM and execute:
   ```powershell
   cd C:\ASTI\OMS\Batch\
   .\SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe /batch
   ```
3. **Verify Results**:
   - Console output: `[SUCCESS] Processed X file(s) successfully.`
   - Check `C:\ASTI\DATA\COMPLETED\SUZUKI-AT\5455\4\` (processed CSV files moved here).
   - Check Oracle Database (`IMP_RUN`, `IMP_FILES_STAGE`, `SUZUKI_SPIRITS_...` tables).
   - Check `app.log` in `C:\ASTI\OMS\Batch\`.

### **Step 4: Windows Task Scheduler Configuration**
1. Open **Task Scheduler** (`taskschd.msc`).
2. Select **Create Task...**
   - **General Tab**:
     - Name: `ASTI_Suzuki_Spirits_Data_Pre_Import`
     - Select: **"Run whether user is logged on or not"**
     - Check: **"Run with highest privileges"**
   - **Triggers Tab**:
     - Set your desired schedule (e.g., Daily at 21:00 or Hourly).
   - **Actions Tab**:
     - Action: `Start a program`
     - Program/script: `C:\ASTI\OMS\Batch\SUZUKI_SPIRITS_DATA_PRE_IMPORT.exe`
     - Add arguments: `/batch`
     - **Start in (Mandatory)**: `C:\ASTI\OMS\Batch\`
3. Save the task and enter the administrative password if prompted.
4. Right-click the task $\rightarrow$ **Run** to test. Confirm the **Last Run Result** shows `(0x0)`.
