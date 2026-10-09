# 📄 SPIRITS データ受注登録 仕様書・設計書 (Stage 2 Specification)

> **文書名:** SPIRITSデータ受注登録 仕様書・設計書  
> **プロジェクト名:** `SUZUKI_SPIRITS_DATA_IMPORT`  
> **システム名:** 受注データ抽出システム  
> **お客様名:** ASTI株式会社 様  
> **作成元:** 天方産業 株式会社  
> **バージョン:** 1.0.0 (2026/08/20)  
> **対象テーブル:** `ORDERS_STAGE` (受注ワーク), `ORDERS` (受注本番), `ORDER_HISTORY` (受注履歴)  

---

## 🏗️ 1. 概要 (Overview)

`SPIRITSデータ取込` (Stage 1: `SUZUKI_SPIRITS_DATA_PRE_IMPORT`) で取得・検証してスズキ専用テーブルに登録したデータから、受注データにあたるものを抽出し、加工・変換を行って `ORDERS_STAGE` (受注ワークテーブル) および本番の `ORDERS` (受注テーブル) に登録する。

### 🎛️ 1.1 実行形態および EXE 仕様 (Execution Modes & Console EXE Architecture)

本機能はコンソールアプリケーション (`SUZUKI_SPIRITS_DATA_IMPORT.exe`) および Web画面 (`OrderImportStage.aspx` / `OrderImport.aspx`) の両方から実行可能とする。

| 実行形態 | トリガー / 実行元 | 実行ユーザー (`CREATED_USER_ID`) | 処理範囲 (Scope) | エラー出力先 | 下流連携 (`STRA納期設定`) |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **🤖 バッチ実行** | Windows タスクスケジューラ<br>`SUZUKI_SPIRITS_DATA_IMPORT.exe /batch` | 固定値 `'SUZUKI-AT'` | `ORDERS_STAGE` 登録 → `ORDERS` 保存 → `STRA納期設定` まで全自動実行 | `[FOLDER_PATH]\エラーリスト\` (CSV) | **自動実行あり** |
| **💻 画面実行** | Web画面 (`OrderImportStage.aspx`)<br>手動ボタン押下 | ログインユーザーID (例: `EMP1040`) | `ORDERS_STAGE` (ワーク) への登録・加工まで (画面上で確認・手動保存) | 画面上のポップアップ/一覧 ＋ CSV | **画面では実行しない** |

#### 💻 コンソール実行 (EXE) コマンドライン引数仕様
```text
SUZUKI_SPIRITS_DATA_IMPORT.exe [/batch | /manual] [/user:<UserID>] [/folder:<FolderPath>] [/customersettingid:<ID>] [/profitcenter:<PC>]
```
* `/batch` : バッチモードで起動 (`SUZUKI-AT` ユーザー、STRA納期設定まで一括実行)。
* `/manual` : 手動モードで起動 (ワーク登録のみ)。
* `/user:<ID>` : 実行ユーザーIDを指定 (省略時: バッチは `'SUZUKI-AT'`)。
* `/customersettingid:<ID>` : 処理対象の `CUSTOMER_SETTING_ID` を指定 (省略時: DBより `CUSTOMER_CODE='5455'` を自動取得)。
* `/folder:<PATH>` : フォルダパスを直接指定 (省略時: `FOLDER_MST` または `App.config` より取得)。
* **終了コード (Exit Code):** 正常終了時 `0` / エラー発生時 `1`。

---

## 🔄 2. 全体処理フロー図 (Process Flow Diagram)

```text
[START]
   │
[DELETE ①] ─── 受注ワークテーブル (ORDERS_STAGE) の物理削除
   │
[処理対象確認] ─── バッチ実行: スズキ専用テーブル / 画面実行: 選択された取込ファイル
   │
[実行場所チェック] 
   ├── 画面実行 ──> [ASTI追加内示チェック]
   │                   ├── HAND_FLAG = 'Y' ──> [ファイル情報取得・ファイル読込] ──┐
   │                   └── HAND_FLAG ≠ 'Y' ──> [受注データチェック] ──────────────┤
   └── バッチ実行 ─────────────────────────────> [受注データチェック] ──────────────┤
                                                                                  │
┌─────────────────────────────────────────────────────────────────────────────────┘
│
├── [対象 0件] ──> [END]
└── [対象 > 0件]
      │
   [INSERT (受注ワーク)] ─── スズキ専用テーブルから ORDERS_STAGE へ挿入 (※0630, 663N, 663Sは除外)
      │
   [UPDATE ① (受注ワーク)] ─── PC設定更新 (USRDEFFLDF) & 特殊コード処理 (0630加算, 663N納期更新, 663S数量更新)
      │
   [エラーチェック] 
      ├── エラーあり ──> [エラー情報出力] ──> [ファイルコピー] ──> [DELETE ② (受注ワーク)] ──> [END]
      └── エラーなし
            │
         [UPDATE ② (受注ワーク)]
            ├── 内示加工 (0602新規扱い、その他内示洗い替え)
            ├── 受注残加工 / 納入指示加工
            │     ├── 出荷先変換 (SUZUKI_SPIRITS_CONVERSION_MST)
            │     └── 内示消込 (0650/0651→0600, 0814→0813, 6624→6604)
            │
         [実行場所チェック]
            ├── 画面実行 ──> [END] (画面上で確認後、手動保存へ)
            └── バッチ実行 ──> [受注データ保存 (ORDERS登録)] ──> [STRA納期設定] ──> [END]
```

---

## 📑 3. 詳細処理手順 (Step-by-Step Specifications)

### 【工程 1】 受注ワーク削除 (`DELETE ①`)
処理開始時、`ORDERS_STAGE` (受注ワークテーブル) から以下すべての条件を満たすレコードをすべて物理削除する。
1. `CUSTOMER_SETTING_ID = 処理中のCUSTOMER_SETTING_ID` (取引先設定ID)
2. `FOLDER_TYPE = 処理中のFOLDER_TYPE` (フォルダ区分)

---

### 【工程 2】 処理対象確認
遷移元 (バッチ実行 / 画面実行) ごとに処理対象データを取得する。
* **バッチ実行の場合:** `SPIRITSデータ取込` で登録したスズキ専用テーブル (`SUZUKI_SPIRITS_*`)。
* **画面実行の場合:** `gvSelectCustomers[chkStageImport]` にチェックがあるデータの `IMP_FILES_STAGE[IMP_FILE_STAGE_ID]` に紐づくスズキ専用テーブル または 受注ファイル。

---

### 【工程 3】 ASTI追加内示チェック (画面実行時のみ)
対象受注ファイルが ASTI追加内示 (≒ハンド作成ファイル) であるか確認する。

| 取得項目 | 参照元 | 書式・条件 |
| :--- | :--- | :--- |
| **HAND_FLAG** (ハンドフラグ) | `chkHandFlag` (ハンドフラグチェックボックス) | `gvStageImportFiles` (取込準備ファイル一覧) = `'Y'` |

* **分岐:**
  * **取得できた場合 (`HAND_FLAG = 'Y'`):** 【工程 4】ファイル読込 へ遷移する。
  * **取得できなかった場合:** 【工程 5】受注データチェック へ遷移する。

---

### 【工程 4】 ASTI追加内示 ファイル情報取得 & ファイル読込 (`HAND_FLAG = 'Y'`)

#### 1. マスタ参照
* **`MAPPING_PROFILE_MST`:** `CUSTOMER_SETTING_ID`, `FOLDER_TYPE`, `ACTIVE_FLAG = 'Y'` で検索し、`PROFILE_ID`, `FILE_ID`, `HEADER_ROW_INDEX`, `DATA_START_ROW_INDEX`, `DEFAULT_SHEET_NAME` を取得。
* **`FILE_MST`:** `PROFILE_ID` に紐づく `FORMAT_TYPE`, `FILE_TYPE`, `DELIMITER`, `ENCLOSURE`, `HEADER_FLAG`, `LINE_ENDING`, `CHARSET` を取得。
* **`FIELD_MAPPING_MST`:** 列並び・セル位置の定義情報を取得。

#### 2. ファイル読み込みマッピング

| 項目名 | `FORMAT_TYPE = 'LIST'` 取得元 | `FORMAT_TYPE = 'MATRIX'` 取得元 | 抽出条件 / 備考 |
| :--- | :--- | :--- | :--- |
| **CUSTOMER_ORDER_NO** (客先発注No) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'CUSTOMER_ORDER_NO'` |
| **ORDER_DATE** (受注日) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'ORDER_DATE'` |
| **DUE_DATE** (希望納期) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'DUE_DATE'` |
| **CUSTOMER_ITEM_NO** (客先品目No) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'CUSTOMER_ITEM_NO'` |
| **DEMAND_QTY** (需要数) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'DEMAND_QTY'` |
| **CURRENCY_CODE** (通貨コード) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'CURRENCY_CODE'` |
| **PRODUCT_CODE** (製品コード) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'PRODUCT_CODE'` |
| **REMARKS** (コメント) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'REMARKS'` |
| **DELIVERY_CODE** (納入先コード) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'DELIVERY_CODE'` |
| **ORDER_TYPE** (受注区分) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'ORDER_TYPE'` (FOLDER_TYPE=4のみ) |
| **PRORATED_TYPE** (分割区分) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `FILE_MST.TARGET_FIELD = 'PRORATED_TYPE'` (FOLDER_TYPE=4のみ) |
| **CUSTOMER_INFO_TYPE** (取引先情報区分) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'CUSTOMER_INFO_TYPE'` |
| **INFO_TYPE** (情報区分) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'INFO_TYPE'` |
| **SELF_FCST_FLAG** (ASTI追加フラグ) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'SELF_FCST_FLAG'` |
| **SELF_FCST_DELETE_FLAG** (ASTI追加削除フラグ)| `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'SELF_FCST_DELETE_FLAG'` ([ASTI追加フラグ]='Y'のみ) |
| **INFO_TYPE_CODE** (情報区分コード) | `SOURCE_COLUMN_INDEX` | `SOURCE_CELL_ADDRESS` | `TARGET_FIELD = 'INFO_TYPE_CODE'` |

---

### 【工程 5】 受注データチェック (`受注データチェック`)
スズキ専用テーブルより、`INFO_TYPE_CODE` (情報区分コード) が以下のいずれかに合致するレコードを対象とする。

* **対象コード一覧 (全11種):**
  * `0600` : 確定受注データ（月次内示）
  * `0602` : 確定受注データ（追加内示）
  * `0630` : 確定受注データ（内示増減）
  * `0650` : 納入指示データ（通常）
  * `0651` : 納入指示データ（臨時）
  * `0813` : 支給品情報データ
  * `0814` : 支給品納入指示データ
  * `6604` : 注文データ（内示）
  * `6624` : 注文データ（確定/指示）
  * `663N` : 注文データ（納期変更）
  * `663S` : 注文データ（数量変更）

* **判定:**
  * 対象データが 1 件以上ある場合 → 【工程 6】`INSERT (受注ワーク)` へ遷移。
  * 対象データが 0 件の場合 → 次の処理対象へ遷移 (または終了)。

---

### 【工程 6】 受注ワーク挿入 (`INSERT ORDERS_STAGE`)
スズキ専用テーブルから `ORDERS_STAGE` へインサートする。

> ⚠️ **重要除外ルール:**  
> `INFO_TYPE_CODE` が **`0630`**, **`663N`**, **`663S`** のデータは、初期インサートを行わない (【工程 7】の UPDATE で反映する)。

| `ORDERS_STAGE` 列名 | インサート値 | 参照元テーブル / 取得元 | 書式・結合条件・設定ルール |
| :--- | :--- | :--- | :--- |
| **CUSTOMER_SETTING_ID** | `CUSTOMER_SETTING_ID` | `IMP_FILES_STAGE` | 取込ファイルワークテーブル |
| **CUSTOMER_CODE** | `CUSTOMER_CODE` | `CUSTOMER_SETTING_MST` | 取引先設定マスタ |
| **BILLING_TO** | `FBILLTOCD` | `SECTM` (部署マスタ) | `SECTD.FSECTTYP = 'CU'` (部署明細マスタの部署タイプ) |
| **CUSTOMER_ORDER_NO** | `[客先発注No]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **DEMAND_STATUS** | `'F'` または `'O'` | 固定値 | `0600/0602/0813/6604` → `'F'`<br>`0650/0651/0814/6624` → `'O'` |
| **SHIP_TO** | `[出荷先]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **ORDER_DATE** | `[受注日]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **DUE_DATE** | 稼働日計算後の希望納期 | スズキ専用テーブル | `yyyy/MM/dd`<br>`SELECT ADD_WORKING_DAYS2('00001', [希望納期], 0) FROM DUAL` |
| **CUSTOMER_ITEM_NO** | `[客先品目No]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **ITEM_NO** | `[品目No]` | スズキ専用テーブル | SPIRITSデータマッピング参照 (`PRDSLSODRM` 紐づけ済品目) |
| **DEMAND_QTY** | `[需要数]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **DEMAND_UNIT** | `FUNIT` (受払単位) | `ITEMM` (工場別品目マスタ) | `ITEMM.FITEMNO = PRDSLSODRM.FPRDCD` |
| **CURRENCY_CODE** | `FCURR` (通貨コード) | `SECTM` (部署マスタ) | `SECTM.FSECTCD = CUSTOMER_SETTING_MST.CUSTOMER_CODE` |
| **SHIP_STOCK_LOCATION**| `FSHPWHCD` (出荷在庫場所) | `ITEMM` (工場別品目マスタ) | `ITEMM.FITEMNO = ORDERS_STAGE.ITEM_NO` |
| **COMPANY_ID** | `'1000'` | 固定値 | - |
| **PRODUCT_CODE** | `[品目No]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **BILLING_STANDARD** | `'S'` | 固定値 | - |
| **SHIP_PROCESS_TYPE** | `'O'` または `'K'` | 固定値 | `0600/0602/0813/6604` → `'O'`<br>`0650/0651/0814/6624` → `'K'` |
| **DELIVERY_INSTR_FLAG**| `'N'` または `'Y'` | 固定値 | `0600/0602/0813/6604` → `'N'`<br>`0650/0651/0814/6624` → `'Y'` |
| **REMARKS** | `[コメント]` | スズキ専用テーブル / 取込ファイル | - |
| **DELIVERY_CODE** | `[納入先コード]` | スズキ専用テーブル / 取込ファイル | - |
| **TOTAL_SHIP_QTY** | 空欄 または `0` | 固定値 | `0600/0602/0813/6604` → `NULL` (空欄)<br>`0650/0651/0814/6624` → `0` |
| **TRANSPORT_METHOD** | `2` | 固定値 | - |
| **PRE_DAILY_ORDER_QTY**| `[需要数]` | スズキ専用テーブル | 日割前受注数 |
| **PRE_DAILY_DELIVERY_DATE**| `[希望納期]` | スズキ専用テーブル | 日割前納期 (`yyyy/MM/dd`) |
| **IMP_FILE_STAGE_ID** | `IMP_FILE_STAGE_ID` | `IMP_FILES_STAGE` | 一時取込ファイルID |
| **ORDER_TYPE** | `1` または `3` | スズキ専用テーブル | `0600/0602/0630/6604/663N/663S` → `1` (内示)<br>`0650/0651/0814/6624` → `3` (納入指示) |
| **PRORATED_TYPE** | `[分割区分]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **CUSTOMER_INFO_TYPE** | `[取引先情報区分]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **INFO_TYPE** | `INFO_TYPE` | `INFO_TYPE_MST` | `CUSTOMER_INFO_TYPE = 取得ファイル.[情報区分]` |
| **SELF_FCST_FLAG** | `[ASTI追加フラグ]` | 取込ファイル | - |
| **SELF_FCST_DELETE_FLAG**| `[ASTI追加削除フラグ]` | 取込ファイル | - |
| **RECONCILE_TYPE** | `RECONCILE_TYPE` | `IMP_RULE_MST` | `IMP_RULE_MST.CUSTOMER_SETTING_ID = IMP_FILES_STAGE.CUSTOMER_SETTING_ID` かつ `IMP_RULE_MST.FOLDER_TYPE = 処理中のFOLDER_TYPE` |
| **IMP_RUN_ID** | `IMP_RUN_ID` | `IMP_RUN` | 取込実行ID |
| **STATUS** | `'IMPORTED'` | 固定値 | - |
| **ACTIVE_FLAG** | `'Y'` | 固定値 | - |
| **CREATED_AT** | `STARTED_AT` | `IMP_RUN` | 実行開始日時 (`yyyy/MM/dd HH:mm:ss`) |
| **CREATED_USER_ID** | `[ログインユーザーID]` | ログイン情報 / 実行引数 | - |
| **CREATED_PG_ID** | `'OrderImport(Execute)'`| 固定値 | - |
| **UPDATED_AT** | `STARTED_AT` | `IMP_RUN` | 実行開始日時 |
| **UPDATED_USER_ID** | `[ログインユーザーID]` | ログイン情報 / 実行引数 | - |
| **UPDATED_PG_ID** | `'OrderImport(Execute)'`| 固定値 | - |
| **DELIVERY_TIME** | `[納入時間]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **CONTAINER_CAPACITY** | `[容器収容数]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **INITIAL_FLAG** | `[初品区分]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **TARGET_REFERENCE_DATE_TYPE**| `[対象基準日区分]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **TARGET_REFERENCE_DATE**| `[対象基準日]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |
| **INFO_TYPE_CODE** | `[情報区分コード]` | スズキ専用テーブル | SPIRITSデータマッピング参照 |

---

### 【工程 7】 受注ワーク更新 (`UPDATE ①`)

#### 1. 取引先設定ID (`CUSTOMER_SETTING_ID`) の解決更新
各レコードの製品コードに基づき、所属する Profit Center (PC) を特定して正しい `CUSTOMER_SETTING_ID` に更新する。
1. **PC取得:**
   ```sql
   SELECT FUSRSTR1 FROM USRDEFFLDF 
   WHERE FTABLEID = 'ITEMM' AND FRECKEY = '[PRODUCT_CODE(製品コード)]';
   ```
2. **取引先設定ID取得 & 更新:**
   ```sql
   SELECT CUSTOMER_SETTING_ID FROM CUSTOMER_SETTING_MST 
   WHERE CUSTOMER_CODE = '5455' AND PROFIT_CENTER = '[取得したPC]';
   ```
   *(※取得できなかった場合は `CUSTOMER_SETTING_ID` を更新しない)*

#### 2. 特殊コード固有の更新処理

##### 《0630: 確定受注データ（内示増減）》
* **更新対象列:** `DEMAND_QTY` (需要数)
* **更新計算式:** `ORDERS_STAGE.DEMAND_QTY + SUZUKI_SPIRITS_0600AND0630.DEMAND_QTY`
* **合致条件:**
  * 同一 `[品目No]`
  * 同一 `[日割前納期]`
  * 同一 `[対象基準日区分]`
  * 同一 `[対象基準日]`
  * `[情報区分コード] = '0600'`
* **フォールバック:** スズキ専用テーブルから取得できなかった場合は、本番の `ORDERS` (受注テーブル) から取得し、取得できたデータを `ORDERS_STAGE` に反映して `DEMAND_QTY` を更新する。

##### 《663N: 注文データ（納期変更）》
* **更新対象列:** `DUE_DATE` (希望納期)
* **更新値:** `SUZUKI_SPIRITS_663NAND663SAND664T.[希望納期]`
* **合致条件:**
  * 同一 `[客先発注No]`
  * `[情報区分コード] = '6624'`
* **フォールバック:** スズキ専用テーブルから取得できなかった場合は、本番の `ORDERS` (受注テーブル) から取得し、取得できたデータを `ORDERS_STAGE` に反映して `DUE_DATE` を更新する。

##### 《663S: 注文データ（数量変更）》
* **更新対象列:** `DEMAND_QTY` (需要数)
* **更新値:** `SUZUKI_SPIRITS_663NAND663SAND664T.[需要数]`
* **合致条件:**
  * 同一 `[客先発注No]`
  * `[情報区分コード] = '6624'`
* **フォールバック:** スズキ専用テーブルから取得できなかった場合は、本番の `ORDERS` (受注テーブル) から取得し、取得できたデータを `ORDERS_STAGE` に反映して `DEMAND_QTY` を更新する。

---

### 【工程 8】 エラーチェック & ロールバック

#### 1. エラー判定条件
* 受注ファイル取込ページの基本チェックに準ずる。
* **追加エラー条件:**
  * `CUSTOMER_SETTING_ID` が `CUSTOMER_CODE = '5455'` かつ `PROFIT_CENTER` (PC) と `CUSTOMER_UNIT_ID` (工場/担当者ID) が **ブランクのまま残っているデータ** (PC特定に失敗したデータ)。

#### 2. エラー発生時の処理
1. **エラー情報通知:**
   * **バッチ実行:** エラー内容を CSV/Excel ファイルにまとめて `[FOLDER_MST.FOLDER_PATH]\エラーリスト\` に出力する。
   * **画面手動実行:** エラー内容を画面上のメッセージ/ダイアログに出力する。
2. **ファイルコピー:** エラー対象となったファイルを取り込み元に戻す (ファイル名: `[元のファイル名]_[ユーザーID]_[処理日時]`)。
3. **物理削除 (`DELETE ②`):** エラー対象となったレコードを `ORDERS_STAGE` から物理削除する。

---

### 【工程 9】 受注ワーク更新 (`UPDATE ②`: 内示加工・受注残加工・納入指示加工)

#### 1. 内示加工処理 (洗い替え条件)
* **`INFO_TYPE_CODE = '0602'` (追加内示) の特別ルール:**
  * 対象基準日系の項目が存在しないため、**常に新規データとして取り扱う** (洗い替え先は考慮せず、既存内示を消し込まずに新規追加する)。
* **その他内示コードの洗い替え条件:**
  * `SHIP_TO = 処理中のSHIP_TO`
  * `ITEM_NO = 処理中のITEM_NO`
  * `TARGET_REFERENCE_DATE_TYPE = 処理中のTARGET_REFERENCE_DATE_TYPE`
  * `TARGET_REFERENCE_DATE = 処理中のTARGET_REFERENCE_DATE`
  * `INFO_TYPE_CODE = 処理中のINFO_TYPE_CODE`
  * `ACTIVE_FLAG = 'Y'`

#### 2. 受注残加工・納入指示加工 (出荷先変換 & 内示消込)

##### ① 出荷先変換ルール (`SUZUKI_SPIRITS_CONVERSION_MST`)
* `SUZUKI_SPIRITS_CONVERSION_MST` (スズキSPIRITS納入ホーム変換マスタ) を参照する。
  * **検索条件:** `DELIVERY_CODE_ORDER (納入指示納入先コード) = 処理中のSHIP_TO` かつ `ACTIVE_FLAG = 'Y'`
  * **変換:** 取得できた `DELIVERY_CODE_PLAN` (内示納入先コード) を `SHIP_TO` に置き換える。
  * *(※取得できなかった場合は変換せず、元の `SHIP_TO` のままとする)*

##### ② 内示消込ルール (Forecast Offset / Consumption)
変換後の `SHIP_TO` および `ITEM_NO` に基づき、処理中の情報区分コードに対応する過去のアクティブ内示データを消し込む (`ACTIVE_FLAG = 'N'` またはステータス更新)。

| 処理中の納入指示コード (`INFO_TYPE_CODE`) | 消込対象の内示コード (`INFO_TYPE_CODE`) |
| :---: | :---: |
| **`0650` / `0651`** (通常/臨時 納入指示) | **`0600`** (月次内示) |
| **`0814`** (支給品納入指示) | **`0813`** (支給品情報データ) |
| **`6624`** (確定注文/指示) | **`6604`** (内示注文) |

---

### 【工程 10】 本番保存 & STRA納期設定

#### 1. 実行場所判定
* **画面手動実行の場合:** `ORDERS_STAGE` の更新完了をもって処理を正常終了する (画面上の「保存」ボタン押下によって別途本番反映される)。
* **バッチ自動実行の場合:** 自動的に【工程 10-2】受注データ保存 へ進む。

#### 2. 受注データ保存 (`INSERT ORDERS` / `INSERT ORDER_HISTORY`)
* `ORDERS_STAGE` から本番テーブル `ORDERS` へ有効な受注データをインサートする。
* 同時に変更履歴として `ORDER_HISTORY` へ履歴レコードを記録する。
* スズキ専用テーブルの対象レコードをステータス `'REGISTERED'` に更新する。

#### 3. STRA納期設定
* バッチ実行の最終ステップとして、下流システム連携の `STRA納期設定` (STRA Delivery Date Setting) をキック/実行する。
* `IMP_RUN` のステータスを `'COMPLETED'` に更新し、実行ログを完了させる。

---

## 📊 4. 設計・開発チェックリスト (Implementation Checklist)

- [ ] **DBアクセス層 (`OMS.Data`):**
  - [ ] `ORDERS_STAGE` の削除 (`DELETE ①`, `DELETE ②`) メソッドの実装
  - [ ] 各スズキ専用テーブルから `ORDERS_STAGE` への抽出・一括 INSERT クエリの実装 (`ADD_WORKING_DAYS2` 含む)
  - [ ] `USRDEFFLDF` / `CUSTOMER_SETTING_MST` による PC 設定更新クエリの実装
  - [ ] `0630`, `663N`, `663S` の特殊加算・更新ロジック (ORDERS フォールバック含む)
  - [ ] `SUZUKI_SPIRITS_CONVERSION_MST` による出荷先変換 & 消込クエリの実装
  - [ ] `ORDERS` / `ORDER_HISTORY` への本番コミット処理の実装
- [ ] **ビジネスロジック層 (`OMS.Business` / `SuzukiDataImportService.vb`):**
  - [ ] `ExecuteOrderRegistration` メソッドへの全工程オーケストレーション実装
  - [ ] バッチ実行 vs 画面実行の引数・分岐制御
  - [ ] エラーリスト CSV 出力ロジックの実装
- [ ] **STRA連携:**
  - [ ] `OrderStraRepository` / STRA 納期設定プロシージャの呼び出し実装
