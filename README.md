# 手術排程與 FU 追蹤系統

版本 **2026.10.11.4**。GAS 負責線上刀表格式、conditional formatting、時間正規化、排序、建立月份／刀日、FU 跨表操作與年度封存。**GAS 不存取 Google Calendar，不登記監控主機，也不負責既有資料的識別欄遷移。**

平台 EXE 辨識 SOAP 的手術／FU Tag，由使用者確認後更新刀表。指定 Windows 主機的監控 EXE 處理刀表 → 日曆，包含人工新增、改期、取消與刪列；日曆不回寫刀表。

## 首次設定

push 到 `master` 後，GitHub Actions 先執行語法、smoke 與共用刀表契約測試，再用既有 secrets 執行 `clasp push`。成功代表雲端 GAS 原始碼已更新；重新整理試算表即可載入新版選單。

**監控設定全部在監控 EXE 完成，不需要 GAS 選單操作：**輸入刀表網址、日曆 ID、服務帳戶金鑰後，程式自動檢查權限、追加／辨識系統欄、配置永久識別碼、隱藏並加警告保護、登記主機及回讀核對。監控維持停用，準備完成後由使用者啟用。

原有嚴格保護不會自動解除或降低；服務帳戶無權編輯時，監控 EXE 會列出需由擁有者授權的分頁／欄位及帳戶 email。授權後在 EXE 重試即可接續。既有識別碼重複或欄位衝突時保留原值，要求核對。

設定、金鑰、指定主機身分與未完成操作保存在監控主機的資料庫。更新 EXE 並保留完整資料資料夾，不需重新匯入或登記。移機前停止舊主機；新資料目錄要接手既有刀表時，在 EXE 明確確認替換主機。

`安裝／修復刀表格式` 僅供需要修復格式、驗證或格式觸發器時使用，與監控首次設定無關。舊 Calendar 觸發器入口是安全空操作；本版不新增定時日曆掃描。

## 欄位與格式

所有欄位按表頭名稱解析，不依賴固定字母／欄序；自訂欄位、人工值、備註與公式保留。若要新欄位出現在日曆或計算中，須更新對應欄位模型與共用契約。

FU 可見欄：`病歷號、姓名、TEL、醫院、Tag、Condition、日期、Plan、心得`。

月表可見欄：`日期／報到時間、醫院、病歷號、姓名、TEL、GA、側別、診斷、Grade、術式、IOL、IOL Target、IOL Final、Axis、Plan、心得、空白分隔欄、SN、CDE、Energy Time、Energy %、屈光數據`。

隱藏系統欄：

| 欄位 | 用途 |
| --- | --- |
| `ScheduleSyncId` | 32 位小寫十六進位 UUID。每筆病人排程及每個刀日標題各自有永久識別碼；改期、搬列均保留。 |
| `CalendarEventId` | 監控 EXE 的日曆事件綁定；既有值保留，GAS 不建立或清除。 |

GAS 建立新月表、刀日及跨表病人時會配置新的 ScheduleSyncId；手動空白列保持空白，監控 EXE 在資料出現後配置。月表 → FU、FU → 月表是另一筆排程，不能沿用來源的兩個 ID。

新系統欄預設隱藏及警告保護，方便平台／監控／人工完整搬列。既有嚴格保護與擁有者授予的編輯者均保留；格式修復不把監控建立的警告保護變成嚴格保護。

FU 日期格式為 `yyyy/m/d`；刀日標題為 `yyyy/m/d ddd` 日期、醫院及 `◆ 刀日`，可選高榮／聯醫。只有功能視窗輸入使用八位數 YYYYMMDD。刀日以列內實際日期及院區為準，整區改期不要求先更改分頁名稱。

診斷管理表頭保留 CATA／Retina／Plasty 與 OD／OS／OU 計數公式；人工公式不覆寫。Plan 條件格式、下拉、欄寬及人工備註沿用既有規則。

## 選單

日常選單保留建立新月刀表、選取列新增刀日、整理目前分頁、水晶體清單、月刀表 → FU、FU → 月刀表及彙整舊月刀表。維護工具只有刀表格式修復、建議欄寬及版本；沒有交接或識別欄遷移入口。

格式修復只安裝 `processScheduleEdit` onEdit，清除安裝者可見的舊 Calendar／onChange／重試觸發器；其他使用者的舊入口亦安全停用。GAS 不再保存新的日曆 ID／主機登記或 Calendar Script Properties 索引。

## 搬列、取消及封存

同一天可以有不同院區的刀日區塊。例如 10/12 高榮與 10/12 聯醫是兩個區塊；監控一律使用患者上方最近有效刀日的完整日期與院區，病人列舊院區由監控正規化。分頁名稱不是事件日期。刀日標題本身有 ScheduleSyncId，沒有 CalendarEventId，也不建立日曆事件。

跨月改期請用列號選取整列（含隱藏欄）剪貼到目的刀日之下；監控以相同 ScheduleSyncId 更新原事件，並沿用 CalendarEventId。只貼可見欄位可能遺留或遺失 ID：原列只剩 ID、ID 重複、同一 ID 改成不同病歷號，或舊排程消失同時出現無 ID 新列，都須先核對，不能把相似姓名／日期當作同一事件。GAS onEdit 不會推測剪貼來源或修補這類綁定。

主程式的隱藏 `__VGHKS_SYNC_STATE` 保存 SOAP 更新的欄位基線，不含 SOAP 內容；GAS 月表/FU 格式與排序忽略它，年度私人備份則保留。監控使用自己的 SQLite 事件紀錄及 spreadsheet developerMetadata 主機登記，不改寫主程式基線。兩程式寫入前重新讀取並回讀核對；這是衝突偵測，Google Sheets 沒有提供跨電腦列寫入 CAS。

- 以列號選取完整列後拖移、剪貼或排序，包含兩個隱藏 ID；改期保留 ID。
- 新增另一筆手術請填空白列；複製既有 ID 會被重複檢查阻擋。
- 病歷號刪節線由監控 EXE 將事件轉灰。整列刪除由監控跨表核對及緩衝後移除事件；整頁消失或 ID 異常不作為批次刪除依據。
- FU 沒有日期時不建立事件；既有事件再清空日期則待確認。
- 年度封存保留私人備份、交易恢復與完整核對；封存保留兩個 ID、ArchiveSourceMonth／ArchiveRowType。日期與來源月份不符時拒絕封存，避免誤封存未來排程。
- `OP` 是舊名稱，不自動改名；如需使用舊追蹤表，依既有方式改為 All。

## 模組與驗證

`code.js` 是選單／設定，`sheet_model.js` 是欄位／格式，`schedule_integrity.js` 是結構／人工編輯格式，`schedule_identity.js` 只驗證識別與保留安全舊入口，`workflows.js` 是日常操作，`annual_archive.js` 是年度封存。原 `calendar_handover.js` 與日曆引擎已移除。

```powershell
node --check code.js
node --check sheet_model.js
node --check schedule_identity.js
node --check schedule_integrity.js
node --check workflows.js
node --check annual_archive.js
node --check tests/smoke.js
node tests/smoke.js
node tests/sheet_contract.js
node tests/monitor_handover.js
git diff --check
```

69 項合成 smoke tests 涵蓋格式、排序、封存、識別衝突、跨月改期、保護設定保留、舊入口無寫入及簡化選單。共用刀表契約另核對欄名、下拉、欄寬、月份日期、診斷公式與欄位換序後的條件格式。監控交接測試使用實際 GAS 解析器，核對同日不同院區、跨月整列搬移、隱藏 ID 保留、只搬可見欄位診斷與主程式狀態頁排除。實際 Google 權限與自動遷移由監控 EXE 的副本驗證負責。
