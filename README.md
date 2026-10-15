# 手術排程與 FU 追蹤系統

版本 **2026.10.15**。GAS 負責線上刀表格式、conditional formatting、時間正規化、排序、建立月份／刀日、FU 跨表操作與年度封存。**GAS 不存取 Google Calendar，不登記監控主機，也不負責既有資料的識別欄遷移。**

平台 EXE 辨識 SOAP 的手術／FU Tag，由使用者確認後更新刀表。指定 Windows 主機的監控 EXE 處理刀表 → 日曆，包含人工新增、改期、取消與刪列；日曆不回寫刀表。

## 使用現有刀表與部署

現有線上刀表已採新格式，不需要 GAS 初始化。直接使用現有刀表與 EXE 即可。GAS 更新後重新整理試算表，就會載入新版選單。

push 到 `master` 後，GitHub Actions 先執行語法、smoke 與共用刀表契約測試，再用既有 secrets 執行 `clasp push`，並唯讀核對正式刀表綁定。

**監控設定全部在監控 EXE 完成，不需要 GAS 選單操作：**輸入刀表網址、日曆 ID、服務帳戶金鑰後，程式自動檢查權限、追加／辨識系統欄、配置永久識別碼、隱藏系統欄、登記主機及回讀核對。監控維持停用，準備完成後由使用者啟用。

原有嚴格保護不會自動解除或降低；服務帳戶無權編輯時，監控 EXE 會列出需由擁有者授權的分頁／欄位及帳戶 email。授權後在 EXE 重試即可接續。既有識別碼重複或欄位衝突時保留原值，要求核對。

設定、金鑰、指定主機身分與未完成操作保存在監控主機的資料庫。更新 EXE 並保留完整資料資料夾，不需重新匯入或登記。移機前停止舊主機；新資料目錄要接手既有刀表時，在 EXE 明確確認替換主機。

`維護工具 → 修復刀表格式` 是選用的修復工具，只處理既有 FU／All 與月表的格式、下拉驗證及編輯觸發器。它不建立 FU、不批次改寫既有報到時間、不配置既有病人或刀日的識別碼，也不作為主程式／監控的啟動步驟。缺少系統欄時，在 EXE 執行「檢查並準備刀表」；缺少臨床欄時先人工核對第一列表頭。

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

新增病人使用指定日期＋院區區塊末端的空白列；有公式（即使顯示空字串）、註記或資料時插入完整新列，保留原內容且不跨入下一刀日。病人醫院欄寫入目的刀日院區，報到時間仍獨立保存。未命名的人工尾欄也納入資料列掃描與排序。

系統欄只自動隱藏，不建立警告或嚴格保護。FU、月表、新月份、格式修復及年度封存均不自動建立、移除或調整受保護範圍；既有人工保護設定保留。完整搬列仍須包含兩個隱藏識別欄。

FU 日期格式為 `yyyy/m/d`；刀日標題為 `yyyy/m/d ddd` 日期、醫院及 `◆ 刀日`，可選高榮／聯醫。只有功能視窗輸入使用八位數 YYYYMMDD。刀日以列內實際日期及院區為準，整區改期不要求先更改分頁名稱。

刀日必須同時有有效完整日期、`高榮` 或 `聯醫`、`◆ 刀日`（標記內空白可正規化），且 `CalendarEventId` 空白。缺院區、缺標記、無效日期或混入事件 ID 時保留原內容並回報結構衝突；不再預設補高榮或清除標記。日期只接受 2000–2099 年完整西元日期，附英文三字星期時必須吻合，日期型別也必須為午夜整日。

舊日期欄名 `時間`、`日期/報到時間` 相容映射至 `日期／報到時間`；診斷統計只接受 `診斷｜` 顯示前綴或 `SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2` 系統公式標記。建立新月沿用最近既有月的全部表頭、公式、註記、格式及欄寬；若月份間臨床、技術或人工欄位位置不同，停止建立並要求先對齊，不默默建立另一種欄序。

每次欄位解析都核對表頭公式，避免第二個無顯示文字的診斷公式漏過重複檢查。新月欄位比對只忽略整張表皆無內容、公式或註記的尾端空欄；未命名但有人工資料或註記的欄位仍保留位置。

診斷管理表頭保留 CATA／Retina／Plasty 與 OD／OS／OU 計數公式；人工公式不覆寫。Plan 條件格式、下拉、欄寬及人工備註沿用既有規則。

## 選單

日常選單保留建立新月刀表、選取列新增刀日、整理目前分頁、水晶體清單、月刀表 → FU、FU → 月刀表及彙整舊月刀表。維護工具只有刀表格式修復、建議欄寬及版本；沒有交接或識別欄遷移入口。

格式修復只重建 `processScheduleEdit` onEdit，移除執行者可見的舊格式／Calendar／匯出觸發器；不更動其他觸發器或 Script Properties。GAS 不再保存日曆設定／主機登記或 Calendar 索引。

僅保留八個歷史已安裝觸發器名稱：`processRowChange` 只轉交目前編輯格式處理；`processCalendarStructureChange`、`retryPendingCalendarSync`、`processMonthlyStructureChange`、`processCalendarChange`、`exportUpcomingWeekData`、`runPendingUpcomingWeekExport`、`autoExportByEdit` 全為無讀寫空操作，避免其他安裝者的舊觸發器報錯。它們不是選單功能；舊日曆設定、交接、識別遷移、健康檢查及修復入口已移除。

## 搬列、取消及封存

年度封存與 Windows 6.12.0 共用原生格式契約及固定雲端互斥標記。兩邊均採「預覽 → 備份 → 複製並回讀核對 → 再次確認移除原月份」，可補入較早月份，以原生插列保留年度分頁 ID 與月份排序。GAS 先保存私人 Drive 整檔備份及壓縮原生交易計畫；Script Properties 僅保存精簡交易進度。未知結果只核對原交易，不重送、不自動解除互斥標記。

啟用 manifest 內的 Sheets v4 進階服務；使用既有綁定專案與部署目標，標準 Cloud 專案亦須啟用 Sheets API。封存保存刪節線、富文字、註記、驗證、列高、隱藏狀態及合併儲存格，固定條件格式顯示；不調整人工保護。既有封存格式轉換列於預覽並先備份。舊 GAS 未保存可靠取消狀態的月份須以私人備份核對，Windows 統計標示不完整；新月份區段以 `VGHKS_ARCHIVE_CONTRACT_V2` 註記辨識。

互斥使用 `DOCUMENT` metadata `vghks.sheet.archive.lock`、固定 ID 2147483646，主程式、GAS 與監控均尊重；同一 ID 的原子建立拒絕第二個封存。未完成的舊版 GAS 交易不自動續跑舊刪除流程。更新時一併更新主程式與監控；統計介面由主程式提供。

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

`code.js` 是選單／設定，`sheet_model.js` 是欄位／格式，`schedule_integrity.js` 是結構／人工編輯格式及集中保留的舊觸發器空入口，`schedule_identity.js` 只驗證識別，`workflows.js` 是日常操作，`annual_archive.js` 是年度封存。原 `calendar_handover.js` 與日曆引擎已移除；FU 安排到月表的 request token 只快取永久 ScheduleSyncId，不再依事件 ID、原月份、列號或內容雜湊猜測同一排程。重送重新定位並核對病歷號；跨月搬移或監控更新事件綁定均不重建病人，識別重複或病人不符時停止。

```powershell
node --check code.js
node --check sheet_model.js
node --check schedule_identity.js
node --check schedule_integrity.js
node --check workflows.js
node --check annual_archive.js
node --check archive_native.js
node --check archive_transaction.js
node --check tests/smoke.js
node tests/smoke.js
node tests/sheet_contract.js
node tests/monitor_handover.js
node tests/parser_contract.js
node tests/workflow_contract.js
node tests/archive_contract.js
git diff --check
```

73 項合成 smoke tests 涵蓋格式、排序、封存、識別衝突、跨月改期、只隱藏不新增保護及保留人工保護、舊觸發器無寫入及簡化選單。實際格式修復回歸另驗證沒有 FU 時不建表、不配置既有識別碼、不批次改報到時間，保留人工值／公式／註記／保護；缺欄提示按系統欄與臨床欄分流。共用刀表契約另核對欄名、下拉、欄寬、月份日期、診斷公式與欄位換序後的條件格式。監控交接測試使用實際 GAS 解析器，核對同日不同院區、跨月整列搬移、隱藏 ID 保留、只搬可見欄位診斷與主程式狀態頁排除。實際 Google 權限與自動遷移由監控 EXE 的副本驗證負責。

2026.10.14 另以 `workflow_contract.js` 執行實際新增月份、刀日與病人寫入，匯出供兩個 EXE 讀取的合成儲存格；核對同日雙院區、新列院區、空白列使用、空顯示公式與註記保留、永久 ID 重送及整列跨月改期。沒有增加雲端系統欄、隱藏分頁或日曆權限。
