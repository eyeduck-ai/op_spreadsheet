# 手術排程與 FU 追蹤系統

版本 **2026.10.11.2**。這是綁定 Google Spreadsheet 的 GAS；本版本只維護線上刀表的格式、conditional formatting、時間正規化、排序、月份／刀日建立、跨表快照與年度封存。**GAS 不讀取、新增、更新或刪除任何 Google Calendar 事件**，也不安裝日曆掃描／重試觸發器。

SOAP → 刀表由平台 EXE 經使用者確認後寫入；刀表 → 日曆由指定 Windows 主機的監控 EXE 負責，包含人工輸入、改期、取消與刪列。日曆不回寫臨床資料；監控程序只維護隱藏的識別欄。

## 模組

- `code.js`：設定、選單、共用工具及格式 onEdit 安裝。
- `sheet_model.js`：依表頭名稱定位欄位、格式、日期區塊、時間正規化及原生排序。
- `schedule_integrity.js`：刀表結構、重複識別檢查、人工編輯格式與安全的舊觸發器入口。舊日曆引擎與 Script Properties 事件索引已移除。
- `calendar_handover.js`：固定排程識別欄遷移、診斷及 protocol 2 主機交接。
- `workflows.js`：FU／月表快照、建立月表與刀日、排序及水晶體清單。
- `annual_archive.js`：私人整檔備份與永久年度封存。

## 欄位與格式

所有欄位按**表頭名稱**解析，不依賴固定字母／欄序。臨床欄可保留中文，自訂欄保留資料與格式；系統欄使用英文。新增臨床欄若只是人工備註，不影響同步。若要納入日曆或計算，需同步更新 GAS 與監控程序的欄位模型、格式規則及測試。

FU 可見標準欄：`病歷號、姓名、TEL、醫院、Tag、Condition、日期、Plan、心得`。

月表可見標準欄：`日期／報到時間、醫院、病歷號、姓名、TEL、GA、側別、診斷、Grade、術式、IOL、IOL Target、IOL Final、Axis、Plan、心得、空白分隔欄、SN、CDE、Energy Time、Energy %、屈光數據`。

FU、All 與 YYYYMM 月表有兩個隱藏系統欄：

| 欄位 | 用途 |
| --- | --- |
| `ScheduleSyncId` | 32 位小寫十六進位 UUID。病人排程與每個 `◆ 刀日` 標題各自有永久識別碼，改期與搬列保留。 |
| `CalendarEventId` | 監控 EXE 管理的 Calendar 事件綁定，保留所有既有值。GAS 不建立或清除事件 ID。 |

新增月表、刀日及跨表建立的病人會有新的 ScheduleSyncId；空白手動輸入列保持空白，由監控 EXE 在有效資料出現後配置識別。月表 → FU、FU → 月表是新的獨立排程，配置新的 ScheduleSyncId，不能複製來源事件綁定。

ScheduleSyncId 欄預設隱藏並限制編輯；CalendarEventId 沿用既有保護，不會自行解除。擁有者須授權平台／監控服務帳戶與需要搬列的操作者編輯這兩個系統欄。安裝／格式修復保留擁有者後續授予的權限。

FU 日期格式為 `yyyy/m/d`。刀日標題使用 `yyyy/m/d ddd` 日期、醫院及 `◆ 刀日` 標記，病人報到時間按既有規則正規化。只有功能視窗輸入使用八位數 YYYYMMDD，直接編輯儲存格不改成八位數格式。

刀日與院區以**實際標題內容**為準。整個刀日改期後，即使實際月份暫時與 YYYYMM 分頁名稱不同，格式、排序、監控及水晶體清單仍依實際刀日運作。新增排程功能仍預設放入日期所屬月表；年度封存會拒絕日期／來源月份不符，避免誤封存未來排程，請先搬到正確月份再封存。

新月表診斷表頭保留既有管理公式：以 CATA／Retina／Plasty 與 OD／OS／OU 計數。一般安裝只更新已受管理的公式，不覆蓋人工公式；重複表頭、公式錯誤與混合日期／事件列需先處理。Plan 條件格式、下拉選項、欄寬與人工備註的既有規則保留。

## 選單與日常操作

`手術排程系統` 保留建立新月刀表、選取列新增刀日（可選高榮／聯醫）、整理目前分頁、水晶體清單、月刀表 → FU、FU → 月刀表及彙整舊月刀表。

`維護工具` 提供：

1. **遷移固定排程識別欄**：第一次升級時明確執行，為既有 FU／All、YYYYMM 與年度封存表追加 ScheduleSyncId，配置病人／刀日識別碼。全表先檢查結構與重複 ID；有衝突時不開始寫入。既有值、公式、備註與 CalendarEventId 保留；部分失敗後可重跑，不重建已配置 ID。
2. **檢查排程識別欄**：列出重複、格式錯誤、只剩 ID 及缺少識別的資料。
3. **交由本系統同步日曆**：輸入日曆 ID 及指定 worker_id，寫入 protocol 2 文件 metadata；不存取日曆。
4. **顯示日曆交接狀態**、**安裝／修復系統**與**套用所有 FU／月表建議欄寬**。

安裝只保留一個 `processScheduleEdit` onEdit 觸發器，處理格式與時間。onChange／Calendar 重試及舊同步觸發器會清除；其他使用者尚未移除的舊入口也安全停用。沒有固定五分鐘 GAS 背景掃描。

## 搬列、改期與移除

- 以列號選取**整列**再剪貼、拖移或排序，包含兩個隱藏 ID；原生排序使用完整表頭寬度，包含自訂欄與系統欄。
- 修改刀日日期／院區時，保留該刀日及病人的 ScheduleSyncId。監控 EXE 會更新原事件。
- 新增另一筆排程請填空白列；複製既有列而保留 ID 會產生重複，監控與 GAS 診斷均阻擋有疑義的操作。
- 病歷號刪節線由監控程序將日曆事件轉灰。純粹調整列順序不重新分配日曆時間。
- 人工刪除整個病人列後，監控程序會跨表重新定位並經緩衝核對後刪除綁定事件；清空 ID、剪貼中間狀態、整張分頁消失及結構／讀取異常都不是直接刪除依據。
- FU 尚未填日期時不建立事件；既有事件再清空日期的疑義交由監控待確認。
- GAS 不因上述操作呼叫 Calendar。已過期排程的處理由監控範圍決定。

## 年度封存

保留私人整檔備份、暫存年度頁、完整回讀驗證與交易恢復流程。來源月表只在驗證成功後移除；永久封存頁包含兩個隱藏識別欄及 ArchiveSourceMonth／ArchiveRowType，臨床值、公式、備註與識別碼按表頭重映射保留。封存頁本身不操作日曆。`OP` 仍是舊名稱，不會自動改名；請依舊流程手動改為 All。

## 部署與交接

`master` 分支 push 後，GitHub Actions 先跑語法、smoke 與刀表契約測試，再以既有 `CLASPRC_JSON`／`CLASP_JSON` secrets 執行 `clasp push`。Actions 成功代表綁定 Apps Script 的原始碼已更新；不會自動執行選單、識別欄遷移、保護範圍授權或主機交接。首次升級後重新整理刀表以載入新版選單。

1. 先備份 GAS，並在 Google 刀表副本與測試日曆驗收。
2. 部署全部 JS、HTML 與 appsscript.json；manifest 不再啟用 Calendar 進階服務。服務帳戶 JSON 不能代替擁有者部署 GAS。
3. 先執行「遷移固定排程識別欄」，檢查新增欄的保護，授權服務帳戶。
4. 執行「安裝／修復系統」。所有曾安裝舊觸發器的帳號應一併移除舊 Calendar 觸發器。
5. 指定 Windows 主機完成 init，於「交由本系統同步日曆」輸入同一 Calendar ID 與 worker_id。
6. 監控 EXE 通過 check 後再啟用。GAS 在交接前也不會寫日曆；升級與交接期間會暫停日曆更新。

文件 metadata `vghks.calendar.control`：

```json
{"protocol":2,"owner":"application","gas_calendar_disabled":true,"calendar_id":"CALENDAR_ID","time_zone":"Asia/Taipei","gas_version":"2026.10.11.2","worker_id":"32位小寫十六進位主機識別"}
```

不保存病人原文。這是指定唯一主機的設定，並非分散式鎖；移機前必須停止舊主機，攜帶完整資料目錄，不可同時執行兩台。

本版本沒有恢復 GAS 日曆同步選單，舊 restore 入口只顯示停用提示。回退舊部署前應先停止監控主機、備份兩個系統欄與主機資料，再由擁有者處理；不要讓舊 GAS 與主機同時操作日曆。

## 本機驗證

```powershell
node --check code.js
node --check sheet_model.js
node --check schedule_integrity.js
node --check calendar_handover.js
node --check workflows.js
node --check annual_archive.js
node --check tests/smoke.js
node tests/smoke.js
node tests/sheet_contract.js
git diff --check
```

目前 71 項 smoke tests 通過，涵蓋格式、排序、封存、識別遷移、重複 ID、跨月整區改期、跨表重試及主機交接。Calendar 引擎與其測試已移除，由監控程序測試接手事件生命週期。另以 `tests/sheet_contract.js` 核對共用刀表契約：欄名、下拉、欄寬、月份日期、診斷公式與欄位換序後的條件格式。平台持有同一份契約，CI 對比實際 GAS 模型。

這些是合成資料測試；不表示雲端已部署，也不代替刀表副本的拖移、保護範圍與觸發器權限驗收。
