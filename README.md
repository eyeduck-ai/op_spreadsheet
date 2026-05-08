# 手術排程系統 GAS

這個專案提供 Google Apps Script (`code.js`) 來管理 OP sheet，包含表格初始化、Tag 下拉選單、條件格式、時間格式正規化、同步 Google Calendar，以及 Calendar 反向同步回 Google Sheet。

## 首次安裝

1. 在 Google 試算表開啟「擴充功能」→「Apps Script」。
2. 將 `code.js` 的內容貼到 Apps Script 專案。
3. 請啟用 Calendar Advanced Service：
   - Apps Script 左側「服務」→「+」→ 選擇「Google Calendar API」。
   - 或將 `appsscript.json` 的內容套用到 Apps Script manifest。
4. 在 Apps Script 函式選單選擇 `setup`，執行一次並完成授權；若尚未設定 Calendar ID，系統會提示輸入。
5. 回到試算表並重新整理頁面，應會看到「手術排程系統」選單。
6. 若之後需要更換日曆，選擇「手術排程系統」→「維護工具」→「設定日曆」；系統會存到 Script Properties 並重建雙向同步綁定。

## 日常使用

- `維護工具` 子選單底部會顯示目前部署版本，例如 `版本：2026.05.09`。
- `維護工具` → `安裝`：重新初始化表格格式、批次同步既有日曆事件，並重裝 Sheet/Calendar 觸發器。
- `輸出日期`：輸出指定日期的 OP 病人與 IOL 清單；水晶體清單與病人清單會橫向並排。
- `輸出一週`：輸出今天起算 7 天的 OP 病人與 IOL 清單，依日期分段寫入同一張「輸出表單」。
- `複製列`：複製目前列，並清空日期、時間、eventID。
- `複製列` 只能在 `OP` 工作表執行，避免誤改輸出表單或其他工作表。
- `歸人整合`：合併非 OP 的同病歷號資料；若將被合併移除的列含日期、時間或 eventID，會中止並提示列號。
- `日期排序`：依日期與時間排序 OP sheet。
- `維護工具` 子選單：
  - `安裝`：執行完整初始化、觸發器安裝與既有列批次同步。
  - `設定日曆`：輸入並儲存手術日曆 ID 到 Apps Script 的 Script Properties；儲存後會重建 Sheet → Calendar 與 Calendar → Sheet 綁定，但不會自動批次同步既有列。
  - `初始化表格`：只重跑格式、驗證、條件格式與時間正規化。
  - `安裝同步`：只重裝 Sheet → Calendar 的 onEdit trigger。
  - `安裝反向同步`：建立 Calendar syncToken baseline，並安裝 Calendar → Sheet trigger。
  - `安裝自動輸出`：重裝未來一周清單輸出 trigger；每日約 06:00 自動更新，OP 表或 Calendar 反向同步變更後會延遲 5 分鐘更新。
  - `清除舊事件`：固定針對 `OP` 工作表的 `日曆eventID` 欄，刪除對應日曆事件；成功後只清空該 ID 儲存格。
  - `版本：2026.05.09`：顯示目前部署版本。

## 功能作用範圍

- 固定針對 `OP` 工作表：`輸出日期`、`輸出一週`、`歸人整合`、`日期排序`、`初始化表格`、`安裝` 批次同步、反向同步、`清除舊事件`。
- 針對目前工作表但限定 `OP`：`複製列`。
- 全域設定或 trigger：`設定日曆`、`安裝同步`、`安裝反向同步`、`安裝自動輸出`、版本顯示。

## 注意事項

- Sheet → Calendar 自動同步需要 installable onEdit trigger；第一次請執行 `setup()`。
- Sheet → Calendar 與 Calendar → Sheet 都需要 Calendar Advanced Service；若未啟用，`setup()` 會略過需要 Calendar API 的同步並顯示提醒。
- 未來一周自動輸出會由 `setup()` 或「安裝自動輸出」建立 time-driven trigger；每日 trigger 約 06:00 執行，Google 可能會讓實際時間前後微幅浮動。
- OP 表修改後不會每次立即重算；系統會刪除舊的 pending trigger 並排程 5 分鐘後更新，避免大量編輯時消耗過多免費 Apps Script 執行時間。
- Calendar ID 優先讀取 Script Properties 內的 `CALENDAR_ID`；`CONFIG.CALENDAR_ID` 只保留為舊版 fallback。
- 更換 Calendar ID 只會重建雙向同步綁定，不會自動把既有列批次同步到新日曆；若需要同步既有資料，請再執行「安裝」。
- 更新 `code.js` 後，請再執行一次 `setup()`，讓既有 Calendar event 套用新的標題與描述邏輯。
- 系統依第 1 列欄位名稱辨識資料欄；可以在資料表中間新增自訂欄位，但必要欄名需保留且避免重複。
- `日曆eventID` 是系統欄位，初始化時會依實際欄位位置自動隱藏；現在只存 Google Calendar API `event.id`。
- Plan 欄包含 `問`、`補`、`通知` 時資料列會套淡黃色；包含 `APPLY` 時套淡綠色且優先於淡黃色。
- 時間欄包含 `GA` 時，該時間儲存格會套紅底，Calendar 仍會以全天事件處理並把 `GA` 放入描述備註。
- 日曆事件標題格式為 `病歷號 | 姓名 | Condition`；分隔符 `|` 前後空白不影響解析，例如 `123|王小明|Cataract` 也可辨識，輸入全形 `｜` 也會自動視為分隔符。
- 反向同步仍會嘗試辨識舊標題格式 `病歷號 姓名 - Condition`，方便逐步轉換。
- description 會包含 TEL 與 Plan，例如 `TEL: 09xxxxxxxx`、`Plan: PE/IOL OD`。
- 有日期但沒有病歷號的列不會同步日曆，系統會在病歷號欄加 note 提醒。
- 時間欄空白或填入 `GA` 這類非時間備註時，會建立全天事件；`2460`、`25:00` 這類無效時間仍會被擋下。
- Calendar 反向同步遇到 description 內的 `Time note: GA` 時，會把 `GA` 保留回時間欄。
- 清空日期時，如果該列已有 eventID，系統會刪除對應 Calendar event 並清空 eventID。
- 日曆事件被刪除時，反向同步會清空日期、時間與 eventID，並在「心得」欄加 note 標記「日曆已刪除」。
- 從雙欄版本升級時，初始化會依欄名辨識舊系統欄 `日曆apiEventID`，必要時將值補回 `日曆eventID` 後移除舊欄。
- 清除舊日曆事件工具固定從 `OP` 工作表的 `日曆eventID` 欄第 2 列開始處理；刪除成功或事件已不存在時，只清空該 ID 儲存格，不會更動病歷號、姓名、日期、時間、Plan、心得等同列資料。
- 輸出表單中，病人清單的「術式」只保留換行前第一段；換行後文字會和原 Plan 欄完整內容合併到「補充說明」欄。
- 水晶體清單只列出成功解析到 `IOL(...)` 或 `MSICS(...)` 的病人，並依品牌排序。
- 未來一周輸出只寫入同一份試算表的「輸出表單」，不會自動產生 PDF、寄信或建立 Drive 檔案。
- 若想使用 Google Sheets 圖形按鈕，可將圖形指定給 `setup` 或其他公開函式。

## 本地檢查

- 語法檢查：`node --check code.js`
- Smoke tests：`node tests/smoke.js`

## Fork 與自動部署

- 這個 repository 只包含 Apps Script 程式碼與部署設定，不包含任何 Google Spreadsheet 資料。
- Fork 使用者需要建立自己的 spreadsheet-bound Apps Script 專案，並把自己的 Script ID 寫入本機 `.clasp.json`。
- GitHub Actions 自動部署需要設定自己的 repository secrets：`CLASP_JSON` 與 `CLASPRC_JSON`。
- `CLASPRC_JSON` 內含 Google OAuth refresh token，能代表授權帳號更新 Apps Script；只能放在 GitHub Secrets 或本機私密環境，不能提交到 git。
- `.clasp.json` 可能包含個人 Apps Script project ID，預設也不提交；請用 GitHub Secret `CLASP_JSON` 提供給 workflow。
- 本專案的 workflow 只在 `master` push 時執行 `clasp push --force`，不會因外部 fork 的 pull request 自動部署。

## 隱私與資料安全

- 這個工具會處理病患個資與醫療資訊，包含病歷號、姓名、TEL、Condition、Plan、心得、手術時間與輸出表單內容。
- 請使用專用 Google Calendar，並將分享權限限制在實際需要排程作業的人員；日曆標題會包含病歷號、姓名與 Condition，description 會包含 TEL 與 Plan。
- `輸出表單` 會產生病歷號、姓名、TEL、疾病、術式與補充說明，應視為含病患資料的工作表，不應公開分享或匯出到不受控的位置。
- `清除舊事件` 只處理目前試算表內固定 `OP` 工作表的 `日曆eventID` 欄，並刪除已設定 Calendar ID 內對應事件；使用前請確認該日曆是專用手術日曆。
- Calendar 反向同步使用的 `syncToken` 會存在 Apps Script `PropertiesService`，通常不包含病患內容，但仍應視為同步狀態資料，不要外洩。
- Apps Script 執行記錄不應包含完整病患資料；若自行新增 log，避免輸出姓名、病歷號、電話、Plan 或完整 Calendar title/description。
- 不要將含病患資料的 Google Sheet 匯出檔、Excel、CSV、PDF、截圖、備份檔或本地測試資料提交到 git。
