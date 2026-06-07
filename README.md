# 手術排程系統 GAS

這個專案提供 Google Apps Script (`code.js`) 來管理 `All` sheet，包含表格初始化、醫院與 Tag 下拉選單、條件格式、時間格式正規化、Sheet → Google Calendar 同步，以及產生分院工作表與手術清單。

## 首次安裝

1. 在 Google 試算表開啟「擴充功能」→「Apps Script」。
2. 將 `code.js` 的內容貼到 Apps Script 專案。
3. 請啟用 Calendar Advanced Service：
   - Apps Script 左側「服務」→「+」→ 選擇「Google Calendar API」。
   - 或將 `appsscript.json` 的內容套用到 Apps Script manifest。
4. 若是從舊版升級，請先在試算表中手動將既有 `OP` 工作表改名為 `All`；系統只會提示，不會自動改名。
5. 在 Apps Script 函式選單選擇 `setup`，執行一次並完成授權；若尚未設定 Calendar ID，系統會提示輸入。
6. 回到試算表並重新整理頁面，應會看到「手術排程系統」選單。
7. 從舊版升級後，請執行「手術排程系統」→「維護工具」→「停用反向同步」一次，清除舊版 Calendar → Sheet trigger。
8. 若之後需要更換日曆，選擇「手術排程系統」→「維護工具」→「設定日曆」；系統會存到 Script Properties 並重建 Sheet → Calendar 同步綁定。

## 日常使用

- `維護工具` 子選單底部會顯示目前部署版本，例如 `版本：2026.06.07.2`。
- `維護工具` → `安裝`：重新初始化表格格式、批次同步既有日曆事件，並重裝 Sheet/Calendar 觸發器。
- `輸出日期`：輸出指定日期的 OP 病人資料，依 `醫院` 分成 `高榮` 與 `聯醫` 區塊，寫入「手術清單」與「水晶體清單」。
- `輸出一週`：輸出今天起算 7 天的 OP 病人資料，依日期與醫院分段寫入「手術清單」與「水晶體清單」。
- `複製列`：複製目前列，並清空日期、時間、CalendarEventId 與 CalendarSheetWriteUpdated。
- `複製列` 只能在 `All` 工作表執行，避免誤改輸出清單、分院鏡像表或其他工作表。
- `歸人整合`：合併非 OP 的同病歷號資料；若將被合併移除的列含日期、時間或 eventID，會中止並提示列號。
- `日期排序`：依日期與時間排序 `All` sheet。
- `維護工具` 子選單：
  - `完整安裝／批次同步`：執行完整初始化、觸發器安裝與既有列批次同步。
  - `設定日曆`：輸入並儲存手術日曆 ID 到 Apps Script 的 Script Properties；儲存後會重建 Sheet → Calendar 綁定，但不會自動批次同步既有列。
  - `初始化表格`：只重跑格式、驗證、條件格式與時間正規化。
  - `資料遷移（不批次同步）`：只補齊/遷移系統欄位，例如將舊 `日曆eventID` 搬到 `CalendarEventId`，確保 `醫院` 在 `Tag` 左側，並重新隱藏與保護系統欄。
  - `檢查 All 工作表`：檢查主要工作表是否已命名為 `All`；若只找到舊 `OP`，會提示手動改名。
  - `更新分院工作表`：重新產生唯讀鏡像 `OP-高榮` 與 `OP-聯醫`。
  - `安裝同步`：只重裝 Sheet → Calendar 的 onEdit trigger。
  - `停用反向同步`：刪除舊版 Calendar → Sheet trigger，並清除舊 syncToken；Calendar 編輯不會回寫 `All`。
  - `安裝自動輸出`：重裝未來一周清單輸出 trigger；每日約 06:00 自動更新，`All` 表變更後會延遲 5 分鐘更新。
  - `清除舊事件`：固定針對 `All` 工作表的 `CalendarEventId` 欄，刪除對應日曆事件；成功後清空該 ID 與同列 `CalendarSheetWriteUpdated`。
  - `版本：2026.06.07.2`：顯示目前部署版本。

## 功能作用範圍

- 固定針對 `All` 工作表：`輸出日期`、`輸出一週`、`歸人整合`、`日期排序`、`初始化表格`、`完整安裝／批次同步`、`清除舊事件`。
- 針對目前工作表但限定 `All`：`複製列`。
- 由 `All` 產生並覆寫：`OP-高榮`、`OP-聯醫`、`手術清單`、`水晶體清單`。
- 全域設定或 trigger：`設定日曆`、`資料遷移`、`安裝同步`、`停用反向同步`、`安裝自動輸出`、版本顯示。

## 注意事項

- Sheet → Calendar 自動同步需要 installable onEdit trigger；第一次請執行 `setup()`。
- Sheet → Calendar 需要 Calendar Advanced Service；若未啟用，`setup()` 會略過需要 Calendar API 的同步並顯示提醒。
- 未來一周自動輸出會由 `setup()` 或「安裝自動輸出」建立 time-driven trigger；每日 trigger 約 06:00 執行，Google 可能會讓實際時間前後微幅浮動。
- `All` 表修改後不會每次立即重算；系統會刪除舊的 pending trigger 並排程 5 分鐘後更新，避免大量編輯時消耗過多免費 Apps Script 執行時間。
- Calendar ID 優先讀取 Script Properties 內的 `CALENDAR_ID`；`CONFIG.CALENDAR_ID` 只保留為舊版 fallback。
- 更換 Calendar ID 只會重建 Sheet → Calendar 綁定，不會自動把既有列批次同步到新日曆；若需要同步既有資料，請再執行「完整安裝／批次同步」。
- 更新 `code.js` 後，請再執行一次 `setup()`，讓既有 Calendar event 套用新的標題與描述邏輯。
- 系統依第 1 列欄位名稱辨識資料欄；可以在資料表中間新增自訂欄位，但必要欄名需保留且避免重複。
- `醫院` 欄位會放在 `Tag` 左側，資料驗證選項為 `高榮`、`聯醫`；空白或任何不是 `聯醫` 的值，都會歸到高榮輸出。
- `OP-高榮` 與 `OP-聯醫` 是由 `All` 產生的鏡像表，會在安裝、資料遷移、編輯、合併、排序或手動更新時重建；請只在 `All` 編輯資料。
- `CalendarEventId` 與 `CalendarSheetWriteUpdated` 是系統欄位，初始化/資料遷移時會依實際欄位位置自動隱藏並嘗試保護；`CalendarEventId` 只存 Google Calendar API `event.id`。
- 系統欄位保護會保留目前執行者，若需額外保留維護者可在 Script Properties 設定 `SYSTEM_COLUMN_PROTECTION_EDITORS`，以逗號、分號或換行分隔 email。
- Plan 欄包含 `問`、`補`、`通知` 時資料列會套淡黃色；包含 `APPLY` 時套淡綠色且優先於淡黃色。
- 時間欄包含 `GA` 時，該時間儲存格會套紅底，Calendar 仍會以全天事件處理並把 `GA` 放入描述備註。
- 日曆事件標題格式為 `病歷號 | 姓名 | Condition`；中間欄位會保留位置，姓名空白時會寫成 `病歷號 |  | Condition`。分隔符 `|` 前後空白不影響解析，例如 `123|王小明|Cataract` 也可辨識。
- description 會包含 TEL 與 Plan，例如 `TEL: 09xxxxxxxx`、`Plan: PE/IOL OD`。
- 有日期但沒有病歷號的列不會同步日曆，系統會在病歷號欄加 note 提醒。
- 時間欄空白或填入 `GA` 這類非時間備註時，會建立全天事件；`2460`、`25:00` 這類無效時間仍會被擋下。
- Sheet 寫入 Calendar 後，系統會把 Calendar 回傳的 `updated` 記錄到 `CalendarSheetWriteUpdated`，用來追蹤 Sheet-originated event 狀態。
- Calendar → Sheet 反向同步已停用；若直接在 Calendar 修改或刪除事件，變更不會回寫 `All`，請以 `All` 作為資料來源。
- 清空日期時，如果該列已有 eventID，系統會刪除對應 Calendar event 並清空 eventID。
- 從舊版本升級時，請執行 `資料遷移` 或 `初始化表格`；系統會辨識舊系統欄 `日曆eventID` / `日曆apiEventID`，必要時將值補回 `CalendarEventId` 後移除無衝突的舊欄。
- 清除舊日曆事件工具固定從 `All` 工作表的 `CalendarEventId` 欄第 2 列開始處理；若尚未遷移，仍會相容舊 `日曆eventID` 欄。刪除成功或事件已不存在時，會清空該 ID 與同列 `CalendarSheetWriteUpdated`，不會更動病歷號、姓名、日期、時間、Plan、心得等同列資料。
- `手術清單` 中，「術式」只保留換行前第一段；換行後文字會和原 Plan 欄完整內容合併到「補充說明」欄。
- `水晶體清單` 只列出成功解析到 `IOL(...)` 或 `MSICS(...)` 的病人，並依品牌排序；若該醫院日期有手術病人但沒有水晶體資料，會顯示「本日無水晶體資料」。
- 未來一周輸出只寫入同一份試算表的「手術清單」與「水晶體清單」，不會自動產生 PDF、寄信或建立 Drive 檔案。
- 若想使用 Google Sheets 圖形按鈕，可將圖形指定給 `setup` 或其他公開函式。

## 本地檢查

- 語法檢查：`node --check code.js`
- 測試語法檢查：`node --check tests/smoke.js`
- Smoke tests：`node tests/smoke.js`

## Fork 與自動部署

- 這個 repository 只包含 Apps Script 程式碼與部署設定，不包含任何 Google Spreadsheet 資料。
- Fork 使用者需要建立自己的 spreadsheet-bound Apps Script 專案，並把自己的 Script ID 寫入本機 `.clasp.json`。
- GitHub Actions 自動部署需要設定自己的 repository secrets：`CLASP_JSON` 與 `CLASPRC_JSON`。
- `CLASPRC_JSON` 內含 Google OAuth refresh token，能代表授權帳號更新 Apps Script；只能放在 GitHub Secrets 或本機私密環境，不能提交到 git。
- `.clasp.json` 可能包含個人 Apps Script project ID，預設也不提交；請用 GitHub Secret `CLASP_JSON` 提供給 workflow。
- 本專案的 workflow 只在 `master` push 時執行 `clasp push --force`，不會因外部 fork 的 pull request 自動部署。

## 隱私與資料安全

- 這個工具會處理病患個資與醫療資訊，包含病歷號、姓名、TEL、Condition、Plan、心得、手術時間與輸出清單內容。
- 請使用專用 Google Calendar，並將分享權限限制在實際需要排程作業的人員；日曆標題會包含病歷號、姓名與 Condition，description 會包含 TEL 與 Plan。
- `手術清單`、`水晶體清單`、`OP-高榮` 與 `OP-聯醫` 會產生或複製病患資料，應視為含病患資料的工作表，不應公開分享或匯出到不受控的位置。
- `清除舊事件` 只處理目前試算表內固定 `All` 工作表的 `CalendarEventId` 欄，並刪除已設定 Calendar ID 內對應事件；使用前請確認該日曆是專用手術日曆。
- 舊版 Calendar 反向同步使用的 `syncToken` 可能仍存在 Apps Script `PropertiesService`；執行「停用反向同步」會清除這些同步狀態資料。
- Apps Script 執行記錄不應包含完整病患資料；若自行新增 log，避免輸出姓名、病歷號、電話、Plan 或完整 Calendar title/description。
- 不要將含病患資料的 Google Sheet 匯出檔、Excel、CSV、PDF、截圖、備份檔或本地測試資料提交到 git。
