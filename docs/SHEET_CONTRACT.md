# 刀表契約與相容性更新

正式來源為 `contracts/sheet-contract.json`，契約版本為 **1.0.0**。`CONFIG.VERSION` 是 GAS 部署日期版本；兩者各自維護。公開規格及案例只包含欄位、規則與合成資料。

主程式負責院內 SOAP 與刀表更新；GAS 負責格式、排序、刀日與封存；獨立 `vghks-sheet2cal` 負責 Calendar 寫入。Python 兩端安裝本 repo 的 `packages/python` 共用套件，固定完整 Git commit，不在執行中下載新規格。

## 正式來源與產物

- `contracts/sheet-contract.json`：欄名、別名、型別語意、院區、永久 ID、列類型與刪除保護。
- `contracts/cases/`：從實際 GAS 匯出的日期、欄位、時間、建立／搬列及原生封存案例。
- `packages/python/src/vghks_sheet_core/`：不依賴 Google 或院內 SDK 的解析、識別與共享基準讀寫。
- `tools/generate_contract.py`：由規格產生 `code.js` 的 CONFIG 區塊及 Python 套件內的規格／案例；其餘 GAS 解析演算法由相同案例驗證。

欄位以表頭解析，保留自訂欄、公式與註記；重複正式欄名或永久 ID 需停止並核對。整列搬移保留兩個 ID；日期與院區由最近有效刀日決定。取消以病歷號顯示刪節線判定。FU 無日期不建立事件；已綁定 FU 清空日期保留事件並要求核對。封存保留 ID，未完成封存阻擋寫入；完整成功掃描的連續缺列確認才允許刪除。

## 更新步驟

1. 修改規格、對應演算法與合成案例，檢查新增欄位是否需要同步至日曆。
2. 執行 `python tools/generate_contract.py`，並審閱生成差異。
3. 執行 `node tests/formal_contract.js`、既有 GAS 測試及 `python -m vghks_sheet_core.verify --gas-repo .`。
4. `tools/verify_consumer.py --bot-repo <checkout>` 核對指定主程式版本與候選契約，包含主程式及日曆解析的 82 個案例。
5. 更新 Python 套件版本與各消費端固定的 Git commit，在各自 CI 執行回歸。
6. GAS 的部署 job 必須等待相容性 job 成功；破壞性改動先發布可同時讀新舊格式的消費端，再遷移資料，最後移除舊格式。

語意版本：新增相容能力遞增 minor；修正且保持語意相容時遞增 patch；移除欄位／別名或改變識別與取消／刪除語意時遞增 major。新 major 必須由消費端明確支援。僅更新一般主程式功能不要求 GAS 與監控同步升版。

`contracts/consumers.json` 記錄部署閘門所驗證的主程式完整 commit。獨立監控自己的 CI 會取固定 GAS commit，執行真實匯出器與監控回歸；調整契約時須一併更新該依賴並確認 CI。
