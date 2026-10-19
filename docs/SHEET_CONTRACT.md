# 共用刀表規格

契約版本 **1.1.0**。正式來源是 `contracts/sheet-contract.json`；GAS 的日期部署版號另行維護。

| 來源 | 用途 |
| --- | --- |
| contracts/sheet-contract.json | 欄位、別名、日期、院區、ID、同步及刪除規則 |
| contracts/cases/ | 實際 GAS 匯出的合成解析、寫入與封存案例 |
| packages/python/src/vghks_sheet_core/ | Python 共用解析、永久識別、共享基準與 Calendar 資料模型 |
| tools/generate_contract.py | 生成 GAS CONFIG、Python 內嵌 JSON 與案例 |
| contracts/consumers.json | 部署前交叉驗證的固定主程式候選 commit |

## 共通行為

- 依表頭名稱定位，保留自訂欄、公式、註記及人工保護。
- 日期與院區由最近有效刀日決定。**月份患者醫院欄為可選人工內容；新列留白，不補填、不覆寫，搬列保留。** FU 使用自身醫院及日期。
- ScheduleSyncId、CalendarEventId 隨整列搬移；重複、孤立或病人重新綁定需先核對。
- 取消依病歷號的顯示刪節線判定，含富文字。
- 獨立監控涵蓋過去與未來月份及 FU／All；刀日不建事件，無日期 FU 不建事件。已綁定 FU 清空日期需核對。
- 封存保留 ID；未完成封存阻擋寫入。只有連續完整成功掃描確認缺列，才允許移除已確認擁有權的事件。

## 更新順序

1. 修改 JSON、對應實作與合成案例；新版患者欄位是否進入 Calendar 須明確定義。
2. 執行 `python tools/generate_contract.py`，核對生成差異及真實 GAS 匯出。
3. 跑 GAS 測試、`python -m vghks_sheet_core.verify --gas-repo .` 及固定主程式交叉驗證。
4. 主程式與監控固定相同完整 GAS Git commit；同步更新依賴鎖及契約 manifest，各自執行回歸。
5. GAS 部署等待相容性檢查；主程式 Release 使用成功 main 建置的原產物，監控另行發布與部署。

Python 套件隨部署保存規格，執行中不下載最新規格。兩個 Python 消費端的完整來源 commit、內嵌規格與案例皆需核對；GAS 的 consumer commit 是 CI 相容性候選，不是另一份規格來源。

新增相容能力遞增 minor，修正遞增 patch；移除欄位／別名或改變識別、取消、刪除語意遞增 major。破壞性更新先發布可讀新舊格式的消費端，再遷移資料，最後移除舊格式。
