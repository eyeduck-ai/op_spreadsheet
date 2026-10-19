# 手術排程與 FU 追蹤

GAS **2026.10.19.5**，共用契約 **1.2.2**。

| 專案 | 負責範圍 |
| --- | --- |
| [vghks-bot](https://github.com/eyeduck-ai/vghks-bot) | 院內 SOAP、刀表預覽／更新、統計及封存 |
| 本 repo | 線上刀表格式、排序、月份／刀日、FU 跨表操作及封存 |
| [vghks-sheet2cal](https://github.com/eyeduck-ai/vghks-sheet2cal) | 單一主機常駐同步今天起的排程，過去改期沿用原事件 |

## 院區與日常操作

- 月份以 `YYYYMM` 命名，依第一列表頭定位欄位。
- 刀日需有完整日期、`高榮` 或 `聯醫`、`◆ 刀日`，且 CalendarEventId 留白。同一天可有不同院區區塊。
- **患者的醫院欄可留白或手填。日期與院區依最近上方有效刀日判斷；新增患者不補院區，更新及搬列保留既有值與公式。** FU 使用自己的醫院及日期欄。
- 病人日期欄是報到時間；整列搬移、剪貼或排序需包含隱藏的 ScheduleSyncId、CalendarEventId。
- 選單提供建立月份／刀日、整理分頁、水晶體清單、月表與 FU 雙向轉送、年度封存。
- 「修復刀表格式」只整理既有分頁的格式、驗證及編輯觸發器，不建立 FU、不配置既有 ID、不批次改寫報到時間或患者院區。

## 識別、保護與封存

| 系統欄 | 用途 |
| --- | --- |
| ScheduleSyncId | 每筆患者及刀日的永久識別，整列搬移時保留 |
| CalendarEventId | 外部監控的事件綁定；GAS 不建立或清除 |

系統欄自動隱藏；既有保護範圍、自訂欄、註記、公式與人工資料保留。重複 ID、錯誤刀日及混合欄位需先核對，不推測綁定。既有資料的系統欄準備由主程式或獨立監控執行，不需要 GAS 初始化或交接選單。

年度封存採「預覽 → 私人整檔備份 → 原生複製與讀回核對 → 再次確認移除來源」，保留格式、刪節線、ID 及恢復計畫。使用 `vghks.sheet.archive.lock` 互斥；未知結果先核對原交易，不重送或自動解除。既有舊格式轉換列於預覽；不調整人工保護。

## Calendar 與同步範圍

GAS 不呼叫 Calendar。獨立監控預設每 60 分鐘掃描所有受支援月份與 FU／All，核對全部排程 ID，但只同步**目前日期為今天或未來**的排程；過去改到今天／未來時更新原事件，今天／未來改到過去時更新原事件一次，確認完成後略過。待核對操作失敗或重啟後續接，不補建過去事件。IVI 不納入，刀日標題不建立事件。FU 無日期不建立事件，封存 ID 用於保護既有事件。詳細設定與維運見獨立 repo。

舊的已安裝觸發器只保留必要格式轉接或安全空操作，避免其他安裝者的舊觸發器報錯；不保留日曆設定、遷移或交接功能空殼。

## 部署與驗證

使用既有綁定專案與 GitHub secrets。合併至 master 後，Actions 先驗證規格、真實 GAS 匯出及固定主程式版本，再執行 clasp push，唯讀確認正式刀表綁定。更新後重新整理試算表載入選單。

```bash
python tools/generate_contract.py --check
node --check code.js
node --check tests/smoke.js
node tests/smoke.js
node tests/sheet_contract.js
node tests/parser_contract.js
node tests/workflow_contract.js
node tests/monitor_handover.js
node tests/archive_contract.js
node tests/formal_contract.js
git diff --check
```

正式 JSON、共用 Python 套件及更新流程見 [共用規格](docs/SHEET_CONTRACT.md)。公開案例全部為合成資料。
