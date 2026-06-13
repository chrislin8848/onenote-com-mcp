# Phase 6 §4 工具描述驗收工作表(真實 Claude Desktop,逐句手測)

> **規則(SPEC §4):** 在**接好 server 的真實 Claude Desktop**逐句輸入,用 §7 診斷日誌看它
> **選了哪個工具、帶什麼參數**。挑錯的回去改該條描述,每句重複幾次(有非決定性)迭代到穩。
> **不要**用 CC sub-agent / API 代測——那會污染「只看描述的天真 Claude」。本檔只是 Chris 手測的清單與對照表。

## 前置(每輪測試前)

1. **用拋棄式測試本/測試節**——Desktop 會真的執行寫入/刪除,別拿正式筆記本。
2. **開診斷日誌**(看工具選擇與參數):
   - `ONENOTE_MCP_LOG_LEVEL=DEBUG`
   - `ONENOTE_MCP_LOG_FILE`(預設 `%LOCALAPPDATA%\OneNoteMCP\logs\`;或指定一個檔)
   - 重啟 Claude Desktop 讓設定生效。
3. 測一句後,在日誌看 `tool call` 記錄的 **name + params**;填到下面對照表。
4. 同一句**重複 3 次**(模型有非決定性);若選擇不穩或選錯,回報我改描述,再重測。

判讀:**選對工具但缺/錯參數 = 等於選錯**(尤其 `delete_page_content` 缺 `object_id`)。

---

## 測試矩陣(代表性說法 → 期望工具)

每組是 §4 點名「會互相混淆」的工具。逐句輸入,回填「實際選的工具 / 帶的參數 / 穩不穩 / 判定」。

### A. 刪除:`delete_node` vs `delete_page_content`

| # | 使用者說法 | 期望工具 | 實際選 | 參數 | 穩(3次) | 判定 |
|---|---|---|---|---|---|---|
| A1 | 刪掉這個表格 | `delete_page_content` |  |  |  |  |
| A2 | 把這頁裡的這張圖移除 | `delete_page_content` |  |  |  |  |
| A3 | 刪掉整個「行程」這一節 | `delete_node` |  |  |  |  |
| A4 | 把這一頁整頁刪掉 | `delete_node` |  |  |  |  |
| A5 | 刪掉這個節群組 | `delete_node` |  |  |  |  |
| A6 | 清掉這頁的這段大綱文字 | `delete_page_content`(或拒絕→改 update) |  |  |  |  |

### B. 寫內容:`update_page_content` vs `create_table` vs `insert_image`

| # | 使用者說法 | 期望工具 | 實際選 | 參數 | 穩(3次) | 判定 |
|---|---|---|---|---|---|---|
| B1 | 在這頁最後加一段「明天行前說明」 | `update_page_content`(append) |  |  |  |  |
| B2 | 把這段文字改成紅色 | `update_page_content`(replace) |  |  |  |  |
| B3 | 幫我在這頁加一個 3 欄的表格 | `create_table` |  |  |  |  |
| B4 | 在這個表格再加一列 | `update_page_content` 或 `create_table`(append rows) |  |  |  |  |
| B5 | 把這張照片插進這一頁 | `insert_image` |  |  |  |  |
| B6 | 在這段話前面插一行小標 | `update_page_content`(insert_before) |  |  |  |  |

### C. 整理結構:`restructure_section` vs `reorder_sections` vs `move_page` vs `rename_node`

| # | 使用者說法 | 期望工具 | 實際選 | 參數 | 穩(3次) | 判定 |
|---|---|---|---|---|---|---|
| C1 | 把這一節裡的頁面重新排序 | `restructure_section` |  |  |  |  |
| C2 | 把這頁變成那頁的子頁 | `restructure_section`(pageLevel) |  |  |  |  |
| C3 | 把這本筆記本的節順序調一下 | `reorder_sections` |  |  |  |  |
| C4 | 把這一頁搬到「已完成」那一節 | `move_page` |  |  |  |  |
| C5 | 把這一節改名叫「2026 行程」 | `rename_node` |  |  |  |  |
| C6 | 把這頁標題改成「報名表」 | `rename_node` |  |  |  |  |
| C7 | 整理一下這個亂掉的筆記本 | (應提案+問範圍,可能 restructure/reorder) |  |  |  |  |

### D. 讀取:`get_page` vs `get_page_images` vs `get_page_files_info` vs `get_page_files`

| # | 使用者說法 | 期望工具 | 實際選 | 參數 | 穩(3次) | 判定 |
|---|---|---|---|---|---|---|
| D1 | 這一頁寫了什麼 | `get_page` |  |  |  |  |
| D2 | 把這頁的圖片給我看 | `get_page_images` |  |  |  |  |
| D3 | 這頁附了哪些檔案 | `get_page_files_info` |  |  |  |  |
| D4 | 看一下這頁那個 Excel 附件裡有什麼 | `get_page_files_info`(→不支援) |  |  |  |  |
| D5 | 讀一下這頁附的那個 PDF 內容 | `get_page_files` |  |  |  |  |
| D6 | 這頁附的 txt 寫了什麼 | `get_page_files` |  |  |  |  |

### E. 其餘工具(覆蓋全套 22)

| # | 使用者說法 | 期望工具 | 實際選 | 參數 | 穩(3次) | 判定 |
|---|---|---|---|---|---|---|
| E1 | 我有哪些筆記本 | `list_notebooks` |  |  |  |  |
| E2 | 列出這本的所有節 | `list_sections` |  |  |  |  |
| E3 | 這一節有哪些頁 | `list_pages` |  |  |  |  |
| E4 | 找一下提到「峇里島」的頁面 | `search_pages` |  |  |  |  |
| E5 | 我現在在哪一頁 | `get_current_context` |  |  |  |  |
| E6 | 在這本新增一節叫「待辦」 | `create_section` |  |  |  |  |
| E7 | 在這一節開一張新頁叫「會議記錄」 | `create_page` |  |  |  |  |
| E8 | 把這一頁複製到「封存」那一節 | `copy_page` |  |  |  |  |
| E9 | 把整節複製到那個節群組底下 | `copy_section` |  |  |  |  |

### F. 行為契約(描述文字裡的契約是否被遵守)

| # | 觀察點 | 預期 | 結果 |
|---|---|---|---|
| F1 | 破壞性工具(delete_*、覆蓋式 update)是否**先提案/確認**才動手 |  先問再做 |  |
| F2 | 結構性工具(restructure/reorder/move/rename)是否提案+建議先 copy 備份 | 提案+建議備份 |  |
| F3 | 刪頁面內物件前,是否**先讀**(get_page/images/files_info)拿 `object_id` | 先讀拿 ID |  |
| F4 | server `instructions` 是否被採用(範圍化讀取、兩步 objectID 規則) | 觀察是否生效 |  |

---

## 輸出:穩定後的「說法 → 工具」對照表

(迭代到穩後,把上表彙整成最終對照表,作為 §4 驗收交付物。)

| 說法 | 最終穩定選擇 | 備註/改過的描述 |
|---|---|---|
|  |  |  |
