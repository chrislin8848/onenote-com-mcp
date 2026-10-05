# OneNote MCP Server — 開發規格 (給 Claude Code)

> 這份文件是**規格**,不是完成的計畫。請先讀完、提出釐清問題,然後**產出一份分階段的開發計畫**(見最後一節)。

---

## 1. 目標

建立一個 **COM-only** 的 OneNote MCP server,讓 Claude(在 Windows 上的 Claude Desktop)能對 OneNote 桌面版做完整 CRUD:

- 讀取:筆記本/節/頁清單、頁面文字、**表格(保留結構)**、**圖片(二進位)**、文字搜尋、**使用者目前檢視位置**(目前停留的筆記本/節/頁)、**附件/嵌入物件**(中繼資料一律可得;內容抽取**僅限**文字類/圖片/PDF,見 §5「附件與嵌入物件」)。
- 寫入:**新增節/頁**、**附加或就地修改內容(含表格)**、**新增表格**、**插入圖片**。(COM 不能在本地新增筆記本,見 §4/§5。)
- 刪除:刪頁/節/筆記本(層級節點);刪頁面內容物件——圖片、表格、大綱。
- 複製:忠實克隆節/頁(保留格式、表格、圖片、**附件/嵌入物件**、子頁階層);搭配既有編輯工具做「複製後改寫」(見 §4 複製工具、§5)。**無整本克隆**(COM 不能建本);要克隆整本須先手動建好目標本再逐節 `copy_section`(見 §5)。
- 整理:重排頁/節順序、調整頁面階層(`pageLevel`)、改名、跨節搬頁——支援「請 Claude 重整雜亂筆記本」情境(見 §4 結構工具、§5 層級重排紀律)。

文字一律**完整保真**(字型/字級/顏色、粗斜底、螢光等):讀得出、編輯不擾動既有格式、可對新內容設定格式(規則見 §5「格式保真」)。

資料一律走**即時** OneNote 桌面 COM API(不解析備份檔),因此沒有延遲、能拿到圖片與結構化表格、且完全不需要 Azure app 或 token(COM 跑在使用者自己的 OneNote session 內)。

### 1.1 參考來源 / Prior Art:`mhzarem/onenote-mcp`

Repo: <https://github.com/mhzarem/onenote-mcp>。**clone 它當參考,但不是拿來擴充的基底**;本專案在「讀取來源」與「執行模型」上是刻意偏離它的。

可取用:

- **工具命名與介面慣例** — `list_notebooks` / `list_sections` / `list_pages` / `create_page` 等命名沿用自它(本規格 §4 已採用並擴充;注意本專案已把它的 `append_to_page` 併入 `update_page_content` 的 append 模式)。
- **MCP / FastMCP 樣板** — `server.py` 怎麼註冊工具,可當最小可動範例參考。
- **建立/附加頁面的 OneNote XML payload 形狀** — 可當 `CreateNewPage` / append 的快速起點;但完整 schema、表格、`UpdatePageContent`、`GetBinaryPageContent` 一律以 §5 的 Microsoft 文件為準(比它完整,它只做 create + append)。

刻意**不**取用(這是本專案與它最大的差異):

- **它的讀取走「解析本機備份 `.one` 檔(pyOneNote)」** — 本專案明確改走即時 COM(理由見 §1):備份檔有延遲、純文字、解析不完整。**不要沿用備份解析路徑。**
- **它的寫入用「PowerShell 子行程驅動 COM」** — 本專案改用 in-process **pywin32**,並收在 §3 的 `OneNoteBackend` 介面後面。**不要沿用 PowerShell subprocess 模型。**

一句話:把 mhzarem 當作「一個可動的最小 OneNote MCP 長什麼樣」來參考其**工具表面與 MCP 接線**,讀寫的**實作一律以本規格為準**。

### 1.2 與現有方案的定位差異(為何不直接用現成的)

公開的 OneNote MCP server(eshlon/onenotemcp、ZubeidHendricks/azure-onenote-mcp-server、purpleslurple/onenote-mcp-server 等)**幾乎清一色走 Microsoft Graph API**,因此都背著同一組包袱:

- **要 Azure app 註冊 + token**(tenant/client/secret 或 device code flow)。本專案走 in-process COM,**完全不需要 Azure/token**,且天然分用戶(各員工只看到自己同步的 OneNote)。
- **全域搜尋撞規模牆。** eshlon 那一版正是先前在你 55+ 筆記本下放棄的主因。本專案以 `FindPages` + notebook/section 範圍化避開(見 §4 範圍化原則)。
- **Graph 的 OneNote HTML API 對「完整格式保真」與就地編輯本來就有損、受限。** 那些專案宣稱的 comprehensive editing 範圍是「所有 *Graph* 操作」,比 COM 能碰的小一圈。本專案直接拿原生頁面 XML,才做得到螢光雙屬性、`QuickStyleDef`、手術式就地編輯不走有損 round-trip、整頁忠實克隆。

商用/SaaS(Merge、CData、Zapier)競爭的是**另一個維度**——企業治理、DLP、多來源整合、免自架,而非 OneNote 內容操作的深度與保真度;那些不是本專案目標。Merge 雖有「複製筆記本」,但那是非同步 Graph 操作,與本專案「COM 整頁 XML 忠實克隆 + 複製後改寫」不是同一層級。

**結論:** 選 COM 不是繞遠路,而是 Graph 路線結構上做不到「完整保真 + 範圍化避牆 + 免 Azure 分用戶」這三者的交集——這正是本專案的存在理由。

---

## 2. 環境與測試架構(單機:Linux host + Windows VM)

整套開發與測試在**同一台實體機**完成:一台 bare-metal Linux server(Intel W-1290,支援 VT-x/VT-d,非巢狀虛擬化,KVM 近原生速度)。Linux host 負責開發;host 內再用 KVM 跑一個 Windows 11 guest 當 COM 測試標的。

| | Linux host | Windows 11 guest (KVM VM) |
|---|---|---|
| 角色 | Claude Code、開發、repo | OneNote + pywin32 + COM |
| COM 可用? | **否** — 無 pywin32/OneNote | 是 |
| 測試層 | Tier 1:純邏輯 + 單元測試(對 fixtures),每次改動都跑 | Tier 2:真實 COM 整合測試,檢查點才跑 |

兩者透過 libvirt 虛擬網路在本機互通(host ↔ guest),不經外網,ferry 全在這台機器內。

### 2.1 不可違反的設計紀律

即使 VM 就在本機,CC 在 Linux host 仍**不能直接跑 COM**(COM 只在 guest 內,且需互動桌面 session,見 §2.3)。因此下列紀律照舊,目的是讓 host 的內層迴圈快、不必每次改動都喚醒 VM:

1. **COM 隔離在介面後面。** 抽象 `OneNoteBackend`,兩個實作:`Win32ComBackend`(pywin32,僅 guest)、`FixtureBackend`/`MockBackend`(回放 fixtures,host)。
2. **pywin32 守護匯入。** 不可在 module 頂層 `import win32com`;lazy import 或 try/except,確保套件在 Linux host 上 `import` 不炸。
3. **所有 XML 解析/組裝寫成純函式。** `one:` 解析與 `UpdatePageContent`/`CreateNewPage` payload 組裝不得夾帶 COM,必須能在 host 對 fixtures 完整單元測試。
4. **整合測試標 `@pytest.mark.windows`,在 host 自動 skip。**

### 2.2 Windows VM 建置(一次性)

- KVM/QEMU + libvirt(virt-manager 圖形管理)。
- Guest:Windows 11,配虛擬 TPM(swtpm)+ OVMF/UEFI;裝 virtio 磁碟/網路驅動;磁碟用 qcow2、thin-provision ~120GB(支援快照,測前回乾淨狀態)。
- 資源:4 vCPU / 8GB RAM 起步(host 24GB 裡留 ~16GB;OneNote 同步吃緊再拉到 10–12GB)。
- 網路:libvirt 預設 NAT 即可,一條同時供 guest 連網登入 M365 / 同步 OneNote,以及 host 從虛擬子網 SSH 進 guest。
- guest 內安裝:Python 3.12+、pywin32、git、OneNote 桌面版(登入 M365、設開機啟動、保持同步)、OpenSSH Server(設開機自動啟動,配好 host→guest 金鑰免密碼)。
- **autologon**:設 Windows 自動登入,讓 guest 開機就停在已登入互動桌面——這是 COM 能運作的前提。

### 2.3 遠端操作 VM 畫面

兩種都掛在現有 SSH 上,不對外開埠:

- **主控台(SPICE/VNC over `qemu+ssh`)** — 用 virt-viewer / virt-manager 連 `qemu+ssh://user@host/system`。安裝階段唯一可用(guest 還沒網路時),且顯示的就是 autologon 主控台 session(COM 跑的那個)。建 VM、裝 Windows/Office、設 autologon、首次登入 OneNote 全程用它。
- **RDP(裝好後可選,較順)** — `ssh -L 3389:guest_ip:3389 user@host` 後連 `localhost:3389`。注意 RDP 會接管/鎖定 session;**跑 COM 測試的常駐 session 固定在 autologon 主控台**,RDP 只做臨時管理,動完確認測試在 RDP 斷線後仍通過。

### 2.4 自動化測試迴圈

**Tier 1(host,每次改動):** CC 跑純邏輯單元測試,不碰 VM。

**Tier 2(VM,檢查點):** CC 在 host 跑一支 `scripts/remote_test.sh` 觸發整串:

1. host 把程式碼送進 guest(git push → guest pull,或 rsync)。
2. host 觸發 guest 執行測試(見下「觸發 vs 執行」)。
3. guest 在**互動 session** 內用 pytest 跑 `@pytest.mark.windows`,對真實 COM 做 create → get → update(含改表格)→ delete round-trip。
4. 結果(pass/fail、log、JUnit)+ 重新抓的 fixtures 回送 host。
5. 腳本以 guest 的 exit code 為自己的 exit code,CC 據此迭代。

**觸發 vs 執行(必須拆開,因為 COM 要互動 session):** SSH 進 guest 落在非互動 session,直接 spawn 會讓 COM 喚 OneNote 失敗。所以——

- **觸發**走 SSH/git(非互動,OK)。
- **執行**落在 autologon 互動 session,二選一:
  - Task Scheduler 工作設「只在使用者登入時執行」,SSH 只下 `schtasks /run` 觸發;或
  - 一支隨 autologon 啟動的常駐 runner,顧一個佇列(資料夾 / git branch / 本機 HTTP),收到工作就在桌面跑 COM、寫回結果。

第 4 步重抓 fixtures 會順手更新 host 的單元測試樣本,所以 Phase 1 不會因「缺樣本」卡住。

**交付物:** `scripts/remote_test.sh`(host)、互動 session 執行器(Task Scheduler 設定或常駐 runner,guest)、`scripts/dump_fixtures.py`(guest,首次抓樣本)。

---

## 3. 架構分層

```
┌─────────────────────────────────────────┐
│ MCP layer  (FastMCP 工具定義/註冊)        │  Linux 可測(工具 schema、參數驗證)
├─────────────────────────────────────────┤
│ Service layer (流程編排、輸入驗證)        │  Linux 可測
├─────────────────────────────────────────┤
│ OneNote XML layer (純函式)                │  Linux 可測 (核心,對 fixtures TDD)
│  - parse: GetHierarchy/GetPageContent XML │
│  - build: CreateNewPage/UpdatePageContent │
├─────────────────────────────────────────┤
│ Backend layer (OneNoteBackend 介面)       │
│  - Win32ComBackend  (pywin32, Windows)    │  僅 Windows 可測
│  - FixtureBackend   (回放, Linux)         │  Linux
└─────────────────────────────────────────┘
```

執行/消費端:MCP server 必須跑在有 OneNote 的 Windows 上,且通常由 MCP client(Claude Desktop 或 Antigravity)在本機以子行程啟動,所以 client + server + OneNote 同機。

- **開發測試期** — 跑在 §2 的 Windows VM(需要 autologon + 互動 session 執行器,那是測試環境的需求)。
- **正式部署** — 跑在**每位員工自己的工作 Windows PC**。該 PC 本就有員工登入的互動桌面,所以正式環境**不需要** autologon / 互動 session 執行器那套;每台只看得到該員工自己帳號、自己同步的 OneNote(天然分用戶,無共用憑證)。由各員工執行一個 installer EXE 完成安裝(見 §8),像裝一般 Windows App;不做集中式自動推送(MDM/GPO)。

若 client 用 Claude Desktop(Microsoft Store 版),設定檔在 `...\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json`,以 Developer → Edit Config 確認。本專案支援的消費端 client 為 **Claude Desktop 與 Antigravity(CLI/IDE)**;各自設定檔路徑與一鍵 `--configure` 見 §8。Claude Code(Linux host)只是開發工具,不是執行環境。

---

## 4. 要實作的 MCP 工具

| 工具 | 用途 | 對應 COM 方法 |
|---|---|---|
| `list_notebooks` | 列所有筆記本 | `GetHierarchy(scope=hsNotebooks)` |
| `list_sections` | 列某本的節(**含節群組巢狀結構**) | `GetHierarchy(notebookId, hsSections)`(回傳含 `one:SectionGroup` 巢狀,解析須保留,不另開工具) |
| `list_pages` | 列某節的頁(**範圍化,含 `pageLevel` 子頁階層**) | `GetHierarchy(sectionId, hsPages)` |
| `search_pages` | 跨/範圍文字搜尋 | `FindPages` |
| `get_page` | 取單頁內容(**保真富文字** + 結構化表格);**`text_only=True`(鼓勵預設使用)** 只回文字(段落=字串、表格=二維文字陣列、圖=OCR 文字、附件=檔名),無 runs/樣式/objectID,小 10–100 倍 | `GetPageContent(pageId)` |
| `get_table` | 取**單一表格**,不含整頁其餘。**`text_only=True`(鼓勵預設使用)** = 每格文字二維陣列(106×39 實表 2.9M→21.9K 字元);預設完整模型(每格文字+樣式+底色 + row/cell/段落 objectID)在大表**並不精簡**(≈ `get_page`)。兩種都可用 `start_row`/`max_rows`/`columns` 分段讀;≥1,000 格附 `write_cost` 提示(該頁顯示中時另註明) | `GetPageContent` → 依 objectID 找該 `one:Table`(含巢狀)投影 |
| `get_object` | 取**單一物件**(段落完整文字+runs+樣式,或表格/圖片/附件)by objectID,不含整頁其餘——比 `get_page` 精簡(不含整頁重複 runs 與全頁樣式表),找得到頁面任何位置(行內/巢狀於 cell/頁層);`text_only=True`(鼓勵)只回文字 | `GetPageContent` → 依 objectID 定位並投影該物件 |
| `find_objects` | 取頁面上**文字含某子字串**的物件 objectID(**頁內**定位,對應 `search_pages` 的「找整頁」)——比對**完整文字**(非 `get_page_info` 的 40 字截斷預覽),用來精準定位要改的段(如錯字所在) | `GetPageContent` → 走訪段落比對全文 |
| `get_page_images` | 取頁面圖片(二進位) | `GetPageContent` → `GetBinaryPageContent(callbackId)` |
| `get_page_files_info` | 列頁面**附件/嵌入物件**(如 Excel 試算表)中繼資料:名稱/大小/型別/objectID,**不解析內容** | `GetPageContent` 解析 `one:InsertedFile` + 讀 `pathCache` 檔案屬性 |
| `get_page_files` | 取**附件內容**供 Claude 分析(**僅限**:文字類解碼/圖片/PDF 抽文字;其餘型別回中繼資料並明示不支援) | 讀 `pathCache` 本機快取檔(**非** `GetBinaryPageContent`)+ server 端型別感知抽取 |
| `get_current_context` | 取使用者目前檢視位置(筆記本/節群組/節/頁,ID+名稱) | `Windows.CurrentWindow` 的 `CurrentNotebookId/CurrentSectionGroupId/CurrentSectionId/CurrentPageId` + 範圍化 `GetHierarchy` 解名稱 |
| `create_section` | 在指定本(或節群組)建立節 | `OpenHierarchy(name+".one", parentId, out id, cftSection)`(parent 可為 notebook 或節群組) |
| `create_page` | 在指定節建新頁(可設 `pageLevel` 子頁);**預設把新頁放在目前所在頁(get_current_context)正下方**,`after_page_id` 可指定別頁,無視窗/目前頁不在該節則落節尾 | `CreateNewPage` (+ `UpdateHierarchy` 設 pageLevel) + `UpdatePageContent` + 預設接 `reposition_page`(錨點=目前頁;讀視窗在建頁前;錨點不在該節→吞掉退回節尾) |
| `update_page_content` | 改頁面內容:append / insert / replace(含改表格儲存格文字、樣式大小/字型/顏色/底色、超連結 `<a href>`) | `GetPageContent` → 改 XML → `UpdatePageContent`(純 append 可免讀全頁) |
| `find_and_replace` | **就地**替換文字 `find`→`replace`(全頁或單一物件範圍),逐 run 保留各自樣式——改錯字/小幅改字最省做法,不必重供整段;回傳替換次數與變動 objectID。跨 run(不同樣式)的命中不替換,改回報於 `found_across_runs`(用 `get_object` + `update_page_content "replace"` 收尾) | `GetPageContent` → 逐 `one:T` 內逐 span 替換 → `UpdatePageContent`(無命中則跳過寫入) |
| `batch_update` | 對一頁套用**多筆**內容編輯於**單一原子寫入**(全成或全不寫;一次 round-trip 取代連發 update_page_content):ops = `replace`/`append`/`insert_before`/`insert_after`/`find_replace`;各 op 針對讀取時已存在的 objectID;`return_ids` 可回傳新建 objectID 與新 last_modified_time | 多個 mutator 套同一棵樹 → **一次** `UpdatePageContent`(單一寫核心,不另開呼叫點) |
| `create_table` | 新增**新**表格(僅建立) | 組 `one:Table` XML → `UpdatePageContent` |
| `modify_table` | 改**既有**表格:**形狀**(`insert_rows`/`insert_columns`(`insert_columns` 可帶 `values` 一次填新欄內容)、`delete_rows`/`delete_columns`(DESTRUCTIVE);列欄對稱;空 cell 補最小段落)、**順序**(`reorder_columns`/`reorder_rows`,傳**完整目標排列** order、cell 隨欄/列移動且保留 objectID,不重打字)、或**內容**(`set_rows` 覆蓋整列/整表;`set_column` 精簡覆蓋**單欄**(一格一值);皆固定維度、保留每格 objectID,cell 給 `None` 不動,`[None,"",…]` = 保留第一欄、清空其餘) | `GetPageContent` → 改 `one:Table`(Columns 重編 index、移動/增/刪/改 Cell)→ `UpdatePageContent` |
| `insert_svg_image` | 從 **SVG markup** 插入向量圖(伺服器端光柵化成 PNG)| `resvg_py` 渲染 SVG→PNG → 組 `one:Image` + base64 `one:Data` → `UpdatePageContent`(僅吃向量;內嵌 raster data: URI 拒絕;照片手動插)|
| `insert_image_from_path` | 從**本機檔路徑**插入點陣圖(PNG/JPEG/GIF)——server 端讀檔 bytes,**不經模型**(舊 base64 插入因模型吐 bytes 過慢已移除);適合使用者指定的照片,或能寫檔/跑 code 的 client(如 agentic IDE)在磁碟上產出的圖(程式畫的圖表、下載的圖)。僅 PNG/JPEG/GIF;只存在 context、未落地的圖無法插 | 讀本機檔 → 組 `one:Image` + base64 `one:Data` → `UpdatePageContent`(沿用 `insert_svg_image` 同一 seam) |
| `delete_node` | 刪頁/節/節群組/筆記本(層級) | `DeleteHierarchy(objectId)` |
| `delete_page_content` | 刪**頁層**內容物件(整個大綱/頁層圖片/頁層附件) | `DeletePageContent(pageId, objectId)` |
| `delete_inline_content` | 刪**大綱內**物件(表格/段落/行內圖片/行內附件);保留同大綱其他段落 | 走編輯 seam:`GetPageContent` → 移除元素並修剪空容器 → `UpdatePageContent`(**非** `DeletePageContent`——COM 對行內 OE 一律拒絕 `0x8004200E`) |
| `copy_page` | 忠實克隆**單頁**到目標節;**預設把副本放在來源頁正下方**(同節複製的自然位置),`after_page_id` 可指定放在別頁之後,跨節複製則落在節尾 | 內部 raw-XML 克隆(見 §5「複製/克隆」)+ 預設接 `reposition_page`(錨點=來源頁;跨節時錨點不在目標節→吞掉退回節尾) |
| `copy_pages` | 忠實克隆**多頁**為**一個連續區塊**(依給定順序、各頁保留 pageLevel)到目標節;`after_page_id` 把整塊放在某頁之後,否則落節尾 | 逐頁 `copy_page`(各落節尾)+ **一次** `reposition_pages` 把整批排成連續區塊到錨點後(node-ID 守恆,whole-batch 紀律由 seam 保證) |
| `copy_page_subtree` | 忠實克隆**一頁連同其子頁**(其後較深 pageLevel 的連續頁)為一個區塊;**同節預設放在來源子頁樹正下方**(像 copy_page),`after_page_id` 可指定別頁,跨節落節尾 | `subtree_page_ids`(依位置式子頁模型算出清單)→ `copy_pages` |
| `copy_section` | 忠實克隆**整節** | 建節 + 逐頁 `copy_page`(保留 pageLevel) |
| `restructure_section` | 同節內**整批**重排**多頁**順序 + 調整 `pageLevel` 階層 | `GetHierarchy` → 重排完整頁清單(每頁帶 `pageLevel`)→ `UpdateHierarchy`(紀律見 §5「層級重排」) |
| `reposition_page` | 把**一頁**移到同節內某頁之後(空=移到節首)+ 可選設 `pageLevel`;只給 ID,不必交完整清單 | 走同一個 hierarchy seam:`GetHierarchy`(節範圍)→ `addnext`/`addprevious` 原地搬一個元素(node-ID 守恆,whole-batch 紀律由 seam 保證)→ `UpdateHierarchy` |
| `reorder_sections` | 重排某本內節的順序 | 同上,對 `one:Section` 元素整批重排(僅節/頁有實證;**筆記本層級排序未驗證、不納入**) |
| `rename_node` | 重新命名頁/節 | `UpdateHierarchy`(改 name 屬性) |
| `move_page` | 把頁搬到別的節 | `UpdateHierarchy`(把頁元素掛到目標節下;**跨節搬移可靠性待 VM 實機驗證**) |

**範圍化原則:** 永遠用 notebook/section 範圍的 `GetHierarchy` 與 `FindPages`,**不要**對全部頁面逐頁掃描——這是先前 Graph 全域搜尋撞牆的同一個坑,逐頁打 COM 又慢又容易踩 RPC 忙碌錯誤。

**工具與 COM 方法是多對一(勿重寫):** `update_page_content`(含 append/insert/replace)、`create_table`、`insert_svg_image`/`insert_image_from_path`、`find_and_replace`、`batch_update` 最後全都走同一個 `UpdatePageContent`——它是 OneNote 唯一的「寫內容進頁面」通用原語。實作上必須**共用單一「套用頁面內容變更」核心**(集中處理 `GetPageContent → 改 XML → UpdatePageContent` 的並發保護 `dateExpectedLastModified` 與手術式格式保真),這些 MCP 工具只是它的薄 facade(各自負責把自己的輸入轉成內容 XML)。**不要各自重寫一套 round-trip 邏輯**,否則程式會分歧、難維護。分開命名是為了 LLM 好用,不是要分開實作。

**大內容框寫入(v1.4.0,VM 實測 2026-10-05):** `UpdatePageContent` 對一個 `one:Outline` 是**整框取代**(框內無 objectID 合併——少送的段落/列會被刪、只送一格的列被拒),故改一格 = 重送整個內容框;成本按物件數(約每格 10ms)計、與 payload 大小無關。**該頁正顯示於 OneNote 時慢 5–6 倍**(4,134 格:134–141s vs 23–26s)。對策:寫入核心在 payload ≥1,000 個 OE 且該頁為 `CurrentWindow.CurrentPageId` 時拒寫(`PageDisplayedError`,請使用者切到別頁;`allow_displayed=True` 可覆寫,9 個內容寫入工具皆有此參數);指引要求大表的所有修改合併成**一次** `batch_update`/`set_rows`/`set_column`,逾時的寫入多半已完成、先 `get_table(text_only)` 確認再重試。讀取側:`get_page`/`get_table`/`get_object` 的 `text_only=True` 為鼓勵預設;`get_page_info` 對 ≥200 格的表格只列一筆摘要(`cell_paragraphs_not_listed`),格內圖片/附件/巢狀表格照列,`include_cells=True` 全展開(實表 637K→640 字元)。

### 工具描述強化(獨立交付物,跨全套工具一次做,非逐工具帶過)

工具描述是**消費端 Claude 選工具用的 UX 文字**,不是給實作者的程式碼註解。它是一個**獨立交付物**,須在全套工具大致到齊後**一次跨全套做**,不可散在各 Phase 逐工具帶過——因為「對比式描述」需要同時看著整組工具、推理哪幾對會混淆,這不在逐一實作的局部視野裡。FastMCP 從 docstring 取描述,預設會給「功能正確但局部、通用」的描述,**不會自動**寫出讓選擇變準的對比界線與行為契約,故須明確當成一項工作來做。20+ 工具的選擇可靠度,瓶頸不在數量而在「哪幾組會被搞混」。

**原則:寫得「分明」不是「冗長」。** 每條描述每次呼叫都吃 context、且影響選擇,要密集、有辨識度、互相劃清界線。每條至少回答:做什麼、何時用、**何時「不要」用(改用哪個工具)**。

**三件必做(功能描述正確之外):**

1. **對比式/反向標註** — 針對本專案會互相混淆的工具組,在描述裡互相點名界線:
   - **刪除三角:** `delete_node`(刪整個頁/節/節群組/筆記本節點) vs `delete_page_content`(刪**頁層**物件:整個大綱/頁層圖/頁層附件,頁面保留) vs `delete_inline_content`(刪**大綱內**物件:表格/段落/行內圖/行內附件)。關鍵界線:**整個表格、段落永遠在大綱內,故刪它們一律用 `delete_inline_content`,絕不用 `delete_page_content`**;三者描述互相點名(both-way,guard-tested)。
   - `update_page_content`(加/改文字、樣式、超連結、**單一**儲存格文字) vs `create_table`(建**新**表格) vs `modify_table`(改既有表格:`insert_rows`/`insert_columns`/`delete_*` 形狀、`reorder_columns`/`reorder_rows` 重排、或 `set_rows`/`set_column` 覆蓋內容) vs `insert_svg_image`(從 SVG 插**向量**圖;照片/點陣須手動)——`update_page_content` 描述須註明「若加的是表格、改的是行列數、或插圖,改用對應工具」;`create_table` 只建新表、`modify_table` 改既有表(形狀/順序/內容),兩者互相點名;表格內容替換的分工 = `update_page_content "replace"`(一格)vs `modify_table set_column`(整欄)vs `modify_table set_rows`(整列/整表);**重排欄/列用 `reorder_columns`/`reorder_rows`,不要用 `set_rows` 清空再重打**;改整欄**樣式/顏色**(非文字)用 `apply_text_style(columns=[j])`。改錯字/小幅改字用 `find_and_replace`(逐 run 替換、保留樣式、不必重供整段;跨 run 命中回報於 `found_across_runs`);一頁有多筆編輯時用 `batch_update`(單一原子寫入,取代連發 `update_page_content`);寫入後可帶 `return_ids` 取回受影響/新建的 objectID,免重讀整頁。
   - `reposition_page`(**同節內**把**一頁**移到某頁之後,只給 ID) vs `restructure_section`(**同節內**重排**多頁** + `pageLevel`,須完整清單) vs `reorder_sections`(**一本內**節順序) vs `move_page`(把頁搬到**別節**) vs `rename_node`(只改名)。關鍵:移**單一**頁用 `reposition_page`(不必交 46 筆清單,避免模型去寫外部暫存檔);`copy_page` / `create_page` **預設**都把新頁放在「自然錨點」正下方——`copy_page` 在**來源頁**之下、`create_page` 在**目前所在頁**(get_current_context)之下,所以「複製這頁」「在這裡建頁」都不必指定位置,`after_page_id` 才是覆蓋;要移**既有**頁才用 `reposition_page`。
   - **複製粒度四選一:** `copy_page`(**單頁**) vs `copy_pages`(**多頁**明確清單) vs `copy_page_subtree`(**一頁 + 其子頁**,自動算出子頁) vs `copy_section`(**整節**)。關鍵界線:複製**多頁到某位置**時**絕不**重複呼叫 `copy_page`——`copy_page` 預設把每份副本貼在**各自來源頁**之下,會把整組副本打散(real-Claude-Desktop 實測:「把 ●ITIN 及其子頁複製到 ●Local 下方」因此亂放);`copy_pages` / `copy_page_subtree` 把副本排成**一個連續區塊**、一次定位到 `after_page_id` 之後。
   - `get_page`(文字 + 結構化表格;大表頁可改用 `get_table` 只取單表、避免巨大 payload) vs `get_table`(**單一**表格的結構化內容) vs `get_page_images`(取圖片二進位供視覺辨識) vs `get_page_files_info`(附件/嵌入物件**中繼資料**,任何型別) vs `get_page_files`(附件**內容**抽取,僅文字類/圖片/PDF)——info 是 files 的前置;非支援型別(docx/xlsx 等)只能取 info,不能取內容。讀**單一**物件完整內容(段落文字+樣式/表格/圖片)用 `get_object`(非整頁);要在**頁內**用文字定位該改哪一段用 `find_objects`(比對全文、回 objectID;`search_pages` 只回整頁、不回頁內位置)。
   - 命名小疙瘩:`restructure_section` 與 `reorder_sections` 的動詞/單複數不一致,若描述尚未對外凍結可考慮統一,降低模型猶豫。
2. **把程式碼強制不了的行為契約寫進描述文字**(那是唯一落地處):
   - 破壞性工具(`delete_node`、`delete_page_content`、`delete_inline_content`、`modify_table` 的 `delete_rows`/`delete_columns`、覆蓋式 `update_page_content`)醒目標 **DESTRUCTIVE**。
   - 結構性工具(`restructure_section` / `reposition_page` / `reorder_sections` / `move_page` / `rename_node`)註明「呼叫前先向使用者提案並取得確認;建議先 `copy_section` 備份」(呼應 §5 層級重排紀律的 propose-confirm)。
3. **參數 schema 也要導引選擇:** `update_page_content` 的 append/insert/replace 做成 **enum 並逐值描述**;必填/選填清楚;參數名自解釋。

**Server 層 `instructions`(跨工具總則):** 在 MCP `initialize` 的 `instructions` 放不屬於任何單一工具的總則——本 server 操作 live OneNote、讀取一律範圍化、**刪頁面內物件前先用 `get_page` / `get_page_images` / `get_page_files_info` 取 `objectID`**、以及標竿流程「日期改寫 = 手動建本 B → `copy_section` 克隆 → `search_pages` 找日期頁 → `get_page` 讀 → `update_page_content(replace)` 改」。Claude Desktop 對 `instructions` 的採用程度須**實測確認**。

**讀取工具須回傳下一步要用的 ID:** `get_page` / `get_page_images` / `get_page_files_info` 帶各內容物件的 `objectID`,否則模型選對 `delete_page_content` 也缺參數可帶(見 §5「刪內容物件」)。選對工具但缺參數 = 等於選錯。

**驗收(用真實 Claude Desktop,不用 sub-agent / API):** 建一組代表性使用者說法(例:「刪掉這個表格」「把這頁搬到另一節」「整理這節」「我現在在哪一頁」「把這張圖換掉」),在**接好 server 的真實 Claude Desktop**(production 消費端本人)上逐句輸入(用拋棄式測試筆記本,因 Desktop 會真的執行),用 §7 診斷日誌看它**選了哪個工具、帶什麼參數**;挑錯的回去改該條描述,每句重複幾次(有非決定性)迭代到穩,輸出「說法 → 工具」對照表。**不要**改用 CC 的 sub-agent 代測:那是 CC 的 harness/語境且 CC 知道設計意圖,會污染結果(測到「被暗示過的 Claude」而非「只看描述的天真 Claude」),且 agent 會動手/自我修正、蓋掉描述缺陷。

---

## 5. COM 技術要點(實作時遵循)

- 連線:`win32com.client.Dispatch("OneNote.Application")`(放在 `Win32ComBackend` 內,lazy import)。
- `GetHierarchy(startNodeId, scope, out xml)`:scope 用 HierarchyScope enum(hsNotebooks/hsSections/hsPages)控制深度;回傳 `one:` namespace 的 XML 樹,每個節點帶 `ID` 屬性。
- `GetPageContent(pageId, out xml, pageInfo)`:回傳頁面 XML。**預設不含二進位**(避免字串過長);圖片等二進位物件在 XML 裡只帶一個 OneNote callback ID,需再呼叫 `GetBinaryPageContent(pageId, callbackId)` 取實際 bytes。表格則直接是完整的 `one:Table / one:Row / one:Cell` 結構。
- 修改的標準流程:`GetPageContent` → 用 XML 工具改動目標節點 → `UpdatePageContent(xml, dateExpectedLastModified, schema, force)`。
  - `UpdatePageContent` 是**合併式**:只動你 XML 裡有指定且變更過的 page-level 物件,不會碰你沒指定的。
  - **並發保護**:帶上 `dateExpectedLastModified`(從讀取時的頁面取得);若頁面在這之間被改過,更新會失敗。**預設不要用 `force=true`**,以免蓋掉使用者正在編輯的內容;改為回報衝突讓上層決定。
- `FindPages`:文字搜尋,依賴 OneNote 自身索引。
- `DeleteHierarchy(objectId, ...)`:刪除節點。
- `OpenHierarchy(path, relativeToId, out objectId, CreateFileType)`:**建立/開啟層級節點**。本專案只用 `cftSection` 建節(name 以 `.one` 結尾、relativeTo 給父 notebook **或節群組**的 ID),及 `cftNone`(只開不建)。⚠️ **`cftNotebook`(建本)不使用**——COM 在本地建出的本不會綁雲端同步,實證不可行,故無 `create_notebook` 工具;**`cftFolder`(建節群組)亦不使用**——本期無建節群組工具,需要時於 OneNote 手動建。`UpdateHierarchy` 用來設 `pageLevel`(子頁,`create_page` 與複製都會用到,屬必用);改名(`rename_node`)、重排(`restructure_section`/`reorder_sections`)、搬移(`move_page`)亦全走它。
  - ⚠️ **為何不建本:** COM 在本地新建的筆記本不會落在會同步到 OneDrive 的位置(只會是本機本),這正是移除 `create_notebook` 的原因。**`create_section` 則安全**——在既有(已同步)notebook 內建節會自動沿用其同步。需要新增整本時,於 OneNote **手動**建好已同步的本,再用 `create_section` / `copy_section` 往裡面加。
- **層級重排紀律(`UpdateHierarchy`):** 順序由提交 XML 中**子元素的排列順序**決定,沒有位置索引屬性。`UpdateHierarchy` 對部分清單會「推斷」意圖——微軟文件明言:只提交部分子元素時,未提交者的落點**不可預期**。因此重排**必須整批提交該層級的完整子元素清單**(照目標順序、頁面各自帶 `pageLevel`),嚴禁只丟想動的那幾個;若該本含節群組,筆記本直屬子元素是 `one:Section` 與 `one:SectionGroup` 的**混合清單**,整批提交須兩種都含。實證範圍:節內頁面與本內節的重排有社群實例;**最上層筆記本清單的排序未驗證,不納入**;**跨節搬頁(`move_page`)的可靠性需 VM 實機驗證後才轉正**。結構性變更(重排/搬移/改名)的工具描述應註明「建議先克隆備份(copy_section),且 Claude 應先提案經使用者確認再批次套用」。**但「把整份清單交給 LLM」是人因懸崖**(實測:46 頁的 section,模型寧可去寫外部暫存檔輔助排序,撞到唯讀沙箱而中斷)——故移**單一**頁的常見需求由 `reposition_page` 承接:呼叫端只給 page_id + after_page_id,service 在 seam 內讀整層、用 `addnext` 原地搬一個元素、靠 node-ID 守恆保證「完整清單」紀律不破(與 `move_page` 同套路)。整批多頁重排或設多頁 `pageLevel` 才用 `restructure_section`。
- **目前檢視位置(`Windows` 介面):** `Application.Windows.CurrentWindow` 取作用中視窗,其 `CurrentPageId / CurrentSectionId / CurrentSectionGroupId / CurrentNotebookId` 即使用者目前停留的位置;`get_current_context` 以此實作,並用範圍化 `GetHierarchy` 把 ID 解析成名稱回報。限制:**無開啟視窗時取不到**(回報明確錯誤,勿猜);多視窗以作用中視窗為準;**頁內游標位置/選取文字無 API 可取**,粒度止於「頁」;工具描述應提醒 Claude 動作前先回報「你目前在 X 頁」供使用者確認(避免使用者已切頁的時間差)。
- 錯誤處理:對 `RPC_E_SERVERCALL_RETRYLATER`(0x8001010A,OneNote 忙碌/同步中)做重試 + 退避。
- OneNote XML schema:核心元素 `one:Page / one:Outline / one:OEChildren / one:OE / one:T`(文字)、`one:Table/Row/Cell`、`one:Image`。注意 `one:` namespace 前綴。

### 圖片/表格處理細節

- **讀圖片** → `get_page_images` 抽出 binary,以 **MCP image content (base64)** 回傳,讓 Claude 用視覺辨識;一併回報該圖在頁面中的相對位置(若可得)。
- **插入圖片** → `insert_svg_image` 接收 **SVG markup**(由模型直接產生的向量圖:路線圖/圖表/橫幅),伺服器端用 `resvg_py` 光柵化成 PNG,再組 `<one:Image>` 含 `<one:Data>` base64 → `UpdatePageContent`。**僅支援向量**:raster 插入(base64 點陣/照片)在 v1.0.9 移除——base64 經模型 token stream 太慢;內嵌 raster data: URI 會被拒絕,CJK 文字須用明確字族(微軟正黑體);照片請使用者手動插入。**點陣圖若已是本機檔(使用者指定的照片,或能寫檔/跑 code 的 client 在磁碟上產出的圖)改用 `insert_image_from_path`:server 端讀檔,bytes 不經模型;只存在 context、無檔落地者仍須手動。** `make_image` + `apply_page_edit` seam 不變(SVG 與本機檔點陣共用)。
- **讀表格** → 解析 `one:Table` 成結構化 rows(list of list of cell text),不要攤平成單一字串。
- **新增/改表格** → 組 `one:Table` XML(列、欄、儲存格),經 `UpdatePageContent` 寫入。
- **刪內容物件** → 依物件在頁面的層級分兩條路徑(VM 實證:`DeletePageContent` 接受頁層 `one:Outline`/`one:Image`/`one:InsertedFile`,但對**行內** OE 一律以 `0x8004200E` 拒絕):
  - **頁層物件(整個大綱/頁層圖片/頁層附件/嵌入物件)** → `delete_page_content(pageId, objectId)` 走 `DeletePageContent`。**不是 `DeleteHierarchy`**(那只刪頁/節/筆記本層級),**也不能靠 `UpdatePageContent` 省略物件來刪**(它合併式、不會移除未指定物件)。
  - **大綱內物件(表格/段落/行內圖片/行內附件)** → `delete_inline_content(pageId, objectId)` 走**編輯 seam**(`GetPageContent` → 從樹上移除該元素並修剪變空的 OE/OEChildren/Outline、表格 cell 補最小段落不留空殼 → `UpdatePageContent`);同大綱其他段落保留(「刪表格、留段落」)。傳表格自身 objectID 刪整表、傳段落/OE objectID 刪段落(行內圖/附件用其外層 OE 的 objectID,即 `get_page` 回報者)。頁層 objectID 傳此處會被擋並導向 `delete_page_content`。
  - `objectId` 由 `get_page` / `get_page_images` / `get_page_files_info` 取得(讀取工具須一併回傳各內容物件的 objectID);兩條路徑皆帶 `dateExpectedLastModified` 做並發保護,預設不 `force`。
- **手寫墨跡** → 視為已知限制:除非 OneNote 已做墨跡辨識存了文字,否則本期不支援;明確記在限制清單。

### 附件與嵌入物件(`one:InsertedFile`)

頁面上的「插入檔案」(附件圖示)與「嵌入物件」(如 Excel 試算表)在頁面 XML 中是 `one:InsertedFile` 元素,帶 `pathSource`(原始來源路徑)、`pathCache`(OneNote 本機快取檔路徑,由 OneNote 自行填寫與管理)、`preferredName`(UI 顯示名)、`objectID`;schema 上可帶 `Previews` 或 `Printout` 子元素(檔案列印/預覽影像)。本期範圍(刻意收斂):

- **二進位位置與圖片不同:** 附件內容是磁碟上 `pathCache` 指向的快取檔——讀內容**直接讀該檔**,**不是** `GetBinaryPageContent`(那是圖片的 callback 路徑)。快取檔可能不存在(尚未同步、快取被清),工具須優雅回報「快取不可用」,不得崩潰。
- **`get_page_files_info`(任何型別都支援):** 回傳每個附件/嵌入物件的 `preferredName`、副檔名/型別、大小(由快取檔取得;無快取則註明不可得)、`objectID`(供 `delete_page_content` 用)、**`kind`(呈現形態:附件圖示/嵌入預覽/檔案列印——依 `Previews`/`Printout` 子元素判別,確切判別法待 VM dump 實證)**。**不解析內容。**附件與嵌入物件(如 Excel 試算表)在 XML 同為 `one:InsertedFile` 家族,**統一由本工具涵蓋,不拆分**;以 `kind` 區分即可。
- **`get_page_files`(內容抽取,僅限三類):** 文字類(txt/csv/json/md/程式碼)解碼回傳;圖片附件以 MCP image content(base64)回傳供視覺辨識;PDF 在 server 端抽出文字(用純 Python 輕量庫,顧 PyInstaller 凍結相容)。**docx/xlsx/pptx 與其他型別一律不解析**,回中繼資料並明示不支援。大型附件設大小上限並截斷回報。
- **克隆保留(`copy_page`/`copy_section` 必達):** 沿用舊 `pathCache` 等於指向來源快取的死引用,不可照抄;克隆時把快取檔複製到暫存位置,改寫元素為 `pathSource` 指向副本、移除舊 `pathCache`,讓 OneNote 重匯並重建自己的快取。**重匯機制與嵌入物件(試算表)克隆後的行為待 VM 實機驗證。**來源快取不可用時,該附件無法保真複製,須在結果中明確回報(不可默默略過)。
- **嵌入式試算表等物件的確切 XML 形態**(`InsertedFile`+`Previews` 或其他)**待 VM dump 實證**;不論形態,原則一致:**info 可得、可刪、克隆保留、不抽內容、不寫**。
- **不支援寫入(無 `insert_file` 工具):** 插入附件技術上走 `pathSource` 重匯可行,但本期刻意不納入;明確記在限制清單。
- **檔案列印(`Printout`)呈現為頁面影像** → 與圖片同限制:燒進像素的內容(含日期)不可改。

### 格式保真(完整保真,必達)

需求定案為**完整保真**:讀得出格式、就地編輯不擾動既有格式、且能對新內容設定格式。格式資訊全在 GetPageContent 的 XML 裡,API 不丟;能否保真取決於以下實作規則:

- **讀取** — 解析 `<one:T>` CDATA 內的 `<span style>` 取得字元層級格式(粗/斜/底線/刪除線、`color`、highlight/螢光色、`font-family`、`font-size`),**並把 OE 的 `quickStyleIndex` 對照頁面 `QuickStyleDef` 解出段落層級的字型/字級/顏色**(別只看 inline span,否則樣式定義在 QuickStyleDef 的段落會漏)。回給 Claude 的表示法須**無損可重建**——用富文字/HTML 或「runs + 解析後 style」模型,**不可只回純文字**。
  - ⚠️ **螢光/底色同時帶兩個屬性:** OneNote 同時使用 `background:yellow`(標準 CSS)和 `mso-highlight:yellow`(Office 相容 hint)。parser 兩個都要讀;builder 新增螢光時兩個都要寫,否則某些 OneNote 版本顯示不一致。顏色值可以是任何 CSS 顏色(yellow、#FFFF00 等)。
- **修改/附加(最關鍵)** — 必須在 GetPageContent 拿到的**原始 XML 樹上做局部手術式修改**(就地改目標節點),把整棵改過的樹交回 `UpdatePageContent`。**嚴禁先轉成精簡/純文字模型再重建**:只要走過有損模型,沒被動到的段落格式就會掉。為此,編輯路徑的資料模型就是 XML 樹本身(lxml/ElementTree)+ 輔助函式,而非簡化 DTO。
- **新增** — 新文字的 `<one:T>` 自帶 `<span style>`,段落用 OE style / `quickStyleIndex`;工具輸入要能表達格式(粗體/顏色/字型/字級等,或接受可映射成 span 的富文字)。
- **表格** — 儲存格內也是 OE/T,套用同樣規則,讀寫都保留儲存格格式。

**誠實的邊界:** 即使我們不主動更動,OneNote 在自己重新算繪/存檔時偶爾會把 span 邊界正規化或重複 QuickStyleDef(已知行為)。所以保真保證落在**語意層級**(使用者看到的字型/顏色一致),不保證 XML 位元完全相同;測試據此驗「語意一致」而非「位元相同」。

### 複製/克隆(忠實克隆 + 複製後改寫)

`copy_*` 工具做的是**忠實克隆**,內部用「整頁原始 XML 克隆」:讀來源頁完整 XML(`GetPageContent`),**把圖片 binary 以 `GetBinaryPageContent` 取出 inline 進去**、**連同該頁的 `QuickStyleDef` 一起帶過去**(否則 `quickStyleIndex` 失準、格式跑掉)、**附件/嵌入物件以「複製 `pathCache` 快取檔 → 改用 `pathSource` 指向副本重匯」帶過去**(見 §5「附件與嵌入物件」)、**重設 object ID**、**依來源設定目標頁 `pageLevel`**(保留子頁階層),再寫到目標節。`copy_section` 則是先建好目標節、再逐頁套用 `copy_page`(保留 `pageLevel`);節是複製的最大單位(無整本克隆,見下「硬限制」)。

⚠️ **克隆走「接近原始的 XML 直通」(讀 → 重設 ID／inline 圖片／帶樣式 → 寫),嚴禁繞經 `get_page` 的結構化/富文字表示再重建。** `get_page` 那條路是給 Claude 閱讀與編輯用的,對整頁複製會有損(理由同 §5 格式保真的「嚴禁有損模型」)。換言之複製**不是** §4「多對一共用核心」那條手術式 merge 路徑,而是自己一條整頁 create 路徑。

**「複製後改寫」是處理日期調整(及其他內容轉換)的推薦流程,優於邊讀邊改:**

1. 先在 OneNote 桌面版**手動建好一個會同步的空筆記本 B**(COM 不能在本地建本;若 A 含節群組,B 的節群組結構也一併手動建好),再用 `copy_section` 把 A 的各節逐一忠實克隆進 B——這步純機械、不經 LLM 改字,起點是完美副本。
2. 在 B 上改日期:`search_pages`(範圍限 B)找出含日期的頁 → `get_page` 讀乾淨文字認出日期 → `update_page_content`(replace,手術式、保真)只改日期那幾段。

如此把「忠實複製」與「日期改寫」拆成兩個可各自驗證的步驟(先驗 B≡A,再驗日期),改寫走既有保真編輯路徑,只動日期文字、不擾動其他格式/表格/圖片。代價是比邊讀邊改多幾趟往返,但穩健度與可驗證性值得。

**此流程的硬限制:**

- **圖片內燒進像素的日期改不了**(那是圖不是文字);只有文字與表格儲存格裡的日期可改。
- **無整本克隆工具**(`create_notebook` / `copy_notebook` 已移除,因 COM 不能在本地建本)。要「克隆整本」須**先手動建好目標筆記本 B**(若來源含節群組,連節群組結構一併手動建),再對每個節用 `copy_section`;`copy_section` 不會自行重建節群組,故 B 的層級骨架是人工準備、內容(節內頁面/格式/表格/圖片)才靠克隆。
- 「所有日期一致位移」是 Claude 的編排工作;哪些日期該跟著移(行程日 vs. 某些固定日期)邊角可能需使用者界定規則。

---

## 6. 測試策略

- **Tier 1 — host(主要開發迴圈):** 對 `tests/fixtures/` 的真實 XML 做純函式單元測試(parser 與 builder 雙向),每次改動都跑。目標 XML 層高覆蓋。
- **Fixtures 自動更新:** `scripts/dump_fixtures.py`(在 VM 內以 COM dump 真實 `GetHierarchy` 與數頁 `GetPageContent`,含文字/圖片/表格);§2.4 的 Tier 2 迴圈每跑一次會順手重抓並回送 host,所以樣本持續貼近真實,不必手動維護。
- **Tier 2 — VM(整合測試,檢查點):** 標 `@pytest.mark.windows` 的測試經 §2.4 的 `remote_test` 迴圈,對真實 OneNote 跑 create → get → update(含改表格)→ delete round-trip,並驗證 `dateExpectedLastModified` 並發保護。host 上自動 skip。
- **格式保真回歸測試(Tier 2 必含):** 讀入含混合格式的頁 → 只改其中一段 → 寫回 → 重讀,斷言**其他段落的字型/字級/顏色未變(語意一致)**,確保編輯不擾動既有格式。

---

## 7. 安全與穩健性

- 完全不需 Azure/Graph/token;不對外連網。
- 寫入預設保守:不 `force`、不刪未確認的節點;破壞性操作(delete、覆蓋式 update)在工具描述中明確標示。
- 一律範圍化存取,避免全域掃描。

### 診斷日誌(Debug Log,預設關閉)

正式運作時**預設不記詳細日誌**(最多記錯誤/警告),避免日誌膨脹、效能負擔,以及把筆記本內容(可能含客戶/行程資料)寫進檔案。需要除錯時才開啟,開關走**環境變數**(採此方案而非 MCP 工具:診斷開關必須帶外,獨立於被診斷的工具/連線管線;且不污染 §4 的工具命名空間):

- **開關:** 環境變數 `ONENOTE_MCP_LOG_LEVEL`(預設 `ERROR` = 詳細日誌關;設 `DEBUG` 開啟)。server 由 client 以子行程啟動、啟動時讀取環境,故**改值後需重啟 Claude Desktop 才生效**(已接受此代價,換取單純)。設定方式:在該 server 於 Claude Desktop 設定檔的 `env` 區塊加入此變數。`--configure` **預設不寫入**此變數(維持預設關),由使用者需要時自行加上。
- **輸出去向:** 輪替檔案(預設 `%LOCALAPPDATA%\OneNoteMCP\logs\`,可由 `ONENOTE_MCP_LOG_FILE` 覆寫)與/或 stderr;**絕不寫 stdout**(見 §8,否則汙染 JSON-RPC)。Claude Desktop 會把 server 的 stderr 收進它自己的 MCP log(`mcp-server-<name>.log`),可作交叉參考。
- **開啟時記什麼:** 每次工具呼叫一筆——時間、工具名、輸入參數(大型/二進位欄位如 base64 影像、inline 圖片資料一律**截斷**至上限,避免日誌爆量)、結果(成功/錯誤 + COM 錯誤碼如 `RPC_E_SERVERCALL_RETRYLATER`)。
- **穩健性:** 日誌本身絕不可拋例外或中斷 server;寫檔失敗應靜默降級(例如退回 stderr),不得影響工具執行。
- **用途:** (a) 正式環境故障診斷;(b) 開發期驗證「某句使用者說法 → Claude 選了哪個工具、帶什麼參數」是否如預期(工具描述/選擇調校的訊號來源)。

---

## 8. 交付與執行

**正式部署(每位員工的工作 Windows PC,一鍵 installer):**

目標是同事像裝一般 Windows App,跑一個 `OneNoteMCP-Setup.exe` 就裝好——免手動裝 Python、免手動編輯設定檔。兩層打包,**皆在 VM 內建置(PyInstaller / Inno Setup 不能從 Linux 跨平台編譯)**:

1. **凍結 server 成獨立執行檔(PyInstaller)** — 內含 Python runtime 與 pywin32,目標機不必另裝 Python。凍結相容性限制:必須用晚繫結 `Dispatch("OneNote.Application")`(勿用 `gencache.EnsureDispatch`,凍結後抓不到產生的快取模組);build 成 **console** 程式(它是 stdio server),且**除 MCP 協定外不得有任何輸出寫到 stdout**(log 一律走 stderr 或檔案,否則汙染 JSON-RPC;日誌開關與規格見 §7「診斷日誌」)。
2. **包成 installer(Inno Setup 或 NSIS)** — 產出 `OneNoteMCP-Setup.exe`,裝進 `%LOCALAPPDATA%`、註冊解除安裝(出現在「新增/移除程式」)。安裝後自動完成設定:
   - **把 `onenote` 項目合併進該使用者「所有偵測到的 MCP client」設定檔(Claude Desktop 與 Antigravity)**——只加自己的 key、保留其他既有 connector,**不可覆蓋整個檔**。由凍結後的 exe 提供 `--configure` 子命令處理偵測與 JSON 合併,installer 只負責呼叫它。**一支 exe、一次安裝即對兩種 client 都設定,不必裝兩次。**
   - ⚠️ **Claude Desktop 兩種變體都要支援,無法假設同事裝哪一種:**
     - **Store(MSIX)版** — 設定檔在 `%LOCALAPPDATA%\Packages\<PackageFamilyName>\LocalCache\Roaming\Claude\claude_desktop_config.json`(目前 PFN 是 `Claude_pzs8sxrjxfjjc`,但**用 `Get-AppxPackage *Claude*` 動態取得 PFN 再組路徑,別硬編**)。偵測安裝:該 appx 套件是否存在。
     - **一般下載版** — 設定檔在 `%APPDATA%\Claude\claude_desktop_config.json`;安裝目錄通常在 `%LOCALAPPDATA%\Programs\Claude\`。偵測安裝:該程式目錄或解除安裝登錄機碼是否存在。
   - ⚠️ **Antigravity(CLI 與 IDE 共用同一設定):** 設定檔在 `%USERPROFILE%\.gemini\config\mcp_config.json`(沿用 `.gemini` 目錄名;CLI 與 IDE 共用,**寫這一個檔即同時涵蓋兩者**)。schema 與 Claude Desktop 形狀一致:單一 `mcpServers` 物件,每個 server 用 `command`/`args`/`env`/`cwd`(stdio)。偵測安裝:Antigravity 本體 / `agy` CLI 是否存在(**確切偵測點待 VM 實證**;Antigravity 很新且頻繁更新,路徑/schema 以實機為準)。⚠️ **已知 day-one bug:全域 MCP 設定的環境變數傳遞不穩**,故 §7 的 `ONENOTE_MCP_LOG_LEVEL` 走 `env` 可能不生效——需實測,必要時改用 `ONENOTE_MCP_LOG_FILE` 預設路徑。另:CLI(TUI)無法以剪貼簿貼圖,但 `get_page_images` 走 server 端回傳不受此限(圖片是否確實餵入模型仍須真實 client 實測,屬 §4 驗收範疇)。
   - `--configure` 邏輯**不是「二選一 + fallback」**,而是:獨立偵測**每一種**支援的 client 變體(Claude Desktop Store 版、Claude Desktop 一般版、Antigravity)各自是否安裝,**對每一個有裝的都寫入**(裝幾個就寫幾個,互不干擾);目標資料夾不存在就建立(同事可能還沒首次開過該 client,設定檔尚未產生);已存在就**合併**自己的 key、保留其他 connector。
   - 偵測「是否安裝」要看**安裝本身**(appx 套件 / 程式目錄 / 登錄),別只看設定檔在不在(首次開啟前不會有)。
   - **完全沒偵測到任何支援的 client** → 不要默默寫一個沒用的檔;明確提示「未偵測到 Claude Desktop 或 Antigravity,請先安裝後重跑設定」。**安裝本身仍正常完成**(server exe 照裝);只是這次沒有可寫入的對象。
   - **client 不需在安裝前就裝好;日後新增 client 用「手動重跑」補設定。** `--configure` 只設定「執行當下偵測到」的 client,所以同事日後才裝 Claude 或 Antigravity 也沒關係。為此 installer 須放一個**好按的開始功能表捷徑**(例:「OneNoteMCP — 重新偵測並設定」),點下去就是跑 `<exe> --configure`——裝完新 client 點一次即補上。**刻意採此手動模式**:不盲寫未安裝 client 的設定檔(會落在猜測/未來可能變動的路徑、靜默失效,尤其 Store 版 PFN 與 Antigravity 新路徑),也不註冊登入時自動重跑的背景工作(避免常駐footprint);以「一鍵重跑」換取零猜測、零常駐。
   - 這段必須以**安裝使用者本人(per-user)**身分執行,才會落在正確的使用者 profile。
   - 提示使用者**重啟對應的 client(Claude Desktop / Antigravity)** 以載入設定。
- 前置檢查:OneNote 桌面版,以及至少一個支援的 client(Claude Desktop 或 Antigravity)已安裝(缺則提示)。員工本人已登入互動桌面,所以正式環境**不需要** autologon / 互動 session 執行器(那些只在測試 VM 用)。
- 注意:未簽章的 installer 會觸發 SmartScreen「不明發行者」警告,公司 AV/政策也可能擋;內部散布前評估是否需程式碼簽章或請 IT 加白名單。
- 「不做自動部署」= 不寫 MDM/GPO/SCCM 集中推送;同事仍是手動執行這個 installer。

**開發測試環境(VM)交付物:** `remote_test.sh`、互動 session 執行器、`dump_fixtures.py`——細節見 §2.4。

**打包交付物(在 VM 內建置):** PyInstaller spec、Inno Setup/NSIS 腳本、exe 的 `--configure` 子命令(JSON 合併 + 多 client 偵測:Claude Desktop Store/一般版 + Antigravity)、開始功能表「重新偵測並設定」捷徑(呼叫 `--configure`,供日後新增 client 時一鍵補設定)。

---

## 9. 給 Claude Code 的任務

> ⛔ **反模式護欄(最高優先,凌駕任何直覺):本專案 COM-only。** 嚴禁 Microsoft Graph、`msal`、`requests`/任何對 `graph.microsoft.com` 或其他微軟雲端端點的 HTTP 呼叫、以及 Azure app/token 那一套。所有 OneNote 存取一律經 §3 的 `OneNoteBackend` → pywin32 COM。**若你發現自己正要 `import msal`、發 HTTP 請求、或寫 OAuth/token 邏輯,立刻停下重想——那代表走錯路了。** 現成社群版幾乎都是 Graph(理由與差異見 §1.2),不要把它們的做法照抄進來。

1. 讀完本規格,列出你需要釐清的問題(若有)。
2. 產出一份**分階段開發計畫**,建議的分期方向(可調整):
   - **Phase 0** — 專案骨架、`OneNoteBackend` 介面、pywin32 守護匯入、host 端 CI(lint + 單元測試)跑得起來;**並行**把 §2 的 Windows VM、autologon 互動 session 執行器、`remote_test` 迴圈建起來(這是後續所有整合測試的前提)。**並在 VM 實機先確認 COM 可喚得起來**(`Dispatch("OneNote.Application")` + 一次 `GetHierarchy` 成功)再往下做——確保裝的是有 COM 的桌面版 OneNote、互動 session 正常,環境問題在此一次清掉。
   - **Phase 1** — OneNote XML 層(parse + build)純函式 + 對 fixtures 的單元測試;先做讀取相關(hierarchy、page text、table、image callback id)。**保真要求:解析 inline span 樣式 + 解析 `quickStyleIndex`→`QuickStyleDef`;build 端能帶樣式(見 §5 格式保真)。**
   - **Phase 2** — 讀取類 MCP 工具(list_*、get_page、get_page_images、get_page_files_info、get_page_files、search_pages、get_current_context)接到 `FixtureBackend`,Linux 全綠(附件的型別感知抽取——文字解碼/PDF 抽文字——是純函式,host 可測;讀 `pathCache` 檔案收在 backend 後面)。
   - **Phase 3** — `Win32ComBackend` 實作 + `dump_fixtures.py`;經 `remote_test` 迴圈在 VM 跑讀取整合測試(不再手動搬機器)。
   - **Phase 4** — 寫入類(create_section、create_page〔可設 `pageLevel`〕、update_page_content〔含 append/insert/replace〕、create_table、insert_image)+ **層級結構類(restructure_section、reorder_sections、rename_node;`move_page` 跨節搬移先在 VM 實證可靠才轉正)** + 並發保護 + **手術式就地修改(保真)** + Windows round-trip 整合測試 + **格式保真回歸測試**。
   - **Phase 5** — 複製/克隆:raw-XML 克隆核心(inline 圖片 binary、帶 `QuickStyleDef`、**附件 `pathCache`→`pathSource` 重匯**、重設 object ID、設 `pageLevel`)+ `copy_page`/`copy_section`;整合測試驗節層級「B≡A」忠實度(**含附件/嵌入物件保留**;嵌入試算表克隆行為在此實證)。「複製後改寫」沿用既有 `get_page` + `update_page_content`,不另開工具。
   - **Phase 6** — 刪除(`delete_node` 層級 + `delete_page_content` 內容物件:圖片/表格/大綱)、錯誤/重試強化、**診斷日誌(§7,環境變數 `ONENOTE_MCP_LOG_LEVEL` 開關、參數截斷、不碰 stdout)**、**工具描述強化(§4,跨全套對比式描述 + 行為契約 + server `instructions` + 真實 Desktop 選擇驗收)**、煙霧測試;**打包:PyInstaller 凍結 → Inno Setup/NSIS installer(含自動寫入設定的 `--configure`,獨立偵測並處理 Claude Desktop〔Store + 一般版〕與 Antigravity〔CLI/IDE 共用〕共三種設定檔路徑、合併不覆蓋),在 VM 內建置產出 `OneNoteMCP-Setup.exe`**。
3. 每個 Phase 標明:在 Linux 可完成/驗證的部分 vs. 必須在 Windows 驗證的部分。
