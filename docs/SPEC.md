# OneNote MCP Server — 開發規格 (給 Claude Code)

> 這份文件是**規格**,不是完成的計畫。請先讀完、提出釐清問題,然後**產出一份分階段的開發計畫**(見最後一節)。

---

## 1. 目標

建立一個 **COM-only** 的 OneNote MCP server,讓 Claude(在 Windows 上的 Claude Desktop)能對 OneNote 桌面版做完整 CRUD:

- 讀取:筆記本/節/頁清單、頁面文字、**表格(保留結構)**、**圖片(二進位)**、文字搜尋。
- 寫入:**新增筆記本/節/頁**、**附加或就地修改內容(含表格)**、**新增表格**、**插入圖片**。
- 刪除:刪頁/節/筆記本(層級節點);刪頁面內容物件——圖片、表格、大綱。
- 複製:忠實克隆筆記本/節/頁(保留格式、表格、圖片、子頁階層);搭配既有編輯工具做「複製後改寫」(見 §4 複製工具、§5)。
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

執行/消費端:MCP server 必須跑在有 OneNote 的 Windows 上,且通常由 Claude client 在本機以子行程啟動,所以 client + server + OneNote 同機。

- **開發測試期** — 跑在 §2 的 Windows VM(需要 autologon + 互動 session 執行器,那是測試環境的需求)。
- **正式部署** — 跑在**每位員工自己的工作 Windows PC**。該 PC 本就有員工登入的互動桌面,所以正式環境**不需要** autologon / 互動 session 執行器那套;每台只看得到該員工自己帳號、自己同步的 OneNote(天然分用戶,無共用憑證)。由各員工執行一個 installer EXE 完成安裝(見 §8),像裝一般 Windows App;不做集中式自動推送(MDM/GPO)。

若 client 用 Claude Desktop(Microsoft Store 版),設定檔在 `...\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json`,以 Developer → Edit Config 確認。Claude Code(Linux host)只是開發工具,不是執行環境。

---

## 4. 要實作的 MCP 工具

| 工具 | 用途 | 對應 COM 方法 |
|---|---|---|
| `list_notebooks` | 列所有筆記本 | `GetHierarchy(scope=hsNotebooks)` |
| `list_sections` | 列某本的節 | `GetHierarchy(notebookId, hsSections)` |
| `list_pages` | 列某節的頁(**範圍化,含 `pageLevel` 子頁階層**) | `GetHierarchy(sectionId, hsPages)` |
| `search_pages` | 跨/範圍文字搜尋 | `FindPages` |
| `get_page` | 取單頁內容(**保真富文字** + 結構化表格) | `GetPageContent(pageId)` |
| `get_page_images` | 取頁面圖片(二進位) | `GetPageContent` → `GetBinaryPageContent(callbackId)` |
| `create_notebook` | 建立筆記本 | `OpenHierarchy(path, "", out id, cftNotebook)` |
| `create_section` | 在指定本建立節 | `OpenHierarchy(name+".one", notebookId, out id, cftSection)` |
| `create_page` | 在指定節建新頁(可設 `pageLevel` 子頁) | `CreateNewPage` (+ `UpdateHierarchy` 設 pageLevel) + `UpdatePageContent` |
| `update_page_content` | 改頁面內容:append / insert / replace(含改表格儲存格/加列) | `GetPageContent` → 改 XML → `UpdatePageContent`(純 append 可免讀全頁) |
| `create_table` | 新增表格 | 組 `one:Table` XML → `UpdatePageContent` |
| `insert_image` | 插入圖片到頁面 | 組 `one:Image` + base64 `one:Data` → `UpdatePageContent` |
| `delete_node` | 刪頁/節/筆記本(層級) | `DeleteHierarchy(objectId)` |
| `delete_page_content` | 刪頁面內容物件(圖片/表格/大綱) | `DeletePageContent(pageId, objectId)` |
| `copy_page` | 忠實克隆單頁到目標節 | 內部 raw-XML 克隆(見 §5「複製/克隆」) |
| `copy_section` | 忠實克隆整節 | 建節 + 逐頁 `copy_page`(保留 pageLevel) |
| `copy_notebook` | 忠實克隆整本 | 建本 + 逐節 `copy_section`(受 `create_notebook` 雲端限制) |
| `restructure_section` | 同節內**整批**重排頁面順序 + 調整 `pageLevel` 階層 | `GetHierarchy` → 重排完整頁清單(每頁帶 `pageLevel`)→ `UpdateHierarchy`(紀律見 §5「層級重排」) |
| `reorder_sections` | 重排某本內節的順序 | 同上,對 `one:Section` 元素整批重排(僅節/頁有實證;**筆記本層級排序未驗證、不納入**) |
| `rename_node` | 重新命名頁/節 | `UpdateHierarchy`(改 name 屬性) |
| `move_page` | 把頁搬到別的節 | `UpdateHierarchy`(把頁元素掛到目標節下;**跨節搬移可靠性待 VM 實機驗證**) |

**範圍化原則:** 永遠用 notebook/section 範圍的 `GetHierarchy` 與 `FindPages`,**不要**對全部頁面逐頁掃描——這是先前 Graph 全域搜尋撞牆的同一個坑,逐頁打 COM 又慢又容易踩 RPC 忙碌錯誤。

**工具與 COM 方法是多對一(勿重寫):** `update_page_content`(含 append/insert/replace)、`create_table`、`insert_image` 最後全都走同一個 `UpdatePageContent`——它是 OneNote 唯一的「寫內容進頁面」通用原語。實作上必須**共用單一「套用頁面內容變更」核心**(集中處理 `GetPageContent → 改 XML → UpdatePageContent` 的並發保護 `dateExpectedLastModified` 與手術式格式保真),這些 MCP 工具只是它的薄 facade(各自負責把自己的輸入轉成內容 XML)。**不要各自重寫一套 round-trip 邏輯**,否則程式會分歧、難維護。分開命名是為了 LLM 好用,不是要分開實作。

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
- `OpenHierarchy(path, relativeToId, out objectId, CreateFileType)`:**建立/開啟層級節點**。建節用 `cftSection`(name 以 `.one` 結尾、relativeTo 給父 `notebookId`);建本用 `cftNotebook`(給路徑/位置);`cftFolder`=節群組、`cftNone`=只開不建。`UpdateHierarchy` 用來設 `pageLevel`(子頁,`create_page` 與複製都會用到,屬必用);改名(`rename_node`)、重排(`restructure_section`/`reorder_sections`)、搬移(`move_page`)亦全走它。
  - ⚠️ **建 notebook 要指定位置(路徑):** 你們是雲端同步筆記本,新建的本要落在會同步到 OneDrive 的位置,否則只是本機筆記本。**建 section 在既有(已同步)notebook 內則自動沿用其同步,較單純**——多數情況建議建節而非建本。
- **層級重排紀律(`UpdateHierarchy`):** 順序由提交 XML 中**子元素的排列順序**決定,沒有位置索引屬性。`UpdateHierarchy` 對部分清單會「推斷」意圖——微軟文件明言:只提交部分子元素時,未提交者的落點**不可預期**。因此重排**必須整批提交該層級的完整子元素清單**(照目標順序、頁面各自帶 `pageLevel`),嚴禁只丟想動的那幾個。實證範圍:節內頁面與本內節的重排有社群實例;**最上層筆記本清單的排序未驗證,不納入**;**跨節搬頁(`move_page`)的可靠性需 VM 實機驗證後才轉正**。結構性變更(重排/搬移/改名)的工具描述應註明「建議先克隆備份(copy_section/copy_notebook),且 Claude 應先提案經使用者確認再批次套用」。
- 錯誤處理:對 `RPC_E_SERVERCALL_RETRYLATER`(0x8001010A,OneNote 忙碌/同步中)做重試 + 退避。
- OneNote XML schema:核心元素 `one:Page / one:Outline / one:OEChildren / one:OE / one:T`(文字)、`one:Table/Row/Cell`、`one:Image`。注意 `one:` namespace 前綴。

### 圖片/表格處理細節

- **讀圖片** → `get_page_images` 抽出 binary,以 **MCP image content (base64)** 回傳,讓 Claude 用視覺辨識;一併回報該圖在頁面中的相對位置(若可得)。
- **插入圖片** → `insert_image` 接收 base64 影像 + media type(可選位置 x/y 與大小),組 `<one:Image>` 含 `<one:Data>` base64 → `UpdatePageContent` 寫入(COM 文件明載 `UpdatePageContent` 可加入 images)。
- **讀表格** → 解析 `one:Table` 成結構化 rows(list of list of cell text),不要攤平成單一字串。
- **新增/改表格** → 組 `one:Table` XML(列、欄、儲存格),經 `UpdatePageContent` 寫入。
- **刪內容物件(圖片/表格/大綱)** → `delete_page_content(pageId, objectId)` 走 `DeletePageContent`。**不是 `DeleteHierarchy`**(那只刪頁/節/筆記本層級),**也不能靠 `UpdatePageContent` 省略物件來刪**(它合併式、不會移除未指定物件)。`objectId` 由 `get_page` / `get_page_images` 取得(讀取工具須一併回傳各內容物件的 objectID);同帶 `dateExpectedLastModified` 做並發保護,預設不 `force`。
- **手寫墨跡** → 視為已知限制:除非 OneNote 已做墨跡辨識存了文字,否則本期不支援;明確記在限制清單。

### 格式保真(完整保真,必達)

需求定案為**完整保真**:讀得出格式、就地編輯不擾動既有格式、且能對新內容設定格式。格式資訊全在 GetPageContent 的 XML 裡,API 不丟;能否保真取決於以下實作規則:

- **讀取** — 解析 `<one:T>` CDATA 內的 `<span style>` 取得字元層級格式(粗/斜/底線/刪除線、`color`、highlight/螢光色、`font-family`、`font-size`),**並把 OE 的 `quickStyleIndex` 對照頁面 `QuickStyleDef` 解出段落層級的字型/字級/顏色**(別只看 inline span,否則樣式定義在 QuickStyleDef 的段落會漏)。回給 Claude 的表示法須**無損可重建**——用富文字/HTML 或「runs + 解析後 style」模型,**不可只回純文字**。
  - ⚠️ **螢光/底色同時帶兩個屬性:** OneNote 同時使用 `background:yellow`(標準 CSS)和 `mso-highlight:yellow`(Office 相容 hint)。parser 兩個都要讀;builder 新增螢光時兩個都要寫,否則某些 OneNote 版本顯示不一致。顏色值可以是任何 CSS 顏色(yellow、#FFFF00 等)。
- **修改/附加(最關鍵)** — 必須在 GetPageContent 拿到的**原始 XML 樹上做局部手術式修改**(就地改目標節點),把整棵改過的樹交回 `UpdatePageContent`。**嚴禁先轉成精簡/純文字模型再重建**:只要走過有損模型,沒被動到的段落格式就會掉。為此,編輯路徑的資料模型就是 XML 樹本身(lxml/ElementTree)+ 輔助函式,而非簡化 DTO。
- **新增** — 新文字的 `<one:T>` 自帶 `<span style>`,段落用 OE style / `quickStyleIndex`;工具輸入要能表達格式(粗體/顏色/字型/字級等,或接受可映射成 span 的富文字)。
- **表格** — 儲存格內也是 OE/T,套用同樣規則,讀寫都保留儲存格格式。

**誠實的邊界:** 即使我們不主動更動,OneNote 在自己重新算繪/存檔時偶爾會把 span 邊界正規化或重複 QuickStyleDef(已知行為)。所以保真保證落在**語意層級**(使用者看到的字型/顏色一致),不保證 XML 位元完全相同;測試據此驗「語意一致」而非「位元相同」。

### 複製/克隆(忠實克隆 + 複製後改寫)

`copy_*` 工具做的是**忠實克隆**,內部用「整頁原始 XML 克隆」:讀來源頁完整 XML(`GetPageContent`),**把圖片 binary 以 `GetBinaryPageContent` 取出 inline 進去**、**連同該頁的 `QuickStyleDef` 一起帶過去**(否則 `quickStyleIndex` 失準、格式跑掉)、**重設 object ID**、**依來源設定目標頁 `pageLevel`**(保留子頁階層),再寫到目標節。`copy_section`/`copy_notebook` 是建好對應層級後逐頁套用。

⚠️ **克隆走「接近原始的 XML 直通」(讀 → 重設 ID／inline 圖片／帶樣式 → 寫),嚴禁繞經 `get_page` 的結構化/富文字表示再重建。** `get_page` 那條路是給 Claude 閱讀與編輯用的,對整頁複製會有損(理由同 §5 格式保真的「嚴禁有損模型」)。換言之複製**不是** §4「多對一共用核心」那條手術式 merge 路徑,而是自己一條整頁 create 路徑。

**「複製後改寫」是處理日期調整(及其他內容轉換)的推薦流程,優於邊讀邊改:**

1. 用 `copy_notebook`(或 copy_section/page)把 A 忠實克隆成 B——這步純機械、不經 LLM 改字,起點是完美副本。
2. 在 B 上改日期:`search_pages`(範圍限 B)找出含日期的頁 → `get_page` 讀乾淨文字認出日期 → `update_page_content`(replace,手術式、保真)只改日期那幾段。

如此把「忠實複製」與「日期改寫」拆成兩個可各自驗證的步驟(先驗 B≡A,再驗日期),改寫走既有保真編輯路徑,只動日期文字、不擾動其他格式/表格/圖片。代價是比邊讀邊改多幾趟往返,但穩健度與可驗證性值得。

**此流程的硬限制:**

- **圖片內燒進像素的日期改不了**(那是圖不是文字);只有文字與表格儲存格裡的日期可改。
- `copy_notebook` 受 `create_notebook` 的雲端落點限制(B 要落在會同步的位置);多數情況可改用「在既有筆記本內 `copy_section`」避開。
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

---

## 8. 交付與執行

**正式部署(每位員工的工作 Windows PC,一鍵 installer):**

目標是同事像裝一般 Windows App,跑一個 `OneNoteMCP-Setup.exe` 就裝好——免手動裝 Python、免手動編輯設定檔。兩層打包,**皆在 VM 內建置(PyInstaller / Inno Setup 不能從 Linux 跨平台編譯)**:

1. **凍結 server 成獨立執行檔(PyInstaller)** — 內含 Python runtime 與 pywin32,目標機不必另裝 Python。凍結相容性限制:必須用晚繫結 `Dispatch("OneNote.Application")`(勿用 `gencache.EnsureDispatch`,凍結後抓不到產生的快取模組);build 成 **console** 程式(它是 stdio server),且**除 MCP 協定外不得有任何輸出寫到 stdout**(log 一律走 stderr 或檔案,否則汙染 JSON-RPC)。
2. **包成 installer(Inno Setup 或 NSIS)** — 產出 `OneNoteMCP-Setup.exe`,裝進 `%LOCALAPPDATA%`、註冊解除安裝(出現在「新增/移除程式」)。安裝後自動完成設定:
   - **把 `onenote` 項目合併進該使用者的 Claude Desktop 設定檔**——只加自己的 key、保留其他既有 connector,**不可覆蓋整個檔**。建議由凍結後的 exe 提供 `--configure` 子命令處理 JSON 合併,installer 只負責呼叫它。
   - ⚠️ **兩種 Claude Desktop 都要支援,無法假設同事裝哪一種:**
     - **Store(MSIX)版** — 設定檔在 `%LOCALAPPDATA%\Packages\<PackageFamilyName>\LocalCache\Roaming\Claude\claude_desktop_config.json`(目前 PFN 是 `Claude_pzs8sxrjxfjjc`,但**用 `Get-AppxPackage *Claude*` 動態取得 PFN 再組路徑,別硬編**)。偵測安裝:該 appx 套件是否存在。
     - **一般下載版** — 設定檔在 `%APPDATA%\Claude\claude_desktop_config.json`;安裝目錄通常在 `%LOCALAPPDATA%\Programs\Claude\`。偵測安裝:該程式目錄或解除安裝登錄機碼是否存在。
   - `--configure` 邏輯**不是「二選一 + fallback」**,而是:獨立偵測兩種變體各自是否安裝,**對每一個有裝的都寫入**(兩個都裝就兩個都寫,互不干擾);目標資料夾不存在就建立(同事可能還沒首次開過 Claude,設定檔尚未產生);已存在就**合併**自己的 key、保留其他 connector。
   - 偵測「是否安裝」要看**安裝本身**(appx 套件 / 程式目錄 / 登錄),別只看設定檔在不在(首次開啟前不會有)。
   - **兩種都沒偵測到** → 不要默默寫一個沒用的檔;明確提示「未偵測到 Claude Desktop,請先安裝再重跑」,並保留可單獨重跑 `--configure` 的入口(裝完 Claude 後能再執行一次補設定)。寫到沒裝的路徑 = 裝了但 Claude 看不到工具。
   - 這段必須以**安裝使用者本人(per-user)**身分執行,才會落在正確的使用者 profile。
   - 提示使用者**重啟 Claude Desktop** 以載入設定。
- 前置檢查:OneNote 桌面版與 Claude Desktop 已安裝(缺則提示)。員工本人已登入互動桌面,所以正式環境**不需要** autologon / 互動 session 執行器(那些只在測試 VM 用)。
- 注意:未簽章的 installer 會觸發 SmartScreen「不明發行者」警告,公司 AV/政策也可能擋;內部散布前評估是否需程式碼簽章或請 IT 加白名單。
- 「不做自動部署」= 不寫 MDM/GPO/SCCM 集中推送;同事仍是手動執行這個 installer。

**開發測試環境(VM)交付物:** `remote_test.sh`、互動 session 執行器、`dump_fixtures.py`——細節見 §2.4。

**打包交付物(在 VM 內建置):** PyInstaller spec、Inno Setup/NSIS 腳本、exe 的 `--configure` 子命令(JSON 合併 + Store/一般版路徑偵測)。

---

## 9. 給 Claude Code 的任務

> ⛔ **反模式護欄(最高優先,凌駕任何直覺):本專案 COM-only。** 嚴禁 Microsoft Graph、`msal`、`requests`/任何對 `graph.microsoft.com` 或其他微軟雲端端點的 HTTP 呼叫、以及 Azure app/token 那一套。所有 OneNote 存取一律經 §3 的 `OneNoteBackend` → pywin32 COM。**若你發現自己正要 `import msal`、發 HTTP 請求、或寫 OAuth/token 邏輯,立刻停下重想——那代表走錯路了。** 現成社群版幾乎都是 Graph(理由與差異見 §1.2),不要把它們的做法照抄進來。

1. 讀完本規格,列出你需要釐清的問題(若有)。
2. 產出一份**分階段開發計畫**,建議的分期方向(可調整):
   - **Phase 0** — 專案骨架、`OneNoteBackend` 介面、pywin32 守護匯入、host 端 CI(lint + 單元測試)跑得起來;**並行**把 §2 的 Windows VM、autologon 互動 session 執行器、`remote_test` 迴圈建起來(這是後續所有整合測試的前提)。**並在 VM 實機先確認 COM 可喚得起來**(`Dispatch("OneNote.Application")` + 一次 `GetHierarchy` 成功)再往下做——確保裝的是有 COM 的桌面版 OneNote、互動 session 正常,環境問題在此一次清掉。
   - **Phase 1** — OneNote XML 層(parse + build)純函式 + 對 fixtures 的單元測試;先做讀取相關(hierarchy、page text、table、image callback id)。**保真要求:解析 inline span 樣式 + 解析 `quickStyleIndex`→`QuickStyleDef`;build 端能帶樣式(見 §5 格式保真)。**
   - **Phase 2** — 讀取類 MCP 工具(list_*、get_page、get_page_images、search_pages)接到 `FixtureBackend`,Linux 全綠。
   - **Phase 3** — `Win32ComBackend` 實作 + `dump_fixtures.py`;經 `remote_test` 迴圈在 VM 跑讀取整合測試(不再手動搬機器)。
   - **Phase 4** — 寫入類(create_notebook、create_section、create_page〔可設 `pageLevel`〕、update_page_content〔含 append/insert/replace〕、create_table、insert_image)+ **層級結構類(restructure_section、reorder_sections、rename_node;`move_page` 跨節搬移先在 VM 實證可靠才轉正)** + 並發保護 + **手術式就地修改(保真)** + Windows round-trip 整合測試 + **格式保真回歸測試**。
   - **Phase 5** — 複製/克隆:raw-XML 克隆核心(inline 圖片 binary、帶 `QuickStyleDef`、重設 object ID、設 `pageLevel`)+ `copy_page`/`copy_section`/`copy_notebook`;整合測試驗「B≡A」忠實度。「複製後改寫」沿用既有 `get_page` + `update_page_content`,不另開工具。
   - **Phase 6** — 刪除(`delete_node` 層級 + `delete_page_content` 內容物件:圖片/表格/大綱)、錯誤/重試強化、煙霧測試;**打包:PyInstaller 凍結 → Inno Setup/NSIS installer(含自動寫入 Claude Desktop 設定的 `--configure`,獨立偵測並處理 Store + 一般版兩種設定檔路徑),在 VM 內建置產出 `OneNoteMCP-Setup.exe`**。
3. 每個 Phase 標明:在 Linux 可完成/驗證的部分 vs. 必須在 Windows 驗證的部分。
