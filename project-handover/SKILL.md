---
name: project-handover
description: 指定一個或多個 repo（本機路徑、本機專案名、或 git URL），深入分析程式碼與 git 歷史，整理成交接／教學文件 handover.md、技術目錄 tech-catalog.md（專案用到的每項技術與要點），並產出 16:9 的 .pptx 簡報（含架構圖、流程圖、講者備註）。凡是使用者提到專案交接、交接簡報、教學簡報、onboarding、新人導讀、「幫我介紹／整理這個 repo」、「把手上專案做成 ppt」、「這專案用了哪些技術」、技術盤點、knowledge transfer、離職或轉調前整理專案，即使沒說「skill」或「ppt」也要使用。Use for any request to explain, document, or hand over a codebase as slides or a handover doc.
---

# Project Handover：repo → 交接文件 → 簡報

目標讀者是**接手的工程師**或**來上課的同事**：他們要在最短時間內知道「這東西做什麼、怎麼跑、從哪讀、哪裡會爆、要找誰」。所以內容要具體到檔案與指令、每個說法可追溯，而不是泛泛的介紹。

產出（預設放在 skill 真實路徑旁的 `output/<slug>-<YYYYMMDD>/`，即 `realpath` 此 skill 目錄後的上一層 `output/`；使用者指定位置則照辦）：

```
output/<slug>-<date>/
├── facts/<repo>/facts.{md,json}   # 掃描結果（客觀事實）
├── modules.md                     # 模組清單與閱讀紀錄（覆蓋率檢查用）
├── handover.md                    # 完整交接文件（source of truth）      ← 選擇性
├── tech-catalog.md                # 技術目錄＋每項技術的要點            ← 選擇性
├── ds-essentials.md               # 只選 DS 章節、沒選簡報時才有         ← 選擇性
├── deck.json                      # 簡報規格                            ← 選了簡報才有
├── <slug>.pptx                    # 簡報                                ← 選擇性
└── preview/                       # PDF + 每頁 PNG（QA 用）
```

`<SKILL>` 以下指本 skill 目錄（啟動時告知的 base directory）。

## 1. 確認範圍（只問真正需要的）

**先問產出與對象（每次啟動都要做，除非請求裡已經講清楚）**：讀 `<SKILL>/references/audience.md`，用**一次** AskUserQuestion 問兩題：

1. 要產出哪些文件（多選）：交接文件 handover.md／技術目錄 tech-catalog.md／教學簡報 .pptx／Data Science Essentials 章節。
2. 文件要給哪個對象角色（單選，可選 Other 自行輸入）：新進 / 接手工程師（預設）／數據科學家 / ML 工程師／PM / 主管／非技術同仁；使用者輸入的其他角色（SRE、資安稽核、實習生…）依 audience.md 找最接近的設定再微調。

之後的步驟只做被選中的產出（相依規則見 audience.md 第 2 節），內容深度、頁數與用詞依角色調整（audience.md 第 3 節）。把「對象｜產出」寫在每份產出的開頭。

**其他範圍資訊**：

- 解析 repo：本機路徑直接用；只給名稱時先在目前目錄與其上層、使用者常用的專案目錄找；git URL 用 `git clone --filter=blob:none` 到 output 目錄下的 `repos/`（保留 git 歷史，掃描需要它）。找不到或有歧義才問。
- 多個 repo → **一份總覽 + 各 repo 章節**（使用者已選定的偏好），除非他明確要分開。
- 下列資訊程式碼裡讀不到，盡量和上面兩題放在同一次提問（AskUserQuestion 最多 4 題），**一次問完**；使用者不知道或想略過時，用「**待確認**」佔位繼續做，不要卡住：
  - 交接人 / 接手人 / 簡報日期
  - 有沒有程式碼外的脈絡要補：進行中的工作、已知地雷、部署環境與權限、相關文件或 Jira/Confluence 連結
- 若對話中已有這些資訊就直接用，不要重問。

## 2. 掃描事實

對每個 repo：

```bash
python3 <SKILL>/scripts/scan_repo.py <repo_path> --out <outdir>/facts/<repo_name>
```

它只讀不跑：目錄樹、語言與行數、相依與 scripts、Docker/CI/k8s 檔、環境變數**名稱**、TODO/FIXME、git 歷史（貢獻者、hotspot、近期 commit、tag）。先讀 `facts.md` 掌握全貌，再決定要深入讀哪些檔案。

`sensitive_files_present` 列出的檔案（`.env`、金鑰…）**不要開啟、不要引用內容**；簡報與文件只能出現變數名稱與「向誰索取」。簡報常被轉寄，洩漏一次就收不回來。

## 3. 建立模組清單（不可跳過）

掃描與 git 活躍度會把注意力拉向「最近最常改的模組」，但接手的人要負責**所有**模組——安靜但仍在服務流量的模組（例如穩定運作多年的規則引擎、資料正規化）最容易被漏掉，交接後出事也最沒人懂。所以在深入閱讀前，先列出完整清單：

- 從程式碼找出每個功能單元：web 框架的 router / blueprint / controller 掛載點（例：FastAPI `include_router`）、獨立的 package 或 app 目錄、排程 / worker / CLI 進入點、notebook 或 pipeline job。
- 對每個模組記錄到 `modules.md`：名稱、對外介面（端點 / 指令 / job）、主要檔案與行數、外部依賴（DB、API、模型服務）、性質（**核心邏輯** / 薄轉接層 / 已停用）、閱讀深度（已深讀 / 已略讀）。
- **核心邏輯**模組（有自己的演算法、規則、資料表或多步流程）一律要深讀；薄轉接層（只是轉呼叫另一個服務）可以合併成一節簡述，但要寫出它轉呼叫誰。
- 性質判斷依程式碼，不依 commit 數：一個 1000 行、規則密集但兩年沒改的檔案，比一個常改的設定檔更需要解釋。

## 4. 深入閱讀（品質的關鍵）

掃描只告訴你「有什麼」，交接要的是「怎麼運作」。依序讀：

1. README / CLAUDE.md / docs —— 但要和程式碼交叉比對，文件常過時；不一致處寫進「已知問題」。
2. 進入點與啟動流程（facts 的 `entrypoints`、package scripts、Dockerfile CMD、k8s command）、設定如何載入。
3. **逐一深讀 modules.md 中每個核心邏輯模組**：從它的對外介面順著呼叫鏈讀到最底層（外部呼叫、DB 讀寫、回傳組裝），記下經過的 `file:line`，並整理出：
   - 這個模組解決什麼問題、輸入輸出長什麼樣（看 request/response model 或範例）
   - 內部的主要元件與它們之間的呼叫關係 —— 這就是該模組架構圖的原料
   - 關鍵規則 / 演算法 / 門檻值、錯誤時的行為（丟例外？回空值？）
   - 演算法的**應用場景素材**：程式裡的分支（命中 / 未命中、有資料 / 冷啟動、門檻內 / 外）、公式的實際係數、切分與評估方式——之後要用來編一個具體的人走一遍
   - 檔案底部的 `__main__`、測試、範例資料 —— 最適合當教學素材與練習
4. 對外依賴總表（DB、API、模型服務、queue、儲存）與各自被哪些模組使用 —— 這決定系統架構圖。
5. 部署：Dockerfile、CI workflow、Helm/k8s manifest、排程設定、資料庫 migration。
6. hotspot 檔案與 TODO、`git log` 里程碑與近期方向（必要時 `git log --stat -- <path>`）。

讀完後回到 `modules.md`，確認每個核心模組都標「已深讀」。

多 repo 時，特別找 repo 之間的連結證據：設定中的 URL/host/port、client 程式、共用套件、docker-compose、README 互相提及。關係是推測的就標「（推測）」。

大型 repo 可平行派 Explore subagent 各讀一個模組或一個 repo（請它回報上述各點與 `file:line`），但結論要自己整合並抽查關鍵行號。

可以執行安全的唯讀指令（`--help`、`git log`、列出檔案）；**不要**執行會連外、寫入資料、部署或需要憑證的指令。測試指令沒有實際跑過，就在文件中註明「未實際驗證」。

## 5. 整理技術目錄（tech-catalog.md）

> 只在選了 tech-catalog.md 時產出文件；沒選時仍要做盤點，供簡報的「用到哪些技術？」頁使用。

依 `<SKILL>/references/tech-catalog.md` 盤點專案實際用到的技術（相依檔 + 實際 import + 建置 / 雲端設定 + 外部服務 + 演算法），分成核心 / 常用 / 周邊三級，寫成：

1. **目錄**：一張表（分類｜技術｜版本｜等級｜在本專案做什麼），技術名稱連到下方小節。
2. **每項技術的要點**：核心技術各一節（一句話、在本專案哪裡用、關鍵設定、3–6 個要點、注意事項、延伸閱讀）；常用技術精簡一節；周邊只出現在目錄表。
3. **宣告了但沒用到**的套件另列一表。

要點要寫「接手的人在這個專案裡需要懂的」，連到實際的設定值與 `file:line`；演算法的公式與應用場景留在 handover.md，這裡連過去即可。

## 5.5 挑出數據科學家必備技術並網路查證

> 只在選了 Data Science Essentials 時做。沒選簡報時改成獨立的 `ds-essentials.md` 文件。

依 `<SKILL>/references/ds-essentials.md`：從技術目錄挑 5–8 個「數據科學家必須懂」的方法（ML 方法、檢索排序、評估與實驗設計、調參、LLM 應用、MLOps 觀念），每項用 WebSearch 查 1–2 個權威來源（論文、官方文件、教科書）確認定義、公式與預設值。之後在教學簡報加一個 **Data Science Essentials** 章節：**技術解說用英文，「在本專案」的連結用中文**，最後附 References 頁。

## 6. 寫 handover.md

> 只在選了 handover.md 時產出；章節取捨依對象角色（audience.md）。

依 `<SKILL>/references/handover-outline.md` 的章節撰寫（使用繁體中文，技術名詞保留原文）。原則：

- 每個重要說法附來源 `path:line` 或 commit hash；推測標「（推測）」；不知道標「**待確認**」並彙整到最後一節。寫錯比寫不知道更糟。
- 具體勝過完整：「`uv run uvicorn bookstore.app:app --reload` 啟動，預設讀 `settings/dev.toml`」勝過「使用 uvicorn 啟動」。
- 「核心流程」一章要涵蓋 modules.md 中**每個**核心邏輯模組，各自一個小節（介紹、主要檔案表、流程步驟附 `file:line`、教學重點）。
- 學習路徑要可執行：第一天讀哪 3 個檔案、第一週可以做哪個小任務（優先找能不靠外部服務完成的，例如把 `__main__` 樣本改寫成測試）。

## 7. 設計簡報並轉成 deck.json

> 只在選了教學簡報時做；頁數與骨架依對象角色調整（audience.md）。講者備註只指向這次有產出的文件，沒有就指向程式碼 `file:line`。

先讀 `<SKILL>/references/teaching-style.md`（**簡報風格，預設採用**），再讀 `<SKILL>/references/deck-spec.md`（格式與各頁型）。範例：`<SKILL>/assets/example-deck.json`。

風格重點（細節與範例頁序見 teaching-style.md）：

- **不要生硬**：標題寫成問句或口語（「v1 為什麼不夠用？」「抓到一個 bug：…」），講者備註用口語，適時放一個可以問台下的問題。
- **架構簡單帶過**：開場用「先看結果」（輸入 → 輸出）＋一張「一張圖看懂」的總覽圖，最多再一頁講模型 / 資料怎麼到線上。目錄導覽、環境變數表、專案歷程、各模組架構圖都留在 handover.md，不上簡報。
- **演算法講細**：簡報的主體是演算法。每個演算法環節依序講：先備知識（看懂後面所需的概念，例如 YOLO 的 conf / NMS、餘弦相似度、隨機遊走）→ 為什麼這樣設計（舊做法哪裡不夠、有數據就放數據）→ 規則或公式本身（決策圖、逐項白話解釋的表）→ **應用場景**（用一個具體的人 / 一筆資料走一遍）→ 容易忽略的細節或 bug。
- **其餘活潑呈現**：已知問題做成「踩雷地圖」cards（簡報者是作者時改成「下一步：還能更好的地方」），部署縮成一頁 flow，結尾是單純的「Q & A」頁，**不放**「接下來想一起做的事」「還沒有答案的問題」這類清單。**不放**「第一週任務」「學習路徑」頁（留在 handover.md）。
- 使用者明確要「正式交接版」時，才改回完整的章節式簡報（系統架構圖 + 每個核心模組的架構圖與流程 + 環境變數 + 歷程 + 交接清單）。

**應用場景頁**（每個演算法環節至少一張，標題以「應用場景①…」編號）：

- 用虛構但合理的角色與 ID（`U1`、`C101`），在 `diagram` 上畫出輸入 → 每一步的中間結果 → 輸出，方塊 desc 寫出該步的實際值（「B990 不在圖上 → 剔除」）。
- 每個重要分支各一個場景（例：正常命中 vs. 冷啟動走 fallback），在 callout 點出這個分支實際何時會發生；多個分支並列時可用 `table`（「四張照片，四種結局」）。
- 公式類用 `table` 逐步代入示意數字（步驟｜公式（出處）｜本例），算術要正確；需要全體分布才能算的值（Z-score、百分位）寫「假設」。
- 門檻 / 參數類用「同一筆輸入、兩組參數」對照，讓聽眾看到參數改變了什麼。
- 訓練 / 評估類畫一位使用者的資料怎麼切、指標怎麼算（例：Recall@10 = 命中 1 / 2 = 0.5），並把指標翻成白話。
- 頁面與 handover.md 都標「示意」，講者備註附上每一步對應的 `file:line`，並說明場景揭露的設計特性或問題。

通用規則：

- 一頁一個重點；細節放 `notes` 並指回 handover.md 章節。
- 圖上每個方塊與箭頭都要能在程式碼裡找到對應；推測內容標「（推測）」。排版要領見 deck-spec.md 的 diagram 一節（相鄰格子、避免交叉、label 短）。
- 純文字 bullets 頁不超過四分之一；規則對照用 `table`、決策用 `diagram`、流程用 `flow`。
- 數字與名稱都來自 facts / 程式碼，不要四捨五入成好看的數字（應用場景的示意數字例外，但要標「示意」，公式與係數仍須與程式一致）。
- 不使用任何真實個資（姓名、證件、帳號）當例子；即使 repo 裡有真實資料也只描述格式。
- `footer` 設為「<專案名> 導讀｜<YYYY-MM>」（正式交接版用「交接」）。

## 8. 產生與檢查

```bash
uv run --with python-pptx python <SKILL>/scripts/build_deck.py <outdir>/deck.json -o <outdir>/<slug>.pptx
bash <SKILL>/scripts/preview.sh <outdir>/<slug>.pptx
```

- 有 `FIT WARNINGS` 就拆頁或精簡，重建到沒有警告。
- 用 Read 實際看 `preview/slide-*.png`（頁數多時可先拼成總覽圖；架構圖頁用 `pdftoppm -r 90` 單獨放大看），檢查：文字溢出或重疊、英文單字被斷行、箭頭交叉或穿過方塊、箭頭標籤蓋到方塊、空白過多、待確認標記是否醒目。
- 覆蓋率檢查：簡報有 Data Science Essentials 章節（5–8 項，英文解說＋中文專案連結＋References 頁），每項的講者備註都有來源 URL；相依檔中實際被 import 的套件、建置設定裡的雲端服務、以及每個演算法，都出現在 tech-catalog.md 的目錄；modules.md 中每個核心模組都在 handover.md 有小節；簡報裡每個演算法環節（推論、規則 / 評分、訓練、評估）都有先備知識或規則頁＋應用場景頁。
- 預覽用 Heiti TC 代替微軟正黑體，字寬略有差異，留一點餘裕。
- 修正後重建，直到沒有明顯問題。

## 9. 交付

回報：這次的對象角色與產出清單、每份產出（pptx、handover.md、tech-catalog.md、ds-essentials.md 中有產出的）的完整路徑、頁數、技術目錄的項目數（核心 / 常用 / 周邊）、以及**待確認事項清單**（請原負責人補的部分）。提醒 handover.md 是可持續更新的版本，簡報只是講解用快照；之後修改內容時先改 handover.md / deck.json 再重建。
