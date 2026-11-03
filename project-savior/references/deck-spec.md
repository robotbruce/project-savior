# deck.json 格式（build_deck.py 的輸入）

```json
{
  "title": "簡報標題（也是預設頁尾）",
  "footer": "頁尾文字（選填）",
  "theme": { "primary": "#1E3A5F", "accent": "#F28C28", "font": "Microsoft JhengHei" },
  "slides": [ { "type": "...", "notes": "講者備註（字串或字串陣列）", ... } ]
}
```

- 所有文字欄位都支援行內標記：`**粗體**`、`` `程式碼/路徑` ``（等寬、強調色）。
- 每張投影片都可以有 `notes`：教學時要講的話、交接細節，放這裡而不是塞滿畫面。
- 項目清單（`bullets`）元素可以是：
  - 字串；前面每兩個空白 = 縮一層（`"  子項目"`）
  - `{"text": "...", "level": 0, "sub": ["子項目", ...]}`
- `callout`（選填，bullets / two_column / table / flow / architecture / cards 皆可用）：底部橘色重點提示條，一句話。
- `theme` 欄位：primary、accent、text、muted、light、border、code_bg、code_fg、font、mono，只需覆寫想改的。

## 投影片類型

| type | 必填欄位 | 選填 | 適用 |
|---|---|---|---|
| `diagram` | `title`, `nodes`, `edges` | `groups`, `cols`, `rows`, `legend`, `callout` | **架構圖**（系統 / 模組），見下方專節 |
| `title` | `title` | `subtitle`, `author`, `date` | 封面 |
| `agenda` | `items`（字串或 `{text, desc}`） | `title`（預設「大綱」） | 大綱；>6 項自動兩欄 |
| `section` | `title` | `number`, `subtitle` | 章節分隔頁 |
| `bullets` | `title`, `bullets` | `subtitle`, `callout`, `max_size` | 一般重點；建議 ≤ 6 個主項目 |
| `two_column` | `title`, `left`, `right` | `subtitle`, `callout` | 對照（本機 vs 部署、現況 vs 目標）。每欄 `{heading, bullets}` 或 `{heading, code}` 或 `{heading, text}` |
| `table` | `title`, `columns`, `rows` | `col_widths`（相對比例）, `callout` | 設定、環境變數、API 一覽；≤ 12 列 |
| `code` | `title`, `code` | `caption` | 關鍵程式片段、指令；≤ 18 行、每行 ≤ 90 字元 |
| `flow` | `title`, `steps`（字串或 `{label, desc, highlight}`） | `callout` | 請求流程、資料管線、部署流程。畫成「數字節點」：≤ 7 步橫排（圓形編號 + 標籤 + 一句話），8 步以上改直式。**label ≤ 8 字、desc 一行 ≤ 16 字**，超過會出 FIT WARNING；細節放 notes |
| `architecture` | `title`, `layers`：`[{name, items:[字串或 {name, desc, highlight}]}]` | `arrows`（預設 true）, `callout` | 分層架構圖（上 = 使用者端，下 = 資料/模型/基礎設施）；≤ 5 層、每層 ≤ 5 元件 |
| `cards` | `title`, `cards`：`[{title, body(字串或陣列), tag, highlight}]` | `columns`, `callout` | 模組介紹、風險/待辦（tag 放優先度）、學習資源 |
| `stats` | `title`, `items`：`[{value, label}]` | `bullets` | 專案規模數字（模組數、行數、服務數、commit 數） |
| `timeline` | `title`, `milestones`：`[{date, label, desc, status}]` | | 專案歷程與 roadmap；`status: "todo"` 顯示空心點 |
| `image` | `title`, `path`（相對於 deck.json） | `caption` | 既有架構圖、截圖 |
| `closing` | | `title`（預設 Q & A）, `bullets` | 結尾：交接清單、窗口、Q&A |

## 版面規則（產生器會估算並在輸出時警告）

- 執行後若出現 `FIT WARNINGS`，代表該頁即使縮到最小字級仍可能溢出 → 拆頁或精簡文字，不要忽略。
- 字級自動在範圍內縮放（項目 22→12pt、卡片內文 15→10pt），但字越小越難讀；寧可多一頁。
- `architecture` 的元件名稱用實際模組/服務名（`bookstore.app`），`desc` 寫一句它做什麼或技術（`FastAPI 路由`）。
- 圖表類（diagram / flow）是本 skill 最有價值的頁面，比 bullets 更能幫助理解。`architecture`（固定分層方塊）只適合非常簡單的分層說明；有實際呼叫關係時一律用 `diagram`。

## diagram：架構圖

以格線擺放節點，產生原生 PowerPoint 圖形（接手者可直接在 PowerPoint 裡修改）。

```json
{
  "type": "diagram", "title": "發票辨識架構", "cols": 5, "rows": 3,
  "groups": [ {"label": "Tag API", "cells": [1, 0, 4, 3], "color": "primary"} ],
  "nodes": [
    {"id": "cl", "label": "呼叫端", "desc": "記帳 App（推測）", "col": 0, "row": 1, "kind": "client"},
    {"id": "rt", "label": "invoice/router.py", "desc": "/invoice/analyze", "col": 1, "row": 1, "kind": "module", "highlight": true},
    {"id": "ap", "label": "API Gateway", "desc": "取使用者資料", "col": 1, "row": 0, "kind": "external"},
    {"id": "my", "label": "MySQL", "desc": "feedback log", "col": 1, "row": 2, "w": 2, "kind": "db"}
  ],
  "edges": [
    {"from": "cl", "to": "rt"},
    {"from": "rt", "to": "ap", "label": "① 取資料"},
    {"from": "rt", "to": "my", "label": "③ log", "style": "dashed"}
  ]
}
```

- **節點** `nodes[]`：`id`、`label`（元件名，用程式裡的真名）、`desc`（一句話：做什麼 / 技術 / 端點）、`col`/`row`（從 0 起算）、`w`/`h`（跨幾格，預設 1）、`kind`、`highlight`（橘框強調）、`size`（最大字級，預設 14）。
- **kind**：`client` 呼叫端（深色）、`module` 本服務模組、`service` 內部元件/步驟、`external` 外部服務（虛線框）、`db` 資料庫（圓柱）、`storage` S3/檔案、`queue` 佇列/排程、`note` 說明便利貼。
- **群組** `groups[]`：`cells: [col起, row起, col迄(不含), row迄(不含)]`，畫虛線框，左側直書標籤（例：「Tag API」「外部服務」「資料層」）。
- **連線** `edges[]`：`from`、`to`、`label`（短，≤ 6 字最好）、`style: "dashed"`（非同步 / 資料流 / 次要）、`both: true`（雙向）、`side: "h" | "v"`（強制從左右邊或上下邊連出，用來避免交叉）、`color`。
- 預設在圖下方顯示圖例；`"legend": false` 關閉。

**排版要領**（做完務必看預覽確認）：

1. 先決定方向：系統圖用「上→下」分層（呼叫端 / 模組 / 外部服務 / 資料層各一列，用 groups 框起來）；單一模組用「左→右」（呼叫端 → 入口 → 處理元件 → 外部依賴）。
2. 有連線的兩個節點盡量放相鄰格子（上下或左右），長距離斜線容易穿過其他方塊。
3. 一個節點連到多個目標時，把目標排在它的同一側，避免箭頭交叉；仍交叉就調換欄位或加 `side`。
4. 系統圖節點多時，資料層可以不畫連線，改在 `desc` 註明「search · report 使用」，讓圖保持可讀。
5. 每列 ≤ 6 個節點、全圖 ≤ 18 個節點；更多就拆成「系統總覽 + 各模組圖」。
6. 有 `callout` 會壓縮圖的高度；節點多的圖把說明放講者備註。
7. **文字預算**：方塊的 `desc` 最多 2 短行、約 22 字（超過會出 FIT WARNING）。圖是給人「看」的，完整說明放講者備註。
8. **應用場景圖**沿用 diagram：label 寫步驟或元件名，desc 寫「這個例子在這一步的實際值」；結果用 `note` 節點，被剔除 / 走到 fallback 的關鍵步驟設 `highlight`。label 盡量短（≤ 12 字），細節放 desc，避免方塊被長名稱擠到縮字。
