# 數據科學家必備技術（Data Science Essentials）章節

使用者需求：把數據科學家必須懂的技術挑出來解說，上網查證後做成簡報，技術部分用英文。預設：**併入教學簡報**；**技術解說用英文，專案連結用中文**。

使用者在第 1 步勾選「Data Science Essentials」才做。有選簡報 → 併入簡報（下述格式）；沒選簡報 → 改寫成獨立的 `ds-essentials.md`（每項一節：What & Why、How it works、Watch out 用英文，「在本專案」用中文，最後列 References）。對象是 PM / 主管或非技術同仁時，每項只留 What & Why 與「在本專案」。

目的：讓聽眾離開時不只懂這個專案，也帶走可以用在別的專案的方法。演算法章節講「這個專案怎麼做」，這個章節講「這個方法本身是什麼」。

## 1. 挑選

從 tech-catalog.md 的目錄裡，挑出符合下列任一條件的方法（不是框架或雲服務）：

- 機器學習 / 深度學習方法：embedding、Node2Vec、object detection、fine-tuning、contrastive / triplet loss…
- 資訊檢索與排序：BM25 / TF-IDF、vector search、two-stage retrieval、reranking…
- 評估與實驗設計：Recall@K、NDCG、ROC AUC、mAP、precision / recall trade-off、time-based split、data leakage、validation set…
- 最佳化與調參：Bayesian optimization（TPE）、early stopping…
- LLM 應用方法：LLM-as-a-judge、structured scoring rubric、prompting…
- MLOps 必備觀念：model registry、promotion gate、reproducibility…

每份簡報挑 **5–8 項**，以「接手的人換到別的專案也一定會用到」為優先。框架用法（FastAPI、Docker）不放這裡。

## 2. 網路查證（必做）

每一項用 WebSearch 查 1–2 個權威來源，確認定義、公式與常見預設值，並記下 URL：

- 優先順序：原始論文（arXiv、ACL / KDD / NeurIPS 等）→ 官方文件（scikit-learn、PyTorch、PyG、Ultralytics、Hyperopt、MLflow、pgvector…）→ 知名教科書或課程（Manning *Introduction to Information Retrieval*、Stanford CS224W…）。
- 不用內容農場或沒有作者的部落格當唯一來源。
- 查到的預設值和專案設定不同時，兩個都寫（「BM25 default b = 0.75; this project uses 0.4」）。
- 來源 URL 放在每頁的講者備註，並在章節最後放一頁 References。

## 3. 每項技術的頁面（英文）

每項 1–2 頁，放在 `section`「Data Science Essentials」之下，位置在演算法章節之後、上線章節之前：

| 頁 | 頁型 | 內容（英文） |
|---|---|---|
| 1 | `two_column` 或 `table` | **What & Why**：one-line definition、the problem it solves；**How it works**：3–4 個步驟或關鍵公式（白話解釋每個符號） |
| 2（選用） | `table` / `cards` | **When to use / Pitfalls**：適用情境、常見錯誤、和替代方法的比較 |

- 標題格式：`<Technique>: <one-line hook>`（例：`BM25: Rare Terms Matter More`）。
- 內文全部英文、短句；公式用純文字（`IDF = ln((N − n + 0.5)/(n + 0.5) + 1)`）。
- 每頁的 `callout` 用**中文**寫「在本專案」：這個方法在專案裡用在哪、設定值多少、對應簡報哪一頁（例：「在本專案：BM25 算標籤—書籍邊的權重，k1 = 1.2、b = 0.4，見第 18 頁」）。
- 講者備註：英文的補充解釋 1–3 句、來源 URL、中文的專案出處（`file:line`、Cell N）。
- 字數預算同其他頁：表格每格一短句；圖上文字要少。

## 4. 章節結構

1. `section`：`Data Science Essentials`，副標中文一句（「換到別的專案也用得到的方法」）。
2. 一頁 `table` 總覽：Technique｜Category｜Used for（英文）＋最後一欄「本專案」（中文，3–8 字）。
3. 每項技術 1–2 頁。
4. 一頁 References（`bullets`，英文，作者／標題／年份＋網址）。

## 5. tech-catalog.md 的對應

目錄表中被選為 DS 必備的技術，等級欄寫「核心・DS」或「常用・DS」，小節最後加「**References**」列出查證來源。
