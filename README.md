# project-savior

Claude Code skill **project-handover**：指定 repo → 分析程式碼與 git 歷史 → 產出交接/教學文件 `handover.md`、技術目錄 `tech-catalog.md` 與 `.pptx` 簡報（可選擇產出項目與對象角色）。

- Skill 原始碼：`project-handover/`（已 symlink 至 `~/.claude/skills/project-handover`）
- 產出位置：`output/<slug>-<date>/`（已列入 .gitignore，不會進版控）

## 使用

在 Claude Code 中：

```
/project-handover ~/code/my-service
幫我把 my-api 跟 my-model 整理成教學簡報
```

## 結構

```
project-handover/
├── SKILL.md                       # 工作流程
├── scripts/
│   ├── scan_repo.py               # 唯讀掃描 repo 事實（stdlib）
│   ├── build_deck.py              # deck.json → pptx（uv run --with python-pptx）
│   └── preview.sh                 # pptx → PDF/PNG 預覽（LibreOffice + pdftoppm）
├── references/
│   ├── audience.md                # 產出選擇與對象角色
│   ├── handover-outline.md        # 交接文件章節與簡報大綱
│   ├── teaching-style.md          # 簡報風格（預設）
│   ├── tech-catalog.md            # 技術目錄寫法
│   ├── ds-essentials.md           # Data Science Essentials 章節
│   └── deck-spec.md               # deck.json 格式與頁型
└── assets/example-deck.json       # 範例（涵蓋所有頁型）
```

相依：python3、uv、git；預覽需 LibreOffice (`soffice`) 與 poppler (`pdftoppm`)。
