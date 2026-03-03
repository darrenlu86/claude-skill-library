# Claude Skill Library

A collection of reusable [Claude Code](https://docs.anthropic.com/en/docs/claude-code) skills — AI-driven automation tools that extend Claude's capabilities.

## Skills

| Skill | Description |
|-------|-------------|
| [gen-pptx](./gen-pptx/) | Generate professional PPTX presentations from JSON |

## Installation

1. Copy the skill folder into your Claude Code skills directory:

```bash
# Example: install gen-pptx
cp -r gen-pptx ~/.claude/skills/gen-pptx
```

2. Configure the skill (if needed):

```bash
cd ~/.claude/skills/gen-pptx
cp config.example.json config.json
# Edit config.json with your settings
```

3. Use it in Claude Code — the skill is automatically available.

## Contributing

Each skill lives in its own directory with:

```
skill-name/
├── SKILL.md              # Skill instructions (read by Claude)
├── config.example.json   # Example config (user copies to config.json)
└── *.py                  # Implementation files
```

## License

MIT

---

# Claude Skill Library（中文）

可重複使用的 [Claude Code](https://docs.anthropic.com/en/docs/claude-code) 技能庫 — 透過 AI 驅動的自動化工具擴展 Claude 的能力。

## 技能列表

| 技能 | 說明 |
|------|------|
| [gen-pptx](./gen-pptx/) | 從 JSON 內容定義產生專業 PPTX 簡報 |

## 安裝方式

1. 將技能資料夾複製到 Claude Code 的 skills 目錄：

```bash
# 範例：安裝 gen-pptx
cp -r gen-pptx ~/.claude/skills/gen-pptx
```

2. 設定技能（如需要）：

```bash
cd ~/.claude/skills/gen-pptx
cp config.example.json config.json
# 編輯 config.json 填入你的設定
```

3. 在 Claude Code 中直接使用，技能會自動載入。

## 貢獻方式

每個技能獨立放在各自的資料夾中，結構如下：

```
skill-name/
├── SKILL.md              # 技能說明文件（供 Claude 讀取）
├── config.example.json   # 範例設定檔（使用者複製為 config.json）
└── *.py                  # 實作程式碼
```

## 授權

MIT
