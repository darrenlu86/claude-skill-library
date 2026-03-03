# gen-pptx — PPTX Presentation Generator

A Claude Code skill that generates professional PPTX presentations from JSON content definitions.

## Features

- **19 slide types**: title, section, bullet, flow, architecture, table, cards, timeline, comparison, and more
- **Dynamic layout**: text measurement prevents overflow; boxes auto-size to content
- **Auto-decoration**: geometric motifs fill sparse slides at 40% opacity
- **Template support**: use your own corporate `.pptx` template
- **Footer system**: auto page number, date, and copyright on every content slide

## Setup

```bash
cp config.example.json config.json
```

Edit `config.json`:

```json
{
  "template": "path/to/your-template.pptx",
  "footer_text": "Your Company Name",
  "output_dir": "path/to/output"
}
```

### Dependencies

```bash
pip install python-pptx
```

## Usage

Ask Claude:

> "Create a presentation about [topic] using the gen-pptx skill"

Or run directly:

```bash
python generate_pptx.py content.json output.pptx
```

## JSON Format

```json
{
  "slides": [
    {"type": "title", "title": "My Presentation", "subtitle": "2024"},
    {"type": "bullet", "title": "Key Points", "columns": [
      {"header": "Topic", "header_color": "BLUE", "items": ["Point 1", "Point 2"]}
    ]},
    {"type": "end", "title": "Thank You", "subtitle": "Q&A"}
  ]
}
```

See `SKILL.md` for full documentation of all 19 slide types and the planning framework.

---

# gen-pptx — PPTX 簡報產生器

Claude Code 技能，從 JSON 內容定義產生專業 PPTX 簡報。

## 功能特色

- **19 種投影片類型**：封面、段落、條列、流程圖、架構圖、表格、卡片、時間軸、比較等
- **動態排版**：根據文字內容自動測量並調整尺寸，避免溢出
- **自動裝飾**：內容較少的頁面自動加入幾何圖案（40% 透明度）
- **範本支援**：可使用自訂企業 `.pptx` 範本
- **頁尾系統**：自動在每頁內容頁加入頁碼、日期、版權文字

## 設定

```bash
cp config.example.json config.json
```

編輯 `config.json`：

```json
{
  "template": "你的範本路徑.pptx",
  "footer_text": "你的公司名稱",
  "output_dir": "輸出目錄路徑"
}
```

### 安裝套件

```bash
pip install python-pptx
```

## 使用方式

請 Claude：

> 「用 gen-pptx 技能幫我製作一份關於 [主題] 的簡報」

或直接執行：

```bash
python generate_pptx.py content.json output.pptx
```

## 設定優先順序

1. **JSON 內容檔** 中的欄位（單次簡報覆寫）
2. **config.json**（使用者預設值）
3. **預設值**（空白簡報、無頁尾）
