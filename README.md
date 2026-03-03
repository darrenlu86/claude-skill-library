# gen-pptx — Claude Code PPTX Skill

A Claude Code skill that generates professional PPTX presentations from JSON content definitions.

## Features

- **19 slide types**: title, section, bullet, flow, architecture, table, cards, timeline, comparison, and more
- **Dynamic layout**: text measurement prevents overflow; boxes auto-size to content
- **Auto-decoration**: geometric motifs fill sparse slides at 40% opacity
- **Template support**: use your own corporate `.pptx` template
- **Footer system**: auto page number, date, and copyright on every content slide

## Quick Start

### 1. Install as Claude Code Skill

Copy the files into your Claude Code skills directory:

```
~/.claude/skills/gen-pptx/
├── pptx_engine.py      # Layout engine
├── generate_pptx.py    # JSON → PPTX runner
├── SKILL.md            # Skill instructions for Claude
└── config.json         # Your local config (create from example)
```

### 2. Configure

```bash
cp config.example.json config.json
```

Edit `config.json` with your settings:

```json
{
  "template": "path/to/your-template.pptx",
  "footer_text": "Your Company Name",
  "output_dir": "path/to/output"
}
```

### 3. Install Dependencies

```bash
pip install python-pptx
```

### 4. Use with Claude Code

Ask Claude to create a presentation:

> "Create a presentation about [topic] using the gen-pptx skill"

Or run directly with a JSON content file:

```bash
python generate_pptx.py content.json output.pptx
```

## JSON Content Format

```json
{
  "template": "optional/override.pptx",
  "footer_text": "Optional Override",
  "slides": [
    {"type": "title", "title": "My Presentation", "subtitle": "2024"},
    {"type": "bullet", "title": "Key Points", "columns": [
      {"header": "Topic", "header_color": "BLUE", "items": ["Point 1", "Point 2"]}
    ]},
    {"type": "end", "title": "Thank You", "subtitle": "Q&A"}
  ]
}
```

See `SKILL.md` for full documentation of all 19 slide types and the presentation planning framework.

## Config Priority

1. **JSON content file** fields (per-presentation override)
2. **config.json** (user defaults)
3. **Fallback** (blank presentation, no footer)

## License

MIT
