---
name: gen-pptx
description: "Generate a technical presentation (PPTX). Claude researches the topic, writes a JSON content file, then calls the generic generator."
argument-hint: "[topic or instructions]"
disable-model-invocation: true
allowed-tools:
  - Read
  - Edit
  - Write
  - Bash
  - Glob
  - Grep
---

# gen-pptx — 簡報產生器 v4

## Overview

Generate a PPTX presentation for any topic. This skill uses a **data-driven approach**:
Claude writes a JSON content file, then calls the generic `generate_pptx.py` runner.
**No per-topic Python code is ever generated. No company/user binding.**

User instructions: $ARGUMENTS

## Workflow (MUST follow this order)

1. **Research** — Read relevant source files, README, configs, data files
2. **Extract** — Identify key facts, numbers, processes, decisions (see §Content Extraction)
3. **Outline** — Map extracted content to slide types using §Layout Selection Rules
4. **Validate** — Check against §Rhythm & Pacing constraints
5. **Write JSON** — `{project_dir}/slides_content.json`
6. **Run generator**:
   ```bash
   python "~/.claude/skills/gen-pptx/generate_pptx.py" "{project_dir}/slides_content.json"
   ```
7. **Verify** output and report slide count

---

## Presentation Planning Framework

This framework governs ALL content and layout decisions. Follow it strictly.

### §1 Content Extraction — How to read raw material

When given a folder or codebase, systematically extract these 7 categories:

| # | Category | What to look for | Example |
|---|----------|-------------------|---------|
| 1 | **Numbers** | Counts, percentages, sizes, KPIs | "2,488 records", "13 users", "99.9% uptime" |
| 2 | **Processes** | Sequential steps, pipelines, workflows | "Excel → Python → JSON → Browser" |
| 3 | **Structures** | Layers, modules, components, hierarchy | "Frontend / Backend / Data layer" |
| 4 | **Comparisons** | Before/after, alternatives, trade-offs | "Manual vs Automated", "Option A vs B" |
| 5 | **Decisions** | Why a choice was made, key principles | "Static HTML to avoid server dependency" |
| 6 | **Lists** | Features, requirements, tech stack items | "Bootstrap, jQuery, Chart.js, DataTables" |
| 7 | **Milestones** | Phases, timeline, roadmap, versions | "Q1: Plan → Q2: Build → Q3: Test → Q4: Launch" |

**Rule: Every slide MUST trace back to at least one extracted fact. Never invent content that isn't in the source material.**

### §2 Slide Budget — How many slides to generate

| Presentation length | Content slides | Total (with title/agenda/end) |
|---------------------|---------------|-------------------------------|
| 5 min (lightning)   | 4–6           | 7–9                           |
| 10 min (short)      | 8–12          | 11–15                         |
| 20 min (standard)   | 14–18         | 17–21                         |
| 30 min (detailed)   | 20–26         | 23–29                         |

- Default to **10 min / short** when the user doesn't specify length
- If user specifies slide count, follow that directly
- **Structural slides** (title, agenda, section, end) don't count toward content slides

### §3 Layout Selection Rules — Content → Slide Type mapping

Match extracted content to slide types using this priority table. **Always check from top to bottom; use the FIRST match.**

| Content pattern | Primary type | Fallback type | When to use fallback |
|----------------|-------------|---------------|---------------------|
| 3–6 standalone numeric KPIs | `stat` | `cards` | Values need labels longer than 2 words |
| Time-ordered phases/milestones (3–6) | `timeline` | `flow` | No dates, just sequential |
| Ordered procedure with descriptions | `steps` | `flow` | No descriptions, just names |
| Two opposing options/states | `comparison` | `bullet` (2 col) | > 8 items per side |
| System layers/modules (2–5) | `arch` | `bullet` (N col) | > 5 layers |
| Two parallel processes | `dual_flow` | `flow` × 2 slides | > 6 steps per flow |
| Key principles/decisions (2–4) | `highlight` | `bullet` | > 4 items |
| Four-quadrant analysis | `matrix` | `table` | > 2×2 cells |
| Tabular data (rows × columns) | `table` | `bullet` | < 3 columns |
| Important quote or core concept | `quote` | `highlight` | Multiple quotes |
| Interaction between actors | `sequence` | `flow` | < 3 messages |
| File/directory structure | `tree` | `bullet` | Very short tree |
| General categorized info | `bullet` | — | Last resort |

**IMPORTANT: `bullet` is the LAST RESORT, not the default.** Before using `bullet`, re-examine the content and ask: "Does this fit a more specific type?"

### §4 Rhythm & Pacing — Avoid monotony

#### Hard rules (MUST follow)

1. **No 3 consecutive same-type slides** — If you have 3 bullets in a row, convert the middle one to another type
2. **Maximum `bullet` ratio: 40%** — In a 15-slide deck, at most 6 can be `bullet`
3. **Every section must have at least 1 non-bullet slide** — If a section is all bullets, convert one
4. **Visual slides (arch/flow/timeline/matrix/stat/cards/sequence) must be ≥ 25%** of content slides

#### Rhythm pattern templates

Use these patterns as starting points. Letters represent visual weight:

```
T = Text-heavy (bullet, steps, table, tree)
V = Visual (arch, flow, dual_flow, timeline, cards, stat, matrix, comparison, sequence)
H = High-impact (highlight, quote, stat)

Short deck (10 slides):
  title → agenda → T → V → T → H → V → T → V → end

Standard deck (17 slides):
  title → agenda → T → V → T → T → section → V → H → T → V → section → T → V → T → H → end

Long deck (25 slides):
  title → agenda → T → V → T → H → section → V → T → T → V → section → T → V → H → T → section → V → T → T → V → H → T → V → end
```

#### Breaking monotony techniques

When you find yourself writing a 3rd `bullet` slide, use one of these conversions:

| Situation | Convert to | How |
|-----------|-----------|-----|
| Bullet with 3–5 short items that have key insight | `highlight` | Make each item a heading + description |
| Bullet with numeric values mixed in | `stat` or `cards` | Extract the numbers out |
| Bullet comparing two things | `comparison` | Split into left/right |
| Bullet listing sequential items | `flow` or `steps` | Add order arrows |
| Bullet with categorized items | `matrix` | Group into 2×2 |

### §5 Section Structure — Standard presentation skeleton

Every presentation follows this skeleton. Adapt section names to topic.

```
1. OPENING          title + agenda                    (2 slides)
2. CONTEXT          Background, problem, motivation   (1–3 slides)
3. SOLUTION         What was built / proposed          (2–4 slides)
4. HOW IT WORKS     Architecture, process, tech stack  (3–6 slides)
5. DETAILS          Data, features, demo points        (2–5 slides)
6. OPERATIONS       Deployment, update, maintenance    (1–3 slides)
7. FUTURE           Roadmap, next steps, suggestions   (1–2 slides)
8. CLOSING          end slide                          (1 slide)
```

- Use `section` slides between major groups (every 4–6 content slides)
- CONTEXT should include at least 1 `highlight` or `quote` for impact
- HOW IT WORKS must include at least 1 visual type (`arch`/`flow`/`dual_flow`)
- DETAILS should use `table`/`stat`/`cards` if data is available

### §6 Content Density — How much text per slide

| Slide type | Max items/points | Max text per item |
|-----------|-----------------|-------------------|
| `bullet` (1 col) | 6 items | 30 characters |
| `bullet` (2 col) | 5 items per column | 25 characters |
| `highlight` | 4 points | Heading: 15 chars, Desc: 40 chars |
| `flow` / `timeline` | 6 steps | 15 characters per step |
| `arch` | 5 layers | Label: 15 chars, Body: 50 chars |
| `steps` | 5 steps | Title: 15 chars, Desc: 40 chars |
| `stat` | 5 values | Value: 8 chars, Label: 15 chars |
| `table` | 8 rows | 20 characters per cell |

**If content exceeds limits, split into 2 slides rather than shrinking or cramming.**

### §7 Color Assignment — Consistent color strategy

Don't pick colors randomly. Follow this system:

| Rule | Application |
|------|-------------|
| **One accent per section** | All slides in "Architecture" section use BLUE family |
| **Semantic colors** | GREEN = positive/success, RED = negative/risk, ORANGE = warning/WIP |
| **Same-topic = same color** | If "Backend" is BLUE on slide 5, it stays BLUE on slide 12 |
| **Column headers alternate** | 2-col bullet: BLUE + GREEN; 3-col: BLUE + GREEN + ORANGE |
| **Light variants for fills** | Box fills use `_LT` variants; text/accents use the strong color |

### §8 Auto-Decoration — Geometric motifs for sparse slides

The engine automatically adds decorative geometric shapes to slides with **< 45% content coverage**.

**How it works:**
1. After each slide is built, the engine estimates content coverage (shape area / zone area)
2. If coverage < 45%, it finds the emptiest corner (top-right, bottom-right, etc.)
3. Places a themed geometric motif (circles, bars, grid, rings, or arrows) at 40% opacity

**Motif assignment by slide type:**

| Slide type | Motif | Default color |
|-----------|-------|---------------|
| `bullet`, `steps` | Dots (scattered circles) | BLUE_LT |
| `arch`, `cards`, `stat` | Bars (abstract bar chart) | BLUE_LT / CYAN_LT |
| `flow`, `dual_flow`, `timeline` | Arrow (forward arrow) | GREEN_LT / ORANGE_LT |
| `highlight`, `quote`, `sequence` | Rings (concentric circles) | PURPLE_LT / BLUE_LT |
| `table`, `tree`, `matrix` | Grid (2×2 squares) | GRAY_LT / AMBER_LT |

**JSON control:**
```json
{
  "auto_decorate": true,
  "slides": [...]
}
```
- Set `"auto_decorate": false` to disable globally
- Per-slide override: `"accent_color": "GREEN_LT"` changes the motif color
- Title/section/end/agenda slides are never decorated

### §9 Dynamic Text Measurement

The engine uses `measure_text(text, sz_pt, box_w_in)` to estimate text dimensions before rendering:
- Distinguishes CJK full-width characters (1.0× width) from Latin half-width (0.55×)
- Calculates wrap line count for a given box width
- All builders measure content FIRST, then compute layout dimensions

**This means:** box heights, step heights, flow box heights, highlight rows, table rows, etc. all adapt to actual content length. No more fixed sizes that overflow or leave excessive whitespace.

---

## File Locations

| File | Purpose | Per-topic? |
|------|---------|------------|
| `~/.claude/skills/gen-pptx/pptx_engine.py` | Layout engine (colors, fonts, builders) | NO — shared |
| `~/.claude/skills/gen-pptx/generate_pptx.py` | Generic JSON→PPTX runner | NO — shared |
| `{project}/slides_content.json` | Content data for this presentation | YES — per-topic |

Template and footer are **configurable per-presentation** — see §JSON Config below.

## Layout Safe Zone

Slide = 13.33" × 7.50". All builders auto-fill within:

```
CONTENT_TOP = 1.20"  (below title)
CONTENT_BOTTOM = 6.65" (above footer)
CONTENT_H = 5.45"  (usable height)
MARGIN_X = 0.55"
CONTENT_W = 12.20"
```

## Font Size Hierarchy (STRICT — minimum body = 16pt)

| Size | Usage |
|------|-------|
| 40pt | Title slide main title |
| 28pt | Slide title (auto, placeholder) |
| 20pt | Section header in body, subtitle |
| 16pt | ALL body text (MINIMUM for readable content) |
| 12pt | Sequence diagram arrow labels ONLY (sole exception) |

**Rule: No font below 16pt for any readable text. SZ_SUB/SZ_SMALL/SZ_TINY all alias to 16pt.**

## Color Palette (use string names in JSON)

```
BLUE / BLUE_LT    GREEN / GREEN_LT    ORANGE / ORANGE_LT
RED / RED_LT       PURPLE / PURPLE_LT  AMBER / AMBER_LT
CYAN / CYAN_LT     DARK   GRAY / GRAY_LT   WHITE   BLACK
```

Also supports hex: `"#2563EB"`.

## JSON Content Format

```json
{
  "output": "{project_dir}/output.pptx",
  "template": "path/to/template.pptx",
  "footer_text": "Copyright of YourCompany",
  "auto_decorate": true,
  "slides": [
    { "type": "title", ... },
    { "type": "agenda", ... },
    { "type": "bullet", ... },
    ...
    { "type": "end" }
  ]
}
```

### §JSON Config — Top-level fields

| Field | Required | Default | Description |
|-------|----------|---------|-------------|
| `output` | YES | `"output.pptx"` | Output file path |
| `template` | no | Engine default template | Path to .pptx template. Use `""` for blank (no template) |
| `footer_text` | no | _(none)_ | Copyright/footer text. Omit to leave footer empty |
| `auto_decorate` | no | `true` | Add geometric motifs to sparse slides |
| `slides` | YES | — | Array of slide definitions |
```

## Slide Type Reference

### title / section / end
```json
{"type": "title", "title": "Main Title", "subtitle": "Department | Date"}
{"type": "title", "title": "Main Title", "title_color": "WHITE", "sub_color": "WHITE"}
{"type": "section", "title": "Section Name", "subtitle": "Optional"}
{"type": "end", "title": "謝謝各位", "subtitle": "Q&A 時間"}
```
Default text color: **DARK** (visible on light backgrounds). Set `"title_color": "WHITE"` for dark template backgrounds.

### agenda
```json
{"type": "agenda", "title": "簡報大綱", "items": [
  ["1", "Topic Name", "Brief description"],
  ["2", "Topic Name", "Brief description"]
]}
```
Items auto-centered vertically.

### bullet (1~N columns)
```json
{"type": "bullet", "title": "Slide Title", "subtitle": "Optional", "columns": [
  {
    "header": "Header Bar Text",
    "header_color": "BLUE",
    "items": ["Point 1", "Point 2", ["Sub-point", 1]]
  },
  {
    "title": "Section Title",
    "title_color": "DARK",
    "items": ["Point A", "Point B"]
  }
]}
```
Bullets fill from header down to CONTENT_BOTTOM. Items can be `"string"` or `["string", indent_level]`.

### arch (architecture layers)
```json
{"type": "arch", "title": "Architecture", "layers": [
  ["Layer Name", "BLUE_LT", "Detail line 1\nDetail line 2"],
  ["Layer Name", "GREEN_LT", "Detail text"]
], "note": "Optional bottom note bar"}
```
Or dict format: `{"label": "...", "color": "BLUE_LT", "body": "..."}`.
Boxes fill vertical space with arrows auto-placed.

### flow (single-row horizontal)
```json
{"type": "flow", "title": "Process", "subtitle": "Optional", "steps": [
  ["Step 1", "BLUE"],
  ["Step 2", "GREEN"],
  ["Step 3", "ORANGE"]
]}
```
Vertically centered.

### dual_flow (stacked rows)
```json
{"type": "dual_flow", "title": "Processes", "flows": [
  {"label": "Flow A", "label_color": "BLUE", "steps": [
    ["Step 1", "BLUE_LT"], ["Step 2", "GREEN_LT"]
  ]},
  {"label": "Flow B", "label_color": "GREEN", "steps": [
    ["Step 1", "GREEN_LT"], ["Step 2", "ORANGE_LT"]
  ]}
]}
```
Rows fill vertical space.

### steps (numbered procedure)
```json
{"type": "steps", "title": "Update Process", "steps": [
  ["Step Title", "Description text"],
  ["Step Title", "Description text"]
], "notes": {"title": "注意事項", "items": ["Note 1", "Note 2"]}}
```
Steps centered vertically; notes panel fills right side.

### cards (KPI row)
```json
{"type": "cards", "title": "KPI Overview", "subtitle": "Optional", "cards": [
  ["Label", "Value", "BLUE"],
  ["Label", "Value", "GREEN"]
]}
```
Cards centered in the zone.

### table
```json
{"type": "table", "title": "Comparison",
 "headers": ["Col A", "Col B", "Col C"],
 "rows": [["1","2","3"], ["4","5","6"]],
 "header_color": "BLUE", "stripe": true}
```
Table centered vertically.

### highlight (key decisions)
```json
{"type": "highlight", "title": "Design Decisions", "points": [
  ["Heading", "Description text"],
  ["Heading", "Description text"]
]}
```
Full-width dark boxes, vertically distributed.

### sequence
```json
{"type": "sequence", "title": "Interaction",
 "actors": [["Client", "BLUE"], ["Server", "GREEN"]],
 "messages": [[0, 1, "Request"], [1, 0, "Response"]]}
```

### tree (file tree / monospace)
```json
{"type": "tree", "title": "Project Structure",
 "text": "dir/\n├── file1\n└── file2",
 "font_size": 16}
```
Centered on light gray background. `font_size` minimum enforced at 14pt.

### timeline (horizontal milestones)
```json
{"type": "timeline", "title": "Project Roadmap", "subtitle": "Optional", "milestones": [
  ["Q1 2026", "Requirement Analysis", "BLUE"],
  ["Q2 2026", "Development Phase", "GREEN"],
  ["Q3 2026", "Testing & QA", "ORANGE"],
  ["Q4 2026", "Production Launch", "PURPLE"]
]}
```
Cards alternate above/below a horizontal line. Each milestone: `[label, description, color]`.

### comparison (side-by-side)
```json
{"type": "comparison", "title": "Before vs After", "subtitle": "Optional",
 "left": {
   "header": "Before", "header_color": "RED",
   "items": ["Manual process", "Error-prone", "Slow"]
 },
 "right": {
   "header": "After", "header_color": "GREEN",
   "items": ["Automated", "Reliable", "Fast"]
 }}
```
Two columns with a **VS** divider in the center. Good for pros/cons, old/new, option A/B.

### quote (centered large text)
```json
{"type": "quote", "title": "Key Insight",
 "quote": "Data is the new oil, but only if you refine it.",
 "author": "Clive Humby",
 "accent_color": "BLUE"}
```
Large centered quote text (20pt) with decorative quotation mark and author attribution.

### matrix (2×2 grid)
```json
{"type": "matrix", "title": "Priority Matrix", "subtitle": "Optional",
 "x_labels": ["Low Effort", "High Effort"],
 "y_labels": ["High Impact", "Low Impact"],
 "cells": [
   [["Quick Wins", "Do first", "GREEN_LT"], ["Major Projects", "Plan carefully", "BLUE_LT"]],
   [["Fill-ins", "If time allows", "AMBER_LT"], ["Avoid", "Low priority", "RED_LT"]]
 ]}
```
Each cell: `[label, description, color]` or `{"label": "...", "desc": "...", "color": "..."}`.
Optional axis labels with `x_labels` (top) and `y_labels` (left).

### stat (big numbers display)
```json
{"type": "stat", "title": "Performance Metrics", "subtitle": "Optional", "stats": [
  ["99.9%", "System Uptime", "GREEN"],
  ["2,488", "Total Records", "BLUE"],
  ["< 2s", "Response Time", "ORANGE"],
  ["13", "Active Users", "PURPLE"]
]}
```
Each stat: `[value, label, color]` or `{"value": "...", "label": "...", "color": "..."}`.
Big number (44pt) with accent color bar and label. Cards vertically centered.

## Complete Slide Type List

| Type | Purpose | Key Feature |
|------|---------|-------------|
| `title` | Cover page | 40pt title, centered |
| `section` | Section divider | Same as title |
| `end` | Closing page | Same as title |
| `agenda` | Table of contents | Numbered items, centered |
| `bullet` | Text content (1~N cols) | Headers + bullet points |
| `arch` | Architecture layers | Boxes + arrows, centered |
| `flow` | Horizontal process | Single row, centered |
| `dual_flow` | Multi-row process | Stacked rows |
| `steps` | Numbered procedure | Steps + optional notes panel |
| `cards` | KPI cards row | Accent-colored cards |
| `table` | Data table | Striped, centered |
| `highlight` | Key decisions | Dark full-width boxes |
| `sequence` | Sequence diagram | Actors + messages |
| `tree` | File tree / monospace | Gray background panel |
| `timeline` | Milestones | Horizontal line, alternating cards |
| `comparison` | Side-by-side | VS divider, two columns |
| `quote` | Centered quote | Large text + author |
| `matrix` | 2×2 grid | 4 colored cells + axis labels |
| `stat` | Big numbers | Large values + accent bars |

## Checklist

### Content & Planning
- [ ] Every slide traces back to extracted source material (§1)
- [ ] Slide count matches budget for presentation length (§2)
- [ ] Each content mapped to best-fit type, not just `bullet` (§3)
- [ ] `bullet` type ≤ 40% of content slides (§4)
- [ ] No 3 consecutive same-type slides (§4)
- [ ] Visual slides (arch/flow/timeline/etc.) ≥ 25% of content slides (§4)
- [ ] Sections follow skeleton structure (§5)
- [ ] Text per slide within density limits (§6)
- [ ] Colors are consistent across slides for same topics (§7)

### Technical
- [ ] JSON is valid (no trailing commas, proper encoding)
- [ ] All slide types are from the Slide Type Reference
- [ ] Colors use string names ("BLUE") or hex ("#2563EB")
- [ ] Items with indent use array format: `["text", 1]`
- [ ] Output path is specified in JSON or CLI
