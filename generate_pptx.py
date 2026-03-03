# -*- coding: utf-8 -*-
"""
generate_pptx.py — Generic PPTX generator.

Reads a JSON content file and produces a PPTX presentation
using pptx_engine. This script is NOT topic-specific.

Usage:
    python generate_pptx.py <content.json> [output.pptx]

The JSON structure:
{
  "template": "path/to/template.pptx",   // optional, uses default
  "output": "path/to/output.pptx",       // optional if given as CLI arg
  "slides": [
    {"type": "title", "title": "...", "subtitle": "..."},
    {"type": "agenda", "title": "...", "items": [["1","Topic","Desc"], ...]},
    {"type": "bullet", "title": "...", "subtitle": "...", "columns": [...]},
    {"type": "arch", "title": "...", "layers": [...], "note": "..."},
    {"type": "flow", "title": "...", "steps": [...], "subtitle": "..."},
    {"type": "dual_flow", "title": "...", "flows": [...]},
    {"type": "steps", "title": "...", "steps": [...], "notes": {...}},
    {"type": "cards", "title": "...", "cards": [...], "subtitle": "..."},
    {"type": "table", "title": "...", "headers": [...], "rows": [...]},
    {"type": "highlight", "title": "...", "points": [...]},
    {"type": "sequence", "title": "...", "actors": [...], "messages": [...]},
    {"type": "tree", "title": "...", "text": "..."},
    {"type": "section", "title": "...", "subtitle": "..."},
    {"type": "end", "title": "...", "subtitle": "..."}
  ]
}
"""
import sys
import json
import os

# Add engine to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_engine import (
    PptxEngine, bh,
    BLUE, BLUE_LT, BLUE_MID,
    GREEN, GREEN_LT,
    ORANGE, ORANGE_LT,
    RED, RED_LT,
    PURPLE, PURPLE_LT,
    AMBER, AMBER_LT,
    CYAN, CYAN_LT,
    DARK, GRAY, GRAY_LT, GRAY_BG, WHITE, BLACK,
    CONTENT_TOP, CONTENT_BOTTOM, CONTENT_H, CONTENT_W, MARGIN_X,
    SZ_TITLE, SZ_SECTION, SZ_BODY, SZ_SUB, SZ_SMALL, SZ_TINY,
)

# ── Color name → RGBColor mapping ──────────────────────────
COLOR_MAP = {
    "BLUE": BLUE, "BLUE_LT": BLUE_LT, "BLUE_MID": BLUE_MID,
    "GREEN": GREEN, "GREEN_LT": GREEN_LT,
    "ORANGE": ORANGE, "ORANGE_LT": ORANGE_LT,
    "RED": RED, "RED_LT": RED_LT,
    "PURPLE": PURPLE, "PURPLE_LT": PURPLE_LT,
    "AMBER": AMBER, "AMBER_LT": AMBER_LT,
    "CYAN": CYAN, "CYAN_LT": CYAN_LT,
    "DARK": DARK, "GRAY": GRAY, "GRAY_LT": GRAY_LT, "GRAY_BG": GRAY_BG,
    "WHITE": WHITE, "BLACK": BLACK,
}


def resolve_color(val):
    """Convert a color string name or hex to RGBColor."""
    if val is None:
        return None
    if isinstance(val, str):
        # Try named color first
        upper = val.upper().replace(" ", "_")
        if upper in COLOR_MAP:
            return COLOR_MAP[upper]
        # Try hex: "#2563EB" or "2563EB"
        h = val.lstrip("#")
        if len(h) == 6:
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    return val  # already RGBColor or passthrough


def resolve_items(items):
    """Convert JSON items to engine format.
    Items can be strings or [text, level] arrays.
    """
    result = []
    for item in items:
        if isinstance(item, list) and len(item) == 2:
            result.append((item[0], item[1]))
        else:
            result.append(item)
    return result


def resolve_columns(columns):
    """Convert JSON column dicts to engine format."""
    result = []
    for col in columns:
        resolved = {}
        if "header" in col:
            resolved["header"] = col["header"]
        if "header_color" in col:
            resolved["header_color"] = resolve_color(col["header_color"])
        if "title" in col:
            resolved["title"] = col["title"]
        if "title_color" in col:
            resolved["title_color"] = resolve_color(col["title_color"])
        if "items" in col:
            resolved["items"] = resolve_items(col["items"])
        result.append(resolved)
    return result


def resolve_layers(layers):
    """Convert JSON layers to (label, color, body) tuples."""
    result = []
    for layer in layers:
        if isinstance(layer, list):
            label = layer[0]
            color = resolve_color(layer[1])
            body = layer[2] if len(layer) > 2 else ""
            result.append((label, color, body))
        elif isinstance(layer, dict):
            result.append((
                layer["label"],
                resolve_color(layer.get("color", "BLUE_LT")),
                layer.get("body", ""),
            ))
    return result


def resolve_flow_steps(steps):
    """Convert JSON flow steps to (label, color) tuples."""
    result = []
    for step in steps:
        if isinstance(step, list):
            result.append((step[0], resolve_color(step[1]) if len(step) > 1 else BLUE))
        elif isinstance(step, dict):
            result.append((step["label"], resolve_color(step.get("color", "BLUE"))))
    return result


def resolve_actors(actors):
    """Convert JSON actors to (name, color) tuples."""
    result = []
    for a in actors:
        if isinstance(a, list):
            result.append((a[0], resolve_color(a[1]) if len(a) > 1 else BLUE))
        elif isinstance(a, dict):
            result.append((a["name"], resolve_color(a.get("color", "BLUE"))))
    return result


def resolve_messages(messages):
    """Convert JSON messages to (from_idx, to_idx, label) tuples."""
    result = []
    for m in messages:
        if isinstance(m, list):
            result.append((m[0], m[1], m[2] if len(m) > 2 else ""))
        elif isinstance(m, dict):
            result.append((m["from"], m["to"], m.get("label", "")))
    return result


def resolve_cards(cards):
    """Convert JSON cards to (label, value, color) tuples."""
    result = []
    for c in cards:
        if isinstance(c, list):
            label = c[0]
            value = c[1]
            color = resolve_color(c[2]) if len(c) > 2 else BLUE
            result.append((label, value, color))
        elif isinstance(c, dict):
            result.append((
                c["label"], c["value"],
                resolve_color(c.get("color", "BLUE")),
            ))
    return result


def resolve_points(points):
    """Convert JSON points to (heading, description) tuples."""
    result = []
    for p in points:
        if isinstance(p, list):
            result.append((p[0], p[1] if len(p) > 1 else ""))
        elif isinstance(p, dict):
            result.append((p["heading"], p.get("description", "")))
    return result


def resolve_step_items(steps):
    """Convert JSON step items to (title, description) tuples."""
    result = []
    for s in steps:
        if isinstance(s, list):
            result.append((s[0], s[1] if len(s) > 1 else ""))
        elif isinstance(s, dict):
            result.append((s["title"], s.get("description", "")))
    return result


def resolve_notes(notes):
    """Convert JSON notes to engine format."""
    if notes is None:
        return None
    if isinstance(notes, dict):
        return {
            "title": notes.get("title", "注意事項"),
            "items": notes.get("items", []),
        }
    if isinstance(notes, list):
        return notes
    return None


def resolve_flows(flows):
    """Convert JSON flows for dual_flow_slide."""
    result = []
    for f in flows:
        result.append({
            "label": f["label"],
            "label_color": resolve_color(f.get("label_color", "BLUE")),
            "steps": resolve_flow_steps(f["steps"]),
        })
    return result


# ── Main generator ──────────────────────────────────────────

def generate(content, output_path=None):
    """Generate a PPTX from a content dict."""
    template = content.get("template")  # path, "" for blank, or omit for default
    footer = content.get("footer_text")  # string or omit
    eng = PptxEngine(template=template, footer_text=footer)

    # Auto-decoration: on by default, can be disabled in JSON
    eng.auto_decorate = content.get("auto_decorate", True)

    # Track slides before/after each builder to get the newly added slide
    skip_decorate = {"title", "section", "end", "agenda"}

    for slide_def in content.get("slides", []):
        stype = slide_def["type"]
        title = slide_def.get("title", "")
        n_before = len(eng.prs.slides)

        if stype == "title":
            kwargs = {}
            if "title_color" in slide_def:
                kwargs["title_color"] = resolve_color(slide_def["title_color"])
            if "sub_color" in slide_def:
                kwargs["sub_color"] = resolve_color(slide_def["sub_color"])
            eng.title_slide(title, slide_def.get("subtitle", ""), **kwargs)

        elif stype == "section":
            kwargs = {}
            if "title_color" in slide_def:
                kwargs["title_color"] = resolve_color(slide_def["title_color"])
            if "sub_color" in slide_def:
                kwargs["sub_color"] = resolve_color(slide_def["sub_color"])
            eng.section_slide(title, slide_def.get("subtitle", ""), **kwargs)

        elif stype == "end":
            kwargs = {}
            if "title_color" in slide_def:
                kwargs["title_color"] = resolve_color(slide_def["title_color"])
            if "sub_color" in slide_def:
                kwargs["sub_color"] = resolve_color(slide_def["sub_color"])
            eng.end_slide(
                slide_def.get("title", "謝謝各位"),
                slide_def.get("subtitle", "Q&A 時間"),
                **kwargs,
            )

        elif stype == "agenda":
            items = [tuple(i) for i in slide_def["items"]]
            eng.agenda_slide(title, items)

        elif stype == "bullet":
            eng.bullet_slide(
                title,
                resolve_columns(slide_def["columns"]),
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "arch":
            eng.arch_slide(
                title,
                resolve_layers(slide_def["layers"]),
                note=slide_def.get("note"),
            )

        elif stype == "flow":
            eng.flow_slide(
                title,
                resolve_flow_steps(slide_def["steps"]),
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "dual_flow":
            eng.dual_flow_slide(title, resolve_flows(slide_def["flows"]))

        elif stype == "steps":
            eng.steps_slide(
                title,
                resolve_step_items(slide_def["steps"]),
                notes=resolve_notes(slide_def.get("notes")),
                notes_title=slide_def.get("notes_title", "注意事項"),
            )

        elif stype == "cards":
            eng.cards_slide(
                title,
                resolve_cards(slide_def["cards"]),
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "table":
            eng.table_slide(
                title,
                headers=slide_def["headers"],
                rows=slide_def["rows"],
                header_color=resolve_color(slide_def.get("header_color", "BLUE")),
                stripe=slide_def.get("stripe", True),
            )

        elif stype == "highlight":
            eng.highlight_slide(title, resolve_points(slide_def["points"]))

        elif stype == "sequence":
            eng.sequence_slide(
                title,
                actors=resolve_actors(slide_def["actors"]),
                messages=resolve_messages(slide_def["messages"]),
            )

        elif stype == "tree":
            eng.tree_slide(
                title,
                slide_def["text"],
                font_size=slide_def.get("font_size", 16),
            )

        elif stype == "timeline":
            milestones = []
            for m in slide_def.get("milestones", []):
                if isinstance(m, list):
                    label = m[0]
                    desc = m[1] if len(m) > 1 else ""
                    color = resolve_color(m[2]) if len(m) > 2 else BLUE
                    milestones.append((label, desc, color))
                elif isinstance(m, dict):
                    milestones.append((
                        m.get("label", ""),
                        m.get("desc", ""),
                        resolve_color(m.get("color", "BLUE")),
                    ))
            eng.timeline_slide(
                title, milestones,
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "comparison":
            left = slide_def.get("left", {})
            right = slide_def.get("right", {})
            for side in [left, right]:
                if "header_color" in side:
                    side["header_color"] = resolve_color(side["header_color"])
                if "items" in side:
                    side["items"] = resolve_items(side["items"])
            eng.comparison_slide(
                title, left, right,
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "quote":
            eng.quote_slide(
                title,
                slide_def.get("quote", ""),
                author=slide_def.get("author", ""),
                accent_color=resolve_color(slide_def.get("accent_color", "BLUE")),
            )

        elif stype == "matrix":
            cells = slide_def.get("cells", [[], []])
            eng.matrix_slide(
                title, cells,
                x_labels=slide_def.get("x_labels"),
                y_labels=slide_def.get("y_labels"),
                subtitle=slide_def.get("subtitle"),
            )

        elif stype == "stat":
            eng.stat_slide(
                title,
                slide_def.get("stats", []),
                subtitle=slide_def.get("subtitle"),
            )

        else:
            print(f"Warning: unknown slide type '{stype}', skipping.")

        # Auto-decorate: add motif to sparse content slides
        if stype not in skip_decorate and len(eng.prs.slides) > n_before:
            new_slide = eng.prs.slides[len(eng.prs.slides) - 1]
            accent = resolve_color(slide_def.get("accent_color"))
            eng.decorate(new_slide, stype, accent)

    # Save
    out = output_path or content.get("output", "output.pptx")
    eng.save(out)
    return out


# ── CLI entry point ─────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_pptx.py <content.json> [output.pptx]")
        sys.exit(1)

    json_path = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else None

    with open(json_path, "r", encoding="utf-8") as f:
        content = json.load(f)

    generate(content, output)
