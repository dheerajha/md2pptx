# md2pptx — Markdown to PPTX Converter

A modular pipeline that converts Markdown documents into structured, visually professional PowerPoint presentations. Built for the MD→PPTX Hackathon 2026.

---

## Quick Start

```bash
npm install
node index.js path/to/input.md output/presentation.pptx
```

---

## Setup Instructions

### Prerequisites

- **Node.js** 18 or higher
- **npm** 9 or higher

### Install

```bash
git clone https://github.com/your-username/md2pptx
cd md2pptx
npm install
```

### Run

```bash
# Basic usage
node index.js input.md

# With explicit output path
node index.js input.md output/my-deck.pptx

# Test with included samples
node index.js test-cases/ai-in-business.md output/ai-in-business.pptx
node index.js test-cases/sustainability-report.md output/sustainability-report.pptx
```

---

## System Architecture

```
md2pptx/
├── index.js              # Entry point & orchestrator
├── src/
│   ├── parser.js         # Markdown → structured AST
│   ├── structurer.js     # AST → slide story (10-15 slides)
│   ├── renderer.js       # Story → PptxGenJS calls
│   └── themes.js         # Color palettes & typography
├── test-cases/           # Sample markdown inputs
└── output/               # Generated PPTX files
```

### Pipeline Stages

```
Input .md
   │
   ▼
[1] PARSER         Tokenises raw markdown into sections, blocks
    parser.js      (paragraphs, lists, tables, code, quotes, headings)
   │
   ▼
[2] STRUCTURER     Builds a 10-15 slide "story" from parsed content
    structurer.js  Decides layout types, detects data, picks theme
   │
   ▼
[3] RENDERER       Translates story into PptxGenJS API calls
    renderer.js    Draws shapes, text, charts, tables per slide type
   │
   ▼
Output .pptx
```

---

## Key Design Decisions

### 1. Separation of Concerns

Each module has a single responsibility:
- `parser.js` — tokenises Markdown only, knows nothing about slides
- `structurer.js` — makes story/layout decisions, knows nothing about drawing
- `renderer.js` — draws slides, knows nothing about content logic
- `themes.js` — stores design tokens only

### 2. Content-Driven Layout Selection

The structurer analyses each section and picks the most appropriate layout:

| Content type | Layout chosen |
|---|---|
| 4+ list items, minimal prose | `icon_cards` — colourful numbered cards |
| 2+ subheadings with descriptions | `two_column_cards` — side-by-side cards |
| Blockquote present | `quote_callout` — bold quote on dark background |
| Section has a markdown table | `chart_table` — bar/line chart + data table |
| Standard prose + bullets | `content` — body text with optional stat callout |

### 3. Automatic Theme Selection

The document title is matched against keyword patterns to choose an appropriate colour palette:

```
tech/ai/data/digital  → Midnight (dark navy + cyan)
finance/revenue       → Ocean (deep blue)
health/medical        → Teal Trust
sustain/green/climate → Forest & Moss
brand/marketing       → Coral Energy
(default)             → Charcoal Minimal
```

### 4. Smart Chart Column Detection

Before building a chart, the renderer finds the best value column — skipping year-like columns (4-digit integers) in position 0 — so charts plot meaningful metrics rather than year numbers on the Y axis.

### 5. Structured Flow Enforcement

Slide order always follows:
```
Title → Agenda (if ≥4 sections) → Executive Summary → Content Slides
     → Data/Chart Slides → Conclusion
```

The `trimSlides()` function enforces the 10-15 slide cap while preserving this order.

### 6. Clean Text Truncation

All text that exceeds display bounds is cut at the nearest sentence boundary (`. `) when possible, or at the nearest word boundary otherwise, then appended with `…`. No mid-word cuts.

---

## Slide Types

| Type | Description |
|---|---|
| `title` | Dark full-bleed with decorative circles and accent bar |
| `agenda` | Two-column numbered list of sections |
| `executive_summary` | Dark bg, accent-bordered card per key point |
| `content` | Light bg, body text with optional stat callout box |
| `icon_cards` | Grid of coloured cards, each with number badge |
| `two_column_cards` | 2×2 cards with accent left border and heading |
| `quote_callout` | Dark bg, oversized quotation mark, italic quote |
| `chart_table` | Bar/line chart on left, data table on right |
| `data_table` | Full-width formatted table |
| `conclusion` | Dark bg with numbered circle takeaway list |

---

## Supported Markdown Features

- `# H1` → Document title
- `## H2` → Major section (new slide)
- `### H3` → Subheading within a section
- `- / * / 1.` → Bullet/numbered lists (with sub-items)
- `| table |` → Auto-detected as chart + table slide
- `> blockquote` → Quote callout slide
- ` ```code``` ` → Code block (preserved)
- `**bold**` / `*italic*` → Stripped cleanly from slide text
- Inline numbers → Auto-detected as stat callouts

---

## Constraints

| Constraint | Value |
|---|---|
| Max input file size | 5 MB |
| Slide count | 10-15 |
| Max bullet items per slide | 6 |
| Max icon cards per slide | 6 |
| Max table rows rendered | 8 |
| Max data slides | 3 |

---

## Output Compatibility

Generated `.pptx` files open correctly in:
- Microsoft PowerPoint (365 / 2019 / 2016)
- Google Slides (via import)
- LibreOffice Impress

---

## Edge Case Handling

| Edge case | Behaviour |
|---|---|
| No `# H1` title | Uses first heading found, or "Presentation" |
| No subtitle line | Extracted from first meaningful paragraph |
| Table with all-text columns | Rendered as `data_table` (no chart) |
| Year column in position 0 | Excluded from chart Y axis; used as X labels |
| Very short markdown (<3 sections) | Agenda skipped; exec summary uses available content |
| Long text exceeding slide area | Truncated at sentence/word boundary with `…` |
| `$`, `%`, `K`, `M`, `B` suffixes | Parsed as numeric values for chart rendering |

---

## Dependencies

| Package | Purpose |
|---|---|
| `pptxgenjs` | PowerPoint file generation |

No external APIs. No internet required at runtime.
