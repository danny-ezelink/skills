# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

This is a **skills library for Claude Code** — a collection of specialized, self-contained skills that Claude invokes to perform design-heavy tasks. There are three skills:

- **`skills/ui-ux-pro-max/`** — Data-driven UI/UX guidance with BM25 search over design datasets
- **`skills/frontend-design/`** — Methodology and constraints for producing non-generic frontend code
- **`skills/pptx/`** — Full pipeline for reading, editing, and creating PowerPoint presentations

## Common Commands

### PPTX Skill

```bash
# Extract text from a presentation
python -m markitdown presentation.pptx

# Generate a visual thumbnail grid of all slides
python skills/pptx/scripts/thumbnail.py presentation.pptx

# Unpack a .pptx for XML-level editing
python skills/pptx/scripts/office/unpack.py input.pptx unpacked/

# Duplicate a slide or add one from a layout
python skills/pptx/scripts/add_slide.py unpacked/ <slide_index>

# Remove orphaned files after slide deletions
python skills/pptx/scripts/clean.py unpacked/

# Repack into a .pptx (preserves original's media/theme when --original is given)
python skills/pptx/scripts/office/pack.py unpacked/ output.pptx --original template.pptx

# Convert to PDF (for visual QA)
python skills/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide        # → slide-01.jpg, slide-02.jpg, ...
```

### UI/UX Pro Max Skill

All commands run from `skills/ui-ux-pro-max/scripts/`:

```bash
# Search a domain (style, color, chart, landing, product, ux, typography, google-fonts, icons, react, web)
python search.py "dark dashboard SaaS" --domain style

# Search for stack-specific guidelines (react, nextjs, vue, svelte, astro, swiftui, react-native, flutter, nuxtjs, nuxt-ui, html-tailwind, shadcn, jetpack-compose, threejs)
python search.py "form validation" --stack nextjs

# Generate a full design system from a query
python search.py "fintech mobile app" --design-system -p "MyApp"

# Generate and persist to MASTER.md + page-level override
python search.py "dashboard" --design-system --persist -p "MyApp" --page "dashboard"
```

### Dependencies

```bash
pip install "markitdown[pptx]"   # PPTX text extraction
pip install Pillow               # Thumbnail generation
npm install                      # pptxgenjs (for creating from scratch)
# System: LibreOffice (soffice), Poppler (pdftoppm)
```

## Architecture

### PPTX Skill

The skill treats `.pptx` files as ZIP archives containing XML. The editing workflow is:
**unpack → edit XML directly → clean → repack**. Slide content lives in `unpacked/ppt/slides/slide{N}.xml`. Structural changes (adding/removing/reordering slides) must be completed in `unpacked/ppt/presentation.xml` (`<p:sldIdLst>`) before editing content. `pack.py` validates the result and can pull missing media/relationships from the original file.

For creating from scratch, `pptxgenjs` (a Node.js library) is used — see `skills/pptx/pptxgenjs.md` for the full API reference.

**QA is mandatory**: after generation, always run `markitdown` for content QA and convert to images for visual QA. Use subagents for visual inspection — fresh eyes catch issues the generating agent misses. The verification loop (generate → inspect → fix → re-verify) must complete at least one cycle before declaring success.

### UI/UX Pro Max Skill

The search engine (`scripts/search.py` + `scripts/core.py`) runs BM25 full-text search in-memory over CSV data files in `data/`. No database or external API is used. Results return up to 3 ranked rows from the matching domain CSV.

The **Master + Overrides pattern** for design system persistence: `--persist` writes `design-system/MASTER.md` as the global token reference; `--page <name>` writes `design-system/pages/<name>.md` for page-specific overrides that complement the master.

Available domains and their backing CSV files:
- `style` → `styles.csv` (25+ design styles with implementation checklists)
- `color` → `colors.csv` (161 product-type color palettes with semantic tokens)
- `product` → `products.csv` (161 product types with style recommendations)
- `typography` → `typography.csv` (57 font pairings)
- `ux` → `ux-guidelines.csv` (99 guidelines with severity, do/don't, code examples)
- `chart` → `charts.csv` (25 chart types with accessibility grades)
- `landing` → `landing.csv` (landing page conversion patterns)
- `icons`, `react`, `web`, `google-fonts` → domain-specific CSVs

### Frontend Design Skill

Pure methodology — no executable code. The `SKILL.md` constrains aesthetic choices to prevent generic AI outputs. Key rules: commit to a bold, specific aesthetic direction before writing code; never use Inter/Roboto/Arial/system fonts; never use purple-gradient-on-white; use CSS variables for color tokens; favor CSS-only animations over JS where possible.

## Skill Invocation

Each skill's `SKILL.md` frontmatter contains a `description` field that determines when Claude should invoke it. The trigger patterns are:

- **pptx**: Any mention of `.pptx`, "deck", "slides", or "presentation" — even if the goal is just reading/extracting content
- **ui-ux-pro-max**: Any task that changes how a UI looks, feels, moves, or is interacted with
- **frontend-design**: Building web components, pages, or applications where design quality matters
