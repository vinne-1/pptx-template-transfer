# PPTX Template Transfer

Apply one deck's visual design -- logos, colors, fonts, layout patterns -- to another deck's content. Built on `python-pptx`.

**Source content is the authority. Target style is the grammar. Zero target body text leaks.**

## The Problem

Branded PPTX files store visual design (logos, decorative shapes, watermarks, backgrounds) as shapes inside slides, not in slide masters/layouts. Traditional layout/master transfer does nothing useful for these files.

## The Solution

The engine **analyzes the target template's visual DNA** (colors, fonts, logo, layout patterns) then **rebuilds each slide from scratch** using the source deck's content. Template slides are never copied -- they're only read for style extraction -- so zero template text can leak into the output.

### Pipeline

```
Source PPTX (content)     Target PPTX (style)
       |                         |
  Extract content          Analyze visual DNA
  Classify slide types     Mine layout patterns
       |                         |
       +--------> Rebuild <------+
                    |
            Output PPTX
       (source meaning, target grammar)
```

## Installation

Requires Python 3.10+.

```bash
# From the project directory
pip install -e .

# Or install dependencies directly
pip install python-pptx lxml
```

This installs the `pptx-transfer` CLI command and the `pptx_template_transfer` Python package.

## Quick Start

### CLI

```bash
# template = style source, content = content source
pptx-transfer template.pptx content.pptx output.pptx

# Or via python -m
python -m pptx_template_transfer template.pptx content.pptx output.pptx

# With verbose diagnostics
pptx-transfer template.pptx content.pptx output.pptx --verbose

# JSON report for automation
pptx-transfer template.pptx content.pptx output.pptx --report report.json

# Quality validation report
pptx-transfer template.pptx content.pptx output.pptx --quality-report quality.json

# Branding controls
pptx-transfer template.pptx content.pptx output.pptx --no-logo --no-footer
pptx-transfer template.pptx content.pptx output.pptx --footer-company "Acme Corp"

# Named arguments (alternative syntax)
pptx-transfer --target template.pptx --source content.pptx --output output.pptx

# Analyze template shape classifications
pptx-transfer --analyze template.pptx

# Extract structured content as JSON
pptx-transfer --extract content.pptx
```

### Programmatic API

```python
from pathlib import Path
from pptx_template_transfer import transfer, TransferConfig, Thresholds, BrandingPolicy

# Basic usage
report = transfer(
    template=Path("template.pptx"),   # style source
    content=Path("content.pptx"),     # content source
    output=Path("output.pptx"),
)

# Full configuration
config = TransferConfig(
    mode="recreate",
    verbose=True,
    preserve_notes=True,
    thresholds=Thresholds(title_min_font_pt=18),
    branding=BrandingPolicy(
        show_logo=True,
        show_footer=True,
        footer_company_override="My Company",
    ),
)
report = transfer(template, content, output, config)

# Report includes provenance, coverage, and quality
print(report["source_coverage"]["overall_pct"])   # e.g. 100.0
print(report["quality"]["overall_score"])          # e.g. 95.0
```

## Package Structure

```
pptx_template_transfer/
  __init__.py              # Public API re-exports
  cli.py                   # CLI entry point and transfer() orchestrator
  helpers.py               # Shared utilities (text_of, word_count, style_runs, etc.)
  models.py                # All dataclasses (config, content, provenance, quality)
  analysis/
    theme_extractor.py     # Extract visual DNA from template (fonts, colors, logo, footer)
    slide_classifier.py    # Classify shape roles and slide types
    layout_patterns.py     # Mine layout archetypes from template slides
  extraction/
    content_extractor.py   # Extract structured content from source slides
    semantic_blocks.py     # Group paragraphs into semantic blocks
  transform/
    slide_builder.py       # Recreate mode: build slides from scratch with type-specific renderers
    clone_injector.py      # Clone mode (legacy): clone template slides, inject content
    layout_mapper.py       # Map content to layout zones
    overflow_resolver.py   # Detect and resolve text overflow
  validation/
    contamination_checker.py  # Detect target body-text contamination via Jaccard n-gram similarity
    source_coverage.py        # Track per-slide source content coverage
    quality_report.py         # Per-slide quality scoring with acceptance gates
    overlap_checker.py        # Detect shape overlaps
    bounds_checker.py         # Detect shapes extending beyond slide edges
```

## How It Works

### 1. Template Analysis

Extracts the target deck's visual DNA:
- **Fonts**: Heading and body fonts from theme XML, with frequency-scan fallback. Exotic fonts (Montserrat, Lato, Poppins, etc.) are auto-resolved to safe system fonts (Calibri, Segoe UI, Arial)
- **Colors**: Primary, secondary, text, muted, background, card, line -- using saturation-based classification. Source deck colors are also extracted for future remapping
- **Logo**: Most frequently repeated image across slides
- **Footer**: Company name, confidentiality notice, page number format (requires 2+ slide repetition)
- **Layout patterns**: Column counts, zone roles, text capacities

### 2. Content Extraction

Parses each source slide into structured data:
- **Title**: Multi-signal scoring (placeholder type, font size, position, word count)
- **Body paragraphs**: Ordered text with level, bold/italic, font size, per-run hyperlinks
- **Tables**: Cell text matrices with formatting (rebuilt in output with template styling)
- **Images**: Content images above area threshold (transferred with position preservation)
- **Charts**: Chart elements for transfer
- **Speaker notes**: Full text from notes pane
- **Slide type**: title, agenda, content_narrative, metrics_dashboard, comparison, image_heavy, etc.

### 3. Slide Building

Each output slide is built from scratch using type-specific renderers:

| Slide Type | Renderer | Layout |
|------------|----------|--------|
| `title` | Title slide | Centered title + subtitle |
| `section` | Section divider | Large section name |
| `agenda` | Agenda/TOC | Bullet list of items |
| `metrics_dashboard` | KPI cards | Numbered metric cards |
| `comparison` | KPI layout | Side-by-side analysis |
| `content_narrative` | Generic content | Section label + title + body |
| `image_heavy` | Generic content | Text + image placement |
| Incident slides | Incident renderer | Summary/Actions + Details cards |

Every slide gets:
- Background color from template
- Decorative accent shapes in template colors
- Logo in top-left (toggleable via `--no-logo`)
- Section label (auto-generated with confidence scoring, body-keyword fallback)
- Footer with company name, confidentiality label, page number (toggleable via `--no-footer`)
- Images transferred with position preservation
- Tables rebuilt with template-styled formatting
- Auto-retry: slides scoring below threshold are retried with the generic renderer

### 4. Content Provenance

Every output block is tagged with its origin:

| Tag | Meaning |
|-----|---------|
| `source_content` | Body text, titles -- from source deck |
| `target_shell` | Footer, logo, decorative shapes -- from template |
| `converter_generated_bridge` | Section labels, bridge text -- generated |

**Hard rule**: Body content must come from `source_content`. Target deck body copy never appears in output.

### 5. Post-Generation Validation

After building, the pipeline automatically runs:

- **Target contamination check**: Jaccard 3-gram similarity between output and target body text. Flags if >40% similarity.
- **Source coverage report**: Per-slide word-level overlap tracking. Warns if overall coverage is weak.
- **Quality scoring**: Per-slide composite score (coverage - overlap/bounds/font penalties). Acceptance gate at 50/100.
- **Text leakage detection**: Finds non-boilerplate sentences duplicated across slides.
- **Bounds/overlap checking**: Detects shapes extending off-slide or overlapping.

## Configuration

### Thresholds

All classification thresholds are configurable:

```python
from pptx_template_transfer import Thresholds

th = Thresholds(
    title_min_font_pt=16,        # Min font size to consider as title
    body_min_area_pct=3.0,       # Min area % for body shapes
    body_max_zones=3,            # Max body text zones per slide
    decorative_max_area_pct=2.0, # Max area % for decorative shapes
    image_min_area_pct=1.5,      # Min area % for content images
    overflow_max_font_scale=0.70,# Min font scale before overflow warning
)
```

### Branding Policy

Control how branding elements appear in the output:

```python
from pptx_template_transfer import BrandingPolicy

branding = BrandingPolicy(
    show_logo=True,                         # Include template logo (--no-logo to disable)
    show_footer=True,                       # Include footer bar (--no-footer to disable)
    show_confidentiality=True,              # Include "Confidential" label in footer
    footer_company_override="My Company",   # Override auto-detected footer text
)
```

## Transfer Report

The `transfer()` function returns a comprehensive report:

```json
{
  "mode": "recreate",
  "slides": [
    {
      "index": 1,
      "source_slide": 1,
      "content_type": "title",
      "title": "Managed Detection and Response (MDR) Report",
      "word_count": 17,
      "status": "ok",
      "provenance": {
        "title": "source_content",
        "body": "source_content",
        "footer": "target_shell",
        "section_label": "converter_generated_bridge"
      }
    }
  ],
  "source_coverage": {
    "overall_pct": 100.0,
    "total_source_slides": 16,
    "total_output_slides": 16,
    "unmapped": [],
    "entries": [...]
  },
  "quality": {
    "overall_score": 100.0,
    "native_count": 16,
    "fallback_count": 0,
    "slides_needing_review": []
  },
  "warnings": [],
  "errors": []
}
```

## Testing

```bash
# Unit tests (76 tests)
python -m pytest test_pptx_template_transfer.py -v

# Regression tests (36 tests, requires sample decks)
python -m pytest test_regression.py -v

# All tests (112 total)
python -m pytest -v
```

Regression tests verify:
- Source content is the primary authority in output
- No target body-text contamination
- Source coverage >= 50%
- Content provenance is tracked correctly
- Target shell (fonts, colors, footer) is preserved
- Slide-type classification works
- Section labels are clean (no trailing punctuation/prepositions)
- Incident slides are labeled correctly
- Per-slide quality scores exist and pass acceptance gate
- Quality score >= 50/100
- No empty slides or text leakage

## Technical Notes

- Built on `python-pptx` with `lxml` for low-level XML manipulation
- Run-level font styling via `style_runs()` -- paragraph-level `.font.name` doesn't persist in XML
- Theme inheritance: `Presentation(template_path)` preserves slide masters/theme
- Aspect ratio detection: 4:3 vs 16:9 triggers reflowed layout (15% threshold)
- Density-aware font scaling based on line count
- Footer detection requires text repetition on 2+ slides to avoid false positives
- Per-slide error isolation: one failed slide doesn't abort the deck
- Font substitution: exotic fonts auto-resolved to safe system fonts via fallback table
- Image transfer preserves position using percentage-based coordinates
- Tables rebuilt with template-styled headers and alternating row colors
- Auto-retry: slides scoring below 40/100 are retried with the generic renderer
- Block-level source coverage: 40% word overlap threshold per text block

## License

MIT
