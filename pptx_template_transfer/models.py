"""Data models for PPTX Template Transfer."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# ============================================================================
# CONFIGURATION
# ============================================================================

@dataclass(frozen=True)
class Thresholds:
    """All classification thresholds in one place — tune per template style."""
    # Shape role classification
    title_min_font_pt: float = 20.0
    title_max_words: int = 20
    body_min_area_pct: float = 4.0
    body_min_area_pct_relaxed: float = 3.0
    body_min_words: int = 10
    body_min_words_relaxed: int = 5
    body_max_zones: int = 2
    decorative_max_area_pct: float = 2.0
    decorative_max_words: int = 5
    decorative_max_font_pt: float = 10.0
    footer_bottom_frac: float = 0.90
    footer_top_frac: float = 0.08
    footer_max_area_pct: float = 5.0
    info_left_frac: float = 0.55
    info_min_words: int = 5
    info_max_words: int = 50
    # Content extraction
    image_min_area_pct: float = 1.5
    subheading_min_font_pt: float = 18.0
    # Matching
    variety_max_pct: float = 0.40
    # Overflow
    overflow_max_font_scale: float = 0.70
    overflow_chars_per_sq_inch: float = 180.0


@dataclass
class BrandingPolicy:
    """Controls how source and target branding are handled in the output.

    All boolean flags default to True for "target" mode.  Set to False
    to suppress individual branding elements deterministically.
    """
    # "target" = use target footer/logo/conf, "source" = keep source, "hybrid" = target shell + source attribution
    mode: str = "target"
    # Override footer company text (None = auto-detect from target)
    footer_company_override: str | None = None
    # Override confidentiality label (None = "Confidential")
    confidentiality_label: str | None = None
    # Whether to add "Presented by" attribution on title slide
    add_source_attribution: bool = False
    # Source company name for attribution (auto-detected if None)
    source_company: str | None = None
    # Deterministic element toggles
    show_logo: bool = True
    show_footer: bool = True
    show_confidentiality: bool = True


@dataclass
class TransferConfig:
    mode: str | None = None
    verbose: bool = False
    slide_map: dict[str, int] | None = None
    preserve_notes: bool = True
    auto_split: bool = False
    thresholds: Thresholds = field(default_factory=Thresholds)
    report_path: Path | None = None
    branding: BrandingPolicy = field(default_factory=BrandingPolicy)


# ============================================================================
# SHAPE CLASSIFICATION
# ============================================================================

@dataclass
class ShapeInfo:
    """Pre-computed properties of a shape for classification."""
    shape: Any
    text: str
    word_count: int
    font_size: float
    area_pct: float
    top_frac: float
    bottom_frac: float
    left_frac: float
    is_picture: bool
    is_table: bool
    is_chart: bool
    is_group: bool
    is_ole: bool
    placeholder_type: int | None
    name_lower: str
    group_text_words: int = 0


# ============================================================================
# CONTENT EXTRACTION
# ============================================================================

@dataclass
class RunData:
    text: str
    bold: bool = False
    italic: bool = False
    font_size: float = 0.0
    hyperlink_url: str | None = None


@dataclass
class ParagraphData:
    text: str
    level: int = 0
    bold: bool = False
    italic: bool = False
    font_size: float = 0.0
    runs: list[RunData] = field(default_factory=list)


@dataclass
class TextBlock:
    """A positioned text group preserving spatial layout from the source slide."""
    paragraphs: list[ParagraphData] = field(default_factory=list)
    left_pct: float = 0.0
    top_pct: float = 0.0
    width_pct: float = 20.0
    height_pct: float = 10.0
    is_heading: bool = False
    is_label: bool = False


@dataclass
class ImageData:
    """A source image with position and dimensions (EMU units)."""
    blob: bytes = b""
    width: int = 0
    height: int = 0
    left: int = 0
    top: int = 0
    content_type: str = ""


@dataclass
class ContentData:
    title: str = ""
    body_paragraphs: list[ParagraphData] = field(default_factory=list)
    text_blocks: list[TextBlock] = field(default_factory=list)
    tables: list[dict] = field(default_factory=list)
    images: list[ImageData] = field(default_factory=list)
    charts: list[Any] = field(default_factory=list)
    has_chart: bool = False
    slide_type: str = "content"
    word_count: int = 0
    primary_color: str | None = None
    notes: str = ""
    semantic_blocks: list[Any] = field(default_factory=list)
    source_slide_index: int = -1  # index in the source deck


# ============================================================================
# PROVENANCE & COVERAGE
# ============================================================================

@dataclass
class SlideProvenance:
    """Tracks the origin of every content block on an output slide."""
    output_slide_index: int
    source_slide_index: int = -1
    title_origin: str = "source_content"        # source_content | target_shell | converter_generated_bridge
    body_origin: str = "source_content"
    footer_origin: str = "target_shell"
    section_label_origin: str = "converter_generated_bridge"
    tables_from_source: int = 0
    images_from_source: int = 0
    charts_from_source: int = 0


@dataclass
class SourceCoverageEntry:
    """Coverage tracking for a single source slide."""
    source_slide_index: int
    source_title: str = ""
    source_word_count: int = 0
    output_slide_indices: list[int] = field(default_factory=list)
    text_used_pct: float = 0.0
    tables_dropped: int = 0
    charts_dropped: int = 0
    images_dropped: int = 0
    rebuild_method: str = "native"  # native | fallback | split | summarized
    blocks_total: int = 0
    blocks_covered: int = 0
    missing_block_texts: list[str] = field(default_factory=list)


@dataclass
class SourceCoverageReport:
    """Aggregate source coverage across all slides."""
    entries: list[SourceCoverageEntry] = field(default_factory=list)
    overall_text_coverage_pct: float = 0.0
    total_source_slides: int = 0
    total_output_slides: int = 0
    unmapped_source_slides: list[int] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ============================================================================
# TEMPLATE ANALYSIS
# ============================================================================

@dataclass
class TemplateStyle:
    """Visual DNA extracted from a template PPTX."""
    slide_width: int = 0
    slide_height: int = 0
    heading_font: str = "Montserrat"
    body_font: str = "Lato"
    color_primary: str = "2563EB"
    color_secondary: str = "F97316"
    color_text: str = "111827"
    color_muted: str = "475569"
    color_background: str = "F7F8FB"
    color_card: str = "FFFFFF"
    color_line: str = "D1D5DB"
    logo_blob: bytes | None = None
    logo_content_type: str = "image/png"
    logo_width: int = 0
    logo_height: int = 0
    footer_company: str = ""
    footer_has_confidential: bool = True
    footer_has_page_number: bool = True
    patterns: list[Any] = field(default_factory=list)


# ============================================================================
# LAYOUT PATTERNS (Phase 2)
# ============================================================================

@dataclass
class LayoutZone:
    """A bounding-box region with a role on a slide."""
    role: str  # "title", "body", "image", "card", "sidebar", "accent"
    left_pct: float = 0.0
    top_pct: float = 0.0
    width_pct: float = 100.0
    height_pct: float = 100.0
    text_capacity: int = 0  # estimated chars that fit


@dataclass
class LayoutPattern:
    """A detected slide archetype from the target deck."""
    name: str  # "full-narrative", "2-col", "3-card-grid", "sidebar+main", etc.
    zones: list[LayoutZone] = field(default_factory=list)
    column_count: int = 1
    total_text_capacity: int = 0
    source_slide_indices: list[int] = field(default_factory=list)
    has_image_zone: bool = False
    has_table_zone: bool = False


# ============================================================================
# SEMANTIC BLOCKS (Phase 4)
# ============================================================================

@dataclass
class SemanticBlock:
    """A semantically-grouped chunk of content."""
    block_type: str  # "numbered_list", "key_value", "metric_group", "section_header", "callout", "plain"
    paragraphs: list[ParagraphData] = field(default_factory=list)
    label: str = ""


# ============================================================================
# VALIDATION (Phase 3)
# ============================================================================

@dataclass
class OverlapIssue:
    slide_index: int
    shape_a: str
    shape_b: str
    overlap_pct: float
    severity: str  # "minor", "major", "complete"


@dataclass
class BoundsIssue:
    slide_index: int
    shape_name: str
    edges: list[str]  # ["right", "bottom"] etc.
    overflow_px: int


@dataclass
class SlideQuality:
    slide_index: int
    build_method: str  # "native", "fallback", "split"
    content_coverage_pct: float = 100.0
    font_warnings: list[str] = field(default_factory=list)
    overlap_issues: list[OverlapIssue] = field(default_factory=list)
    bounds_issues: list[BoundsIssue] = field(default_factory=list)
    needs_manual_review: bool = False
    review_reasons: list[str] = field(default_factory=list)


@dataclass
class QualityReport:
    slides: list[SlideQuality] = field(default_factory=list)
    overall_score: float = 100.0
    native_count: int = 0
    fallback_count: int = 0
    split_count: int = 0
    warnings: list[str] = field(default_factory=list)
