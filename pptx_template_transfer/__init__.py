"""PPTX Template Transfer — apply one deck's visual design to another's content.

This package re-exports all public names so that existing code using
``import pptx_template_transfer as mod`` or
``from pptx_template_transfer import ...`` continues to work unchanged.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# Models (dataclasses)
# ---------------------------------------------------------------------------
from pptx_template_transfer.models import (
    BoundsIssue,
    BrandingPolicy,
    ContentData,
    ImageData,
    LayoutPattern,
    LayoutZone,
    OverlapIssue,
    ParagraphData,
    QualityReport,
    RunData,
    SemanticBlock,
    ShapeInfo,
    SlideProvenance,
    SlideQuality,
    SourceCoverageEntry,
    SourceCoverageReport,
    TemplateStyle,
    TextBlock,
    Thresholds,
    TransferConfig,
)

# ---------------------------------------------------------------------------
# Helpers (public names)
# ---------------------------------------------------------------------------
from pptx_template_transfer.helpers import (
    DATE_PATTERN,
    FOOTER_PATTERNS,
    JUST_NUMBER_RE,
    NSMAP,
    PAGE_NUM_PATTERN,
    PH_BODY_SET,
    PH_FOOTER_SET,
    PH_TITLE,
    PH_TITLE_SET,
    is_allcaps_short,
    is_chart,
    is_group,
    is_picture,
    is_table,
    max_font_pt,
    placeholder_type_int,
    rgb,
    shape_area_pct,
    shape_bottom_frac,
    shape_top_frac,
    style_runs,
    text_of,
    update_rids_in_tree,
    word_count,
)

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------
from pptx_template_transfer.analysis.slide_classifier import (
    classify_all_shapes,
    classify_shape_role,
    classify_slide_type,
    classify_template_structure,
    get_slide_zones,
)
from pptx_template_transfer.analysis.theme_extractor import analyze_template
from pptx_template_transfer.analysis.layout_patterns import mine_layout_patterns

# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------
from pptx_template_transfer.extraction.content_extractor import (
    extract_content,
    extract_all_content,
)
from pptx_template_transfer.extraction.semantic_blocks import detect_semantic_blocks

# ---------------------------------------------------------------------------
# Transform
# ---------------------------------------------------------------------------
from pptx_template_transfer.transform.slide_builder import (
    build_slide,
    apply_recreate,
    _find_blank_layout,
)
from pptx_template_transfer.transform.clone_injector import (
    apply_design,
    inject_content,
    build_slide_mapping,
    _is_protected_shape,
)
from pptx_template_transfer.transform.overflow_resolver import resolve_overflow
from pptx_template_transfer.transform.layout_mapper import map_content_to_layout

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
from pptx_template_transfer.validation.overlap_checker import check_overlaps
from pptx_template_transfer.validation.bounds_checker import check_bounds
from pptx_template_transfer.validation.quality_report import generate_quality_report
from pptx_template_transfer.validation.contamination_checker import check_target_contamination
from pptx_template_transfer.validation.source_coverage import compute_source_coverage

# ---------------------------------------------------------------------------
# CLI / Public API
# ---------------------------------------------------------------------------
from pptx_template_transfer.cli import (
    detect_mode,
    transfer,
)

# ---------------------------------------------------------------------------
# Backward-compatible underscore aliases
# ---------------------------------------------------------------------------
# Tests access these via ``mod._name`` — alias public helpers back to their
# original underscore-prefixed names so the test suite keeps passing.
_word_count = word_count
_is_allcaps_short = is_allcaps_short
_rgb = rgb
_shape_area_pct = shape_area_pct
_shape_bottom_frac = shape_bottom_frac
_shape_top_frac = shape_top_frac
_style_runs = style_runs
_validate_input = None  # replaced below

# Import private names that tests reference directly
from pptx_template_transfer.cli import _validate_input  # noqa: F811

__all__ = [
    # Models
    "BoundsIssue", "BrandingPolicy", "ContentData", "ImageData",
    "LayoutPattern", "LayoutZone",
    "OverlapIssue", "ParagraphData", "QualityReport", "RunData",
    "SemanticBlock", "ShapeInfo", "SlideQuality", "TemplateStyle",
    "TextBlock", "Thresholds", "TransferConfig",
    # Helpers
    "text_of", "word_count", "max_font_pt", "shape_area_pct",
    "is_picture", "is_table", "is_chart", "is_group",
    "shape_bottom_frac", "shape_top_frac", "is_allcaps_short",
    "placeholder_type_int", "rgb", "style_runs", "update_rids_in_tree",
    # Analysis
    "classify_all_shapes", "classify_shape_role", "classify_slide_type",
    "classify_template_structure", "get_slide_zones",
    "analyze_template", "mine_layout_patterns",
    # Extraction
    "extract_content", "extract_all_content", "detect_semantic_blocks",
    # Transform
    "build_slide", "apply_recreate", "apply_design",
    "inject_content", "build_slide_mapping",
    "resolve_overflow", "map_content_to_layout",
    # Validation
    "check_overlaps", "check_bounds", "generate_quality_report",
    # CLI / API
    "detect_mode", "transfer",
]
