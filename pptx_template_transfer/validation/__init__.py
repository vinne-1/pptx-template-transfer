"""Validation and quality checking for output presentations."""
from pptx_template_transfer.validation.overlap_checker import check_overlaps
from pptx_template_transfer.validation.bounds_checker import check_bounds
from pptx_template_transfer.validation.quality_report import (
    generate_quality_report,
    _detect_text_leakage,
    _detect_body_in_forbidden_zones,
)
from pptx_template_transfer.validation.contamination_checker import (
    check_target_contamination,
)
from pptx_template_transfer.validation.source_coverage import (
    compute_source_coverage,
)
