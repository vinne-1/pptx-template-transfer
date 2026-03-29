"""Analysis modules for inspecting template and source decks."""
from pptx_template_transfer.analysis.slide_classifier import (
    classify_all_shapes,
    classify_shape_role,
    get_slide_zones,
)
from pptx_template_transfer.analysis.theme_extractor import analyze_template
from pptx_template_transfer.analysis.layout_patterns import mine_layout_patterns
