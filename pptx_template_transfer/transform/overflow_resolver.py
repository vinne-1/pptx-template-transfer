"""Overflow resolution — font shrinking, spacing reduction, content splitting."""
from __future__ import annotations

from dataclasses import dataclass, field

from pptx_template_transfer.models import ParagraphData


@dataclass
class OverflowResult:
    strategy: str = "none"  # "none", "font_shrink", "spacing_reduction", "split"
    font_scale: float = 1.0
    kept_paragraphs: list[ParagraphData] = field(default_factory=list)
    overflow_paragraphs: list[ParagraphData] = field(default_factory=list)


def _estimate_zone_capacity(width_pct: float, height_pct: float,
                            slide_w: int, slide_h: int,
                            font_pt: float = 12.0) -> int:
    """Estimate characters that fit in a zone at a given font size."""
    w_inches = (slide_w * width_pct / 100) / 914400.0
    h_inches = (slide_h * height_pct / 100) / 914400.0
    if w_inches <= 0 or h_inches <= 0:
        return 0
    # Characters per line depends on font size
    chars_per_inch = max(1, 72 / font_pt)  # Approximate
    chars_per_line = max(1, int(w_inches * chars_per_inch))
    # Lines per inch depends on font size + spacing
    line_height_inches = (font_pt * 1.3) / 72  # 1.3 line spacing
    lines = max(1, int(h_inches / line_height_inches))
    return chars_per_line * lines


def _total_chars(paragraphs: list[ParagraphData]) -> int:
    return sum(len(p.text) for p in paragraphs)


def resolve_overflow(
    paragraphs: list[ParagraphData],
    width_pct: float,
    height_pct: float,
    slide_w: int,
    slide_h: int,
    base_font_pt: float = 12.0,
    min_scale: float = 0.70,
) -> OverflowResult:
    """Determine how to fit paragraphs into a zone.

    Returns an OverflowResult with the strategy used.
    """
    if not paragraphs:
        return OverflowResult(kept_paragraphs=[])

    total = _total_chars(paragraphs)
    if total == 0:
        return OverflowResult(kept_paragraphs=list(paragraphs))

    # Check at base font size
    capacity = _estimate_zone_capacity(width_pct, height_pct, slide_w, slide_h, base_font_pt)
    if total <= capacity:
        return OverflowResult(kept_paragraphs=list(paragraphs))

    # Try font shrinking in steps
    for scale_pct in [95, 90, 85, 80, 75, 70]:
        scale = scale_pct / 100.0
        if scale < min_scale:
            break
        font_pt = base_font_pt * scale
        capacity = _estimate_zone_capacity(width_pct, height_pct, slide_w, slide_h, font_pt)
        if total <= capacity:
            return OverflowResult(
                strategy="font_shrink",
                font_scale=scale,
                kept_paragraphs=list(paragraphs),
            )

    # Try tighter line spacing (reduce height estimate by 15%)
    capacity_tight = _estimate_zone_capacity(
        width_pct, height_pct * 1.15, slide_w, slide_h, base_font_pt * min_scale,
    )
    if total <= capacity_tight:
        return OverflowResult(
            strategy="spacing_reduction",
            font_scale=min_scale,
            kept_paragraphs=list(paragraphs),
        )

    # Split at paragraph boundary
    capacity_at_min = _estimate_zone_capacity(
        width_pct, height_pct, slide_w, slide_h, base_font_pt * min_scale,
    )
    kept = []
    overflow = []
    chars_used = 0
    split = False
    for p in paragraphs:
        if not split and chars_used + len(p.text) <= capacity_at_min:
            kept.append(p)
            chars_used += len(p.text)
        else:
            split = True
            overflow.append(p)

    # Ensure at least one paragraph is kept
    if not kept and overflow:
        kept = [overflow.pop(0)]

    return OverflowResult(
        strategy="split",
        font_scale=min_scale,
        kept_paragraphs=kept,
        overflow_paragraphs=overflow,
    )
