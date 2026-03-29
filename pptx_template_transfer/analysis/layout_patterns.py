"""Mine layout patterns from a target deck — detect archetypes like 2-col, card grid, sidebar+main."""
from __future__ import annotations

from collections import defaultdict

from pptx import Presentation

from pptx_template_transfer.models import LayoutPattern, LayoutZone, Thresholds
from pptx_template_transfer.helpers import (
    text_of, word_count, max_font_pt, shape_area_pct,
    is_picture, is_table, is_chart,
)


def _detect_columns(shapes: list, slide_w: int) -> int:
    """Detect how many columns a set of shapes forms by clustering left positions."""
    if not shapes or slide_w == 0:
        return 1

    # Get left positions as percentages
    lefts = sorted(set(
        round((s.left or 0) / slide_w * 100 / 5) * 5  # Round to 5% buckets
        for s in shapes
        if (s.width or 0) / slide_w < 0.8  # Skip full-width shapes
    ))

    if len(lefts) <= 1:
        return 1

    # Count distinct position clusters (>15% apart)
    clusters = [lefts[0]]
    for l in lefts[1:]:
        if l - clusters[-1] >= 15:
            clusters.append(l)

    return min(len(clusters), 4)


def _estimate_text_capacity(width_pct: float, height_pct: float,
                            slide_w: int, slide_h: int) -> int:
    """Estimate characters that fit in a zone."""
    w_inches = (slide_w * width_pct / 100) / 914400.0
    h_inches = (slide_h * height_pct / 100) / 914400.0
    # ~12pt font, ~10 chars per inch width, ~3 lines per inch height
    chars_per_line = max(1, int(w_inches * 10))
    lines = max(1, int(h_inches * 3))
    return chars_per_line * lines


def _classify_pattern_name(zones: list[LayoutZone], col_count: int,
                           has_image: bool, has_table: bool) -> str:
    """Name a layout pattern based on its zone geometry."""
    body_zones = [z for z in zones if z.role == "body"]
    image_zones = [z for z in zones if z.role == "image"]

    if has_table:
        return "data-table"
    if col_count >= 3:
        return "3-card-grid" if len(body_zones) >= 3 else "3-col"
    if col_count == 2:
        if image_zones and body_zones:
            return "image+text-split"
        return "2-col"
    if has_image and body_zones:
        # Check if image is on the side
        if image_zones:
            img_z = image_zones[0]
            if img_z.left_pct > 50:
                return "text+right-image"
            elif img_z.left_pct < 20 and img_z.width_pct < 50:
                return "left-image+text"
        return "text+image"
    if len(body_zones) == 1 and body_zones[0].width_pct > 70:
        return "full-narrative"
    if not body_zones:
        title_zones = [z for z in zones if z.role == "title"]
        if title_zones:
            return "title-centered"
        return "minimal"
    return "standard"


def mine_layout_patterns(prs: Presentation,
                         th: Thresholds | None = None) -> list[LayoutPattern]:
    """Scan a presentation and detect layout patterns for each slide."""
    if th is None:
        th = Thresholds()

    sw, sh = prs.slide_width, prs.slide_height
    patterns: list[LayoutPattern] = []

    for si, slide in enumerate(prs.slides):
        shapes = list(slide.shapes)
        zones: list[LayoutZone] = []
        has_image = False
        has_table = False

        # Classify each shape into a zone
        for s in shapes:
            left_pct = (s.left or 0) / sw * 100 if sw else 0
            top_pct = (s.top or 0) / sh * 100 if sh else 0
            width_pct = (s.width or 0) / sw * 100 if sw else 0
            height_pct = (s.height or 0) / sh * 100 if sh else 0

            if is_picture(s):
                has_image = True
                area = shape_area_pct(s, sw, sh)
                if area >= th.image_min_area_pct:
                    zones.append(LayoutZone(
                        role="image",
                        left_pct=left_pct, top_pct=top_pct,
                        width_pct=width_pct, height_pct=height_pct,
                    ))
                continue

            if is_table(s):
                has_table = True
                zones.append(LayoutZone(
                    role="table",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                ))
                continue

            if is_chart(s):
                has_table = True
                zones.append(LayoutZone(
                    role="chart",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                ))
                continue

            t = text_of(s)
            wc = word_count(t)
            fs = max_font_pt(s)

            # Skip empty or tiny decorative shapes
            if not t or (width_pct < 3 and height_pct < 3):
                if width_pct > 1 or height_pct > 1:
                    zones.append(LayoutZone(
                        role="accent",
                        left_pct=left_pct, top_pct=top_pct,
                        width_pct=width_pct, height_pct=height_pct,
                    ))
                continue

            # Footer zone
            if top_pct > 88:
                zones.append(LayoutZone(
                    role="footer",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                ))
                continue

            # Title zone (large font, near top)
            if fs >= 18 and top_pct < 30 and wc <= 15:
                cap = _estimate_text_capacity(width_pct, height_pct, sw, sh)
                zones.append(LayoutZone(
                    role="title",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                    text_capacity=cap,
                ))
                continue

            # Body zone
            area = shape_area_pct(s, sw, sh)
            if area > 2 and wc > 3:
                cap = _estimate_text_capacity(width_pct, height_pct, sw, sh)
                zones.append(LayoutZone(
                    role="body",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                    text_capacity=cap,
                ))
            elif wc <= 5 and area < 5:
                zones.append(LayoutZone(
                    role="card",
                    left_pct=left_pct, top_pct=top_pct,
                    width_pct=width_pct, height_pct=height_pct,
                    text_capacity=_estimate_text_capacity(width_pct, height_pct, sw, sh),
                ))

        # Detect column count from body shapes
        body_shapes = [s for s in shapes
                       if text_of(s) and shape_area_pct(s, sw, sh) > 2
                       and word_count(text_of(s)) > 3]
        col_count = _detect_columns(body_shapes, sw)

        # Calculate total text capacity
        total_cap = sum(z.text_capacity for z in zones if z.role in ("title", "body", "card"))

        name = _classify_pattern_name(zones, col_count, has_image, has_table)

        patterns.append(LayoutPattern(
            name=name,
            zones=zones,
            column_count=col_count,
            total_text_capacity=total_cap,
            source_slide_indices=[si],
            has_image_zone=has_image,
            has_table_zone=has_table,
        ))

    return patterns
