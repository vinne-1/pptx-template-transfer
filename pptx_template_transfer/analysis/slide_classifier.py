"""Shape role classification and slide zone detection."""
from __future__ import annotations

import re
from typing import Any

from pptx_template_transfer.models import ShapeInfo, Thresholds
from pptx_template_transfer.helpers import (
    FOOTER_PATTERNS, PH_TITLE_SET, PH_BODY_SET, PH_FOOTER_SET,
    text_of, word_count, max_font_pt, shape_area_pct,
    shape_top_frac, shape_bottom_frac, shape_left_frac,
    is_picture, is_table, is_chart, is_group, is_ole_or_embedded,
    is_allcaps_short, group_text_words, placeholder_type_int,
)


def _precompute_shape_info(shape, slide_w: int, slide_h: int) -> ShapeInfo:
    text = text_of(shape)
    is_grp = is_group(shape)
    return ShapeInfo(
        shape=shape,
        text=text,
        word_count=word_count(text),
        font_size=max_font_pt(shape),
        area_pct=shape_area_pct(shape, slide_w, slide_h),
        top_frac=shape_top_frac(shape, slide_h),
        bottom_frac=shape_bottom_frac(shape, slide_h),
        left_frac=shape_left_frac(shape, slide_w),
        is_picture=is_picture(shape),
        is_table=is_table(shape),
        is_chart=is_chart(shape),
        is_group=is_grp,
        is_ole=is_ole_or_embedded(shape),
        placeholder_type=placeholder_type_int(shape),
        name_lower=(shape.name or "").lower(),
        group_text_words=group_text_words(shape) if is_grp else 0,
    )


def _classify_shape(
    si: ShapeInfo,
    th: Thresholds,
    *,
    largest_font: float,
    median_font: float,
    title_assigned: bool,
    body_count: int,
    info_count: int,
    similar_ids: set,
) -> tuple[str, float]:
    """Classify a shape's role. Returns (role, confidence)."""

    if si.placeholder_type is not None:
        if si.placeholder_type in PH_TITLE_SET:
            return ("title", 0.95) if not title_assigned else ("decorative", 0.6)
        if si.placeholder_type in PH_BODY_SET:
            return ("body", 0.95) if body_count < th.body_max_zones else ("decorative", 0.5)
        if si.placeholder_type in PH_FOOTER_SET:
            return ("footer", 0.95)

    if si.is_picture or si.is_chart or si.is_table:
        return ("media", 0.95)
    if si.is_ole:
        return ("media", 0.90)
    if si.is_group:
        if si.group_text_words > 20 and body_count < th.body_max_zones:
            return ("body", 0.55)
        return ("media", 0.80)

    if si.bottom_frac >= th.footer_bottom_frac and si.area_pct < th.footer_max_area_pct:
        return ("footer", 0.85)
    if si.top_frac <= th.footer_top_frac and si.area_pct < 3 and si.word_count <= 10:
        return ("footer", 0.80)
    if si.text and FOOTER_PATTERNS.search(si.text):
        return ("footer", 0.85)

    if not si.text:
        return ("decorative", 0.90)

    if si.area_pct < th.decorative_max_area_pct and si.word_count <= th.decorative_max_words:
        return ("decorative", 0.80)
    if 0 < si.font_size <= th.decorative_max_font_pt:
        return ("decorative", 0.75)
    if si.word_count <= 3:
        return ("decorative", 0.70)
    if is_allcaps_short(si.text) and si.area_pct < 5:
        return ("decorative", 0.75)
    if re.match(r"^\d{1,2}$", si.text.strip()):
        return ("decorative", 0.90)
    if id(si.shape) in similar_ids:
        return ("decorative", 0.70)

    effective_title_font = th.title_min_font_pt
    if median_font > 0 and median_font < 14:
        effective_title_font = max(median_font * 1.3, 14)

    conf = 0.0
    if not title_assigned and si.top_frac < 0.45 and si.word_count <= th.title_max_words:
        if si.font_size >= effective_title_font and si.font_size >= largest_font - 2:
            conf = 0.85
        if any(kw in si.name_lower for kw in ("title", "heading")):
            conf = max(conf, 0.70)
        if conf >= 0.55:
            return ("title", conf)

    if body_count < th.body_max_zones:
        if si.area_pct > th.body_min_area_pct and si.word_count > th.body_min_words:
            conf = 0.80
            if any(kw in si.name_lower for kw in ("body", "content", "text")):
                conf = 0.85
            return ("body", conf)
        if si.area_pct > th.body_min_area_pct_relaxed and si.word_count > th.body_min_words_relaxed:
            return ("body", 0.60)

    if (info_count < 1
            and si.left_frac >= th.info_left_frac
            and th.info_min_words <= si.word_count <= th.info_max_words
            and si.area_pct > 2):
        return ("info", 0.65)

    return ("decorative", 0.40)


def _detect_repeated_patterns(infos: list[ShapeInfo], slide_w: int, slide_h: int) -> set:
    result: set[int] = set()
    if len(infos) < 3:
        return result

    dimension_groups: dict[tuple, list] = {}
    for si in infos:
        w = si.shape.width or 0
        h = si.shape.height or 0
        if w == 0 or h == 0:
            continue
        bw = round(w / (slide_w * 0.02)) if slide_w else 0
        bh = round(h / (slide_h * 0.02)) if slide_h else 0
        dimension_groups.setdefault((bw, bh), []).append(si)

    for grp in dimension_groups.values():
        if len(grp) < 3:
            continue
        top_buckets: dict[int, int] = {}
        for si in grp:
            bucket = round((si.shape.top or 0) / (slide_h * 0.05)) if slide_h else 0
            top_buckets[bucket] = top_buckets.get(bucket, 0) + 1
        if top_buckets and max(top_buckets.values()) >= 3:
            for si in grp:
                if si.word_count <= 15:
                    result.add(id(si.shape))
    return result


def classify_all_shapes(
    slide, slide_w: int, slide_h: int, th: Thresholds,
) -> list[tuple[Any, str, float]]:
    """Classify all shapes. Returns [(shape, role, confidence), ...]."""
    shapes = list(slide.shapes)
    infos = [_precompute_shape_info(s, slide_w, slide_h) for s in shapes]

    fonts = [si.font_size for si in infos if si.font_size > 0]
    largest_font = max(fonts) if fonts else 0.0
    sorted_fonts = sorted(fonts)
    median_font = sorted_fonts[len(sorted_fonts) // 2] if sorted_fonts else 0.0

    similar_ids = _detect_repeated_patterns(infos, slide_w, slide_h)

    sorted_infos = sorted(infos, key=lambda si: ((si.shape.top or 0), (si.shape.left or 0)))

    title_assigned = False
    body_count = 0
    info_count = 0
    results: dict[int, tuple[str, float]] = {}

    for si in sorted_infos:
        role, conf = _classify_shape(
            si, th,
            largest_font=largest_font,
            median_font=median_font,
            title_assigned=title_assigned,
            body_count=body_count,
            info_count=info_count,
            similar_ids=similar_ids,
        )
        results[id(si.shape)] = (role, conf)
        if role == "title":
            title_assigned = True
        elif role == "body":
            body_count += 1
        elif role == "info":
            info_count += 1

    return [(s, *results[id(s)]) for s in shapes]


def classify_shape_role(
    shape, slide_width: int, slide_height: int,
    slide=None, th: Thresholds | None = None,
) -> str:
    if th is None:
        th = Thresholds()
    si = _precompute_shape_info(shape, slide_width, slide_height)
    role, _ = _classify_shape(
        si, th, largest_font=si.font_size, median_font=si.font_size,
        title_assigned=False, body_count=0, info_count=0, similar_ids=set(),
    )
    if role == "media":
        return "decorative"
    if role == "info":
        return "body"
    return role


def get_slide_zones(
    slide, slide_width: int, slide_height: int, th: Thresholds | None = None,
) -> dict[str, list]:
    if th is None:
        th = Thresholds()
    classifications = classify_all_shapes(slide, slide_width, slide_height, th)
    zones: dict[str, list] = {"title": [], "body": [], "decorative": [], "footer": []}
    for shape, role, _conf in classifications:
        if role == "title":
            zones["title"].append(shape)
        elif role in ("body", "info"):
            zones["body"].append(shape)
        elif role == "footer":
            zones["footer"].append(shape)
        else:
            zones["decorative"].append(shape)
    return zones


# ============================================================================
# ENHANCED SLIDE TYPE CLASSIFICATION (Phase 2)
# ============================================================================

def classify_slide_type(slide, slide_index: int, total_slides: int,
                        slide_w: int = 0, slide_h: int = 0) -> str:
    """Enhanced slide type classification with 15+ archetypes."""
    from pptx_template_transfer.helpers import is_picture as _is_pic, is_table as _is_tbl, is_chart as _is_ch

    shapes = list(slide.shapes)
    text_shapes = [s for s in shapes if text_of(s)]
    images = [s for s in shapes if _is_pic(s)]
    tables = [s for s in shapes if _is_tbl(s)]
    charts = [s for s in shapes if _is_ch(s)]

    total_words = sum(word_count(text_of(s)) for s in text_shapes)
    big = [s for s in text_shapes if max_font_pt(s) >= 20]
    all_text = " ".join(text_of(s) for s in text_shapes).lower()

    # --- Definitive types ---
    if not text_shapes and not images:
        return "blank"
    if not text_shapes and images:
        return "image_heavy"

    # --- Title slide ---
    if slide_index == 0 and big:
        return "title"
    if total_words <= 20 and big and len(text_shapes) <= 5:
        return "title"

    # --- Section divider ---
    if len(text_shapes) <= 3 and total_words <= 15 and big:
        return "section_divider"

    # --- Data-heavy ---
    if tables or charts:
        return "data_table"

    # --- Agenda (check BEFORE closing — agenda slides list closing-sounding items) ---
    agenda_kw = {"agenda", "outline", "table of contents"}
    title_text = ""
    if big:
        title_text = text_of(big[0]).lower()
    # "overview" only counts as agenda if standalone or preceded by generic words
    # (not "incident overview", "deployment overview", etc.)
    _overview_disqualifiers = {"incident", "case", "deployment", "security", "service", "project"}
    if "overview" in title_text:
        words_before = title_text.split("overview")[0].strip().split()
        if not words_before or not any(w in _overview_disqualifiers for w in words_before):
            agenda_kw = agenda_kw | {"overview"}
    if any(kw in title_text for kw in agenda_kw):
        return "agenda"
    if any(kw in all_text for kw in agenda_kw) and total_words <= 40:
        return "agenda"

    # --- Closing/conclusion ---
    closing_kw = {"thank", "contact", "questions", "q&a", "reference", "next steps", "conclusion"}
    if slide_index == total_slides - 1 and total_words <= 40:
        return "closing"
    if any(kw in all_text for kw in closing_kw) and total_words <= 60:
        return "closing"

    # --- Metrics/KPI dashboard ---
    # Look for 3+ short text blocks containing numbers
    number_shapes = []
    for s in text_shapes:
        t = text_of(s).strip()
        if re.match(r"^[\d,.%$€£¥]+[kKmMbB]?$", t) and word_count(t) <= 2:
            number_shapes.append(s)
    if len(number_shapes) >= 3:
        return "metrics_dashboard"

    # --- Process flow ---
    numbered = [s for s in text_shapes if re.match(r"^\d{1,2}$", text_of(s).strip())]
    if len(numbered) >= 3:
        return "process_flow"

    # --- Comparison ---
    comparison_kw = {"vs", "versus", "comparison", "compared", "before", "after"}
    if any(kw in all_text for kw in comparison_kw):
        return "comparison"

    # --- Timeline ---
    year_pattern = re.compile(r"\b20[12]\d\b")
    year_matches = year_pattern.findall(all_text)
    if len(set(year_matches)) >= 3:
        return "timeline"

    # --- Image-heavy ---
    if len(images) >= 3 and total_words < 30:
        return "image_heavy"
    if len(images) >= 2 and slide_w > 0:
        img_area = sum(shape_area_pct(s, slide_w, slide_h) for s in images)
        if img_area > 40:
            return "image_heavy"

    # --- Content refinement ---
    # Count bullet-level paragraphs
    bullet_count = 0
    for s in text_shapes:
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                if p.level and p.level > 0:
                    bullet_count += 1

    if bullet_count >= 5:
        return "content_bullets"

    if total_words > 20:
        return "content_narrative"

    return "section_divider"


def classify_template_structure(
    slide, slide_w: int, slide_h: int,
    slide_index: int = -1, total_slides: int = -1,
) -> str:
    """Classify a template slide's visual structure."""
    shapes = list(slide.shapes)
    text_shapes = [s for s in shapes if text_of(s)]
    images = [s for s in shapes if is_picture(s)]
    tables = [s for s in shapes if is_table(s)]
    charts = [s for s in shapes if is_chart(s)]

    total_words = sum(word_count(text_of(s)) for s in text_shapes)
    big = [s for s in text_shapes if max_font_pt(s) >= 20]

    if tables or charts:
        return "data"
    if len(images) >= 2 and total_words < 30:
        return "visual"

    numbered = [s for s in text_shapes if re.match(r"^\d{1,2}$", text_of(s).strip())]
    if len(numbered) >= 3:
        infos = [_precompute_shape_info(s, slide_w, slide_h) for s in shapes]
        if _detect_repeated_patterns(infos, slide_w, slide_h):
            return "grid"
        return "list"

    if slide_index == 0 and big:
        return "title"

    closing_kw = {"thank", "contact", "questions", "q&a", "reference"}
    all_lower = " ".join(text_of(s) for s in text_shapes).lower()
    if slide_index == total_slides - 1 and total_slides > 1:
        if any(kw in all_lower for kw in closing_kw) or total_words <= 40:
            return "closing"

    if big and total_words <= 20 and len(text_shapes) <= 5:
        return "title"
    if big and len(text_shapes) <= 3 and total_words <= 15:
        return "section"
    if any(kw in all_lower for kw in closing_kw) and total_words <= 40:
        return "closing"

    body = [s for s in text_shapes if shape_area_pct(s, slide_w, slide_h) > 4 and word_count(text_of(s)) > 10]
    if body:
        return "narrative"
    return "narrative" if total_words > 20 else "section"
