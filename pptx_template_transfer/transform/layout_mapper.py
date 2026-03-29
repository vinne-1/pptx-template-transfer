"""Layout mapper — assigns content to spatial zones for smarter placement.

Analyzes slide content density and type to determine optimal zone
assignments (e.g., left-text/right-image, full-width, grid) before
the slide builder places elements.
"""
from __future__ import annotations

import logging
from typing import Any

from pptx_template_transfer.models import ContentData, TemplateStyle

log = logging.getLogger("pptx_template_transfer")


def map_content_to_layout(
    content_list: list[ContentData],
    style: TemplateStyle,
) -> list[dict[str, Any]] | None:
    """Return per-slide zone assignments for layout-aware placement.

    Each entry in the returned list corresponds to a slide and contains
    a dict with zone hints such as ``{"layout": "two-column",
    "text_zone": "left", "image_zone": "right"}``.

    Returns ``None`` when no meaningful layout enhancement is possible
    (e.g., all slides are title/section type).
    """
    if not content_list:
        return None

    assignments: list[dict[str, Any]] = []

    for cd in content_list:
        zone: dict[str, Any] = {"layout": "default"}

        # Title / section slides don't need zone mapping
        if cd.slide_type in ("title", "section"):
            zone["layout"] = cd.slide_type
            assignments.append(zone)
            continue

        has_images = bool(cd.images)
        has_table = bool(cd.tables)
        has_text = bool(cd.body_paragraphs) or bool(cd.text_blocks)

        if has_images and has_text and not has_table:
            # Two-column: text left, images right
            zone["layout"] = "two-column"
            zone["text_zone"] = "left"
            zone["image_zone"] = "right"
        elif has_table:
            # Full-width for tables
            zone["layout"] = "full-width"
        elif has_images and not has_text:
            # Image-only — centered or grid
            zone["layout"] = "image-grid" if len(cd.images) > 2 else "image-center"
        else:
            zone["layout"] = "text-only"

        assignments.append(zone)

    # If every slide got "default" or a trivial type, return None
    non_trivial = [
        a for a in assignments
        if a["layout"] not in ("default", "title", "section")
    ]
    if not non_trivial:
        return None

    return assignments
