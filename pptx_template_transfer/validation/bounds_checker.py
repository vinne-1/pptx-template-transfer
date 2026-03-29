"""Detect shapes running off slide edges."""
from __future__ import annotations

from pptx_template_transfer.models import BoundsIssue


def _has_text_content(shape) -> bool:
    """Return True if the shape contains meaningful text."""
    return (
        hasattr(shape, 'has_text_frame')
        and shape.has_text_frame
        and bool(shape.text_frame.text.strip())
    )


def check_bounds(slide, slide_index: int,
                 slide_w: int, slide_h: int) -> list[BoundsIssue]:
    """Check all shapes on a slide for out-of-bounds placement.

    Decorative shapes (no text) are allowed to extend beyond slide edges
    since this is a common design pattern for background accents.
    Only text-bearing shapes trigger bounds warnings.
    """
    issues: list[BoundsIssue] = []

    for shape in slide.shapes:
        left = shape.left or 0
        top = shape.top or 0
        width = shape.width or 0
        height = shape.height or 0
        right = left + width
        bottom = top + height

        edges = []
        overflow = 0

        if left < 0:
            edges.append("left")
            overflow = max(overflow, abs(left))
        if top < 0:
            edges.append("top")
            overflow = max(overflow, abs(top))
        if right > slide_w:
            edges.append("right")
            overflow = max(overflow, right - slide_w)
        if bottom > slide_h:
            edges.append("bottom")
            overflow = max(overflow, bottom - slide_h)

        if edges:
            # Only report if overflow is significant (>1% of slide dimension)
            min_dim = min(slide_w, slide_h)
            if overflow > min_dim * 0.01:
                # Skip decorative (no-text) shapes — off-slide accents are
                # intentional design elements.
                if not _has_text_content(shape):
                    continue
                issues.append(BoundsIssue(
                    slide_index=slide_index,
                    shape_name=shape.name or "unnamed",
                    edges=edges,
                    overflow_px=overflow,
                ))

    return issues
