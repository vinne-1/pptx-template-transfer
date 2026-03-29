"""Detect semantic block types within body paragraphs."""
from __future__ import annotations

import re

from pptx_template_transfer.models import ParagraphData, SemanticBlock


_NUMBERED_RE = re.compile(r"^\d+[.)]\s|^Step\s+\d+", re.I)
_KV_RE = re.compile(r"^[A-Z][\w\s]{1,30}:\s")
_METRIC_RE = re.compile(r"^[\d,.]+[%$€£¥kKmMbB]*$")


def detect_semantic_blocks(paragraphs: list[ParagraphData]) -> list[SemanticBlock]:
    """Analyze body paragraphs and group them into semantic blocks."""
    if not paragraphs:
        return []

    blocks: list[SemanticBlock] = []
    current_type = "plain"
    current_paras: list[ParagraphData] = []

    def _flush():
        nonlocal current_paras, current_type
        if current_paras:
            label = ""
            if current_type == "section_header" and current_paras:
                label = current_paras[0].text
            blocks.append(SemanticBlock(
                block_type=current_type,
                paragraphs=list(current_paras),
                label=label,
            ))
            current_paras = []
            current_type = "plain"

    for p in paragraphs:
        text = p.text.strip()
        if not text:
            continue

        # Detect type of this paragraph
        if p.bold and len(text.split()) <= 6:
            # Section header — flush previous, start new section
            _flush()
            current_type = "section_header"
            current_paras = [p]
            _flush()
            continue

        if _NUMBERED_RE.match(text):
            if current_type != "numbered_list":
                _flush()
                current_type = "numbered_list"
            current_paras.append(p)
            continue

        if _KV_RE.match(text):
            if current_type != "key_value":
                _flush()
                current_type = "key_value"
            current_paras.append(p)
            continue

        if _METRIC_RE.match(text) and len(text) <= 15:
            if current_type != "metric_group":
                _flush()
                current_type = "metric_group"
            current_paras.append(p)
            continue

        # Default: plain paragraph
        if current_type != "plain":
            _flush()
            current_type = "plain"
        current_paras.append(p)

    _flush()
    return blocks
