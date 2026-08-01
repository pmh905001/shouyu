"""Screenshot text extraction, backed by RapidOCR.

See docs/screenshot-ocr-design.md. RapidOCR returns recognized text as a
flat list of (quad_box, text, confidence) triples with no guaranteed
reading order - this module regroups them into lines by vertical position
and sorts each line left-to-right before joining, so multi-line screenshots
(code, chat, logs) come out close to their original layout instead of
scrambled. OCR is inherently imperfect on code/terminal screenshots (easy
to confuse 0/O, 1/l, full/half-width punctuation) - always worth a glance
before trusting it verbatim.
"""
from __future__ import annotations

from typing import List

import numpy as np
from PIL import Image

_engine = None


def _get_engine():
    global _engine
    if _engine is None:
        from rapidocr_onnxruntime import RapidOCR

        _engine = RapidOCR()
    return _engine


def extract_text(image: Image.Image) -> str:
    """Run OCR on a PIL image and return its text, reconstructed line by
    line (top-to-bottom, left-to-right within each line). Returns '' if
    nothing was recognized."""
    engine = _get_engine()
    result, _ = engine(np.array(image.convert("RGB")))
    if not result:
        return ""

    boxes = []
    for quad, text, _confidence in result:
        ys = [point[1] for point in quad]
        xs = [point[0] for point in quad]
        boxes.append(
            {
                "text": text,
                "y_center": sum(ys) / len(ys),
                "x_left": min(xs),
                "height": max(ys) - min(ys),
            }
        )
    boxes.sort(key=lambda b: b["y_center"])

    lines: List[List[dict]] = []
    for box in boxes:
        if lines:
            last_line = lines[-1]
            ref_y = sum(b["y_center"] for b in last_line) / len(last_line)
            ref_height = max((b["height"] for b in last_line), default=1) or 1
            if abs(box["y_center"] - ref_y) <= ref_height * 0.6:
                last_line.append(box)
                continue
        lines.append([box])

    line_texts = []
    for line in lines:
        line.sort(key=lambda b: b["x_left"])
        line_texts.append(" ".join(b["text"] for b in line))
    return "\n".join(line_texts)
