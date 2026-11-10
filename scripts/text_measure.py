#!/usr/bin/env python3
"""Shared CJK-aware text estimation for ppt-suggestion.

Used by render_pptx.py (layout: how tall a block is) and validate_pptx.py
(overflow check) so both measure text with the same ruler.

Conventions:
  - CJK character (U+2E80 and above) counts as CJK_WEIGHT
  - other characters count as LATIN_WEIGHT
  - char advance ≈ font-size points (CJK) scaled by weight
  - line height ≈ 1.22 x font-size (typical PowerPoint leading base)

Calibration: set_calibration(cjk_weight, latin_weight) lets a project tune
the weights; defaults are 1.0 / 0.55.
"""

from __future__ import annotations

import math

CJK_WEIGHT = 1.0
LATIN_WEIGHT = 0.55
LINE_HEIGHT_FACTOR = 1.22


def set_calibration(cjk_weight: float = 1.0, latin_weight: float = 0.55) -> None:
    global CJK_WEIGHT, LATIN_WEIGHT
    CJK_WEIGHT = cjk_weight
    LATIN_WEIGHT = latin_weight


def char_weight(ch: str) -> float:
    return CJK_WEIGHT if ord(ch) > 0x2E80 else LATIN_WEIGHT


def effective_len(text: str) -> float:
    return sum(char_weight(c) for c in text)


def estimate_width(text: str, size_pt: float) -> float:
    """Estimated occupied width in inches (single line, no wrap)."""
    return max(0.0, effective_len(text) * size_pt / 72.0)


def estimate_lines(text: str, size_pt: float, width_in: float) -> int:
    """Rough number of wrapped lines for the given box width."""
    cpl = max(1.0, (width_in * 72.0) / size_pt)
    lines = 0
    for raw in text.split("\n"):
        lines += max(1, math.ceil(effective_len(raw) / cpl))
    return lines


def estimate_height(text: str, size_pt: float, width_in: float,
                    line_spacing: float = 1.15) -> float:
    """Estimated occupied height in inches (wrapped)."""
    lines = estimate_lines(text, size_pt, width_in)
    return lines * size_pt * LINE_HEIGHT_FACTOR * line_spacing / 72.0


def estimate_occupied(text: str, size_pt: float, width_in: float,
                      box_w_in: float, box_h_in: float,
                      line_spacing: float = 1.15):
    """(w, h) inches actually occupied by text inside a box (clamped to box)."""
    w = min(box_w_in, estimate_width(text, size_pt))
    h = min(box_h_in, estimate_height(text, size_pt, width_in, line_spacing))
    return w, h


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 4:
        text, size, width = sys.argv[1], float(sys.argv[2]), float(sys.argv[3])
        print(f"len={effective_len(text):.2f} width={estimate_width(text, size):.2f}in "
              f"lines={estimate_lines(text, size, width)} "
              f"height={estimate_height(text, size, width):.2f}in")
    else:
        sys.stderr.write("usage: text_measure.py <text> <size_pt> <width_in>\n")
        sys.exit(2)
