"""Preflight checks for text capacity and replacement-image resolution."""
from __future__ import annotations

import math
import struct
from pathlib import Path

try:
    from PIL import Image
except ImportError:  # Pillow is optional; PNG/JPEG headers still suffice.
    Image = None

from validate_pptx import _inches, estimate_text_height, paragraph_width, paragraph_size

EMU_IN = 914400.0
PX_PER_IN = 96.0
MIN_SOURCE_RATIO = 0.70


def _vertical(shape) -> bool:
    try:
        body_pr = shape._element.xpath('.//*[local-name()="bodyPr"]')
        value = body_pr[0].get("vert", "horz") if body_pr else "horz"
        return value not in {"horz", "none", ""}
    except Exception:
        return False


def _label_capacity(shape, height):
    """Return a rough character capacity for vertical template labels."""
    tf = shape.text_frame
    margin = _inches(tf.margin_top) + _inches(tf.margin_bottom)
    usable = max(0.01, height - margin)
    size = max((paragraph_size(p) for p in tf.paragraphs), default=14.0)
    return max(1, math.floor(usable / (size / 72 * 1.22)))


def check_text(shape, text, label, width=None, height=None, size=None):
    """Return hard capacity problems for a proposed text replacement."""
    if not getattr(shape, "has_text_frame", False):
        return []
    # Never mutate the source during preflight: assigning shape.text discards
    # run formatting, so restoring only the string corrupts subsequent checks.
    from copy import deepcopy
    shape = deepcopy(shape)
    width = _inches(width if width is not None else shape.width)
    height = _inches(height if height is not None else shape.height)
    value = str(text)
    if not value.strip():
        return []
    from slide_clone import replace_text
    replace_text(shape, value)
    tf = shape.text_frame
    if size is not None:
        from power_pptx.util import Pt
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(float(size))
    if _vertical(shape):
        capacity = _label_capacity(shape, height)
        count = len(value.replace("\n", ""))
        if count > capacity:
            return [f"{label} 竖排文本容量不足：{count} 字超过约 {capacity} 字，请缩短标签或换页"]
        return []
    estimated = estimate_text_height(tf, width)
    if estimated > height + max(0.04, height * 0.10):
        return [f"{label} 文本框溢出，容量不足：估算高度 {estimated:.2f}in 超过文本框 {height:.2f}in，请换更大版式或拆页"]
    if tf.word_wrap is False:
        longest = max((paragraph_width(p, line) for p in tf.paragraphs
                       for line in p.text.replace("\v", "\n").split("\n")), default=0)
        inner = max(0.01, width - _inches(tf.margin_left) - _inches(tf.margin_right))
        if longest > inner + 0.04:
            return [f"{label} 不换行文本过宽：估算 {longest:.2f}in 超过可用宽度 {inner:.2f}in，请缩短或换版式"]
    return []


def check_box(text, width, height, size, label):
    """Check an explicitly sized new text box before OfficeCLI runs."""
    from text_measure import estimate_width

    width = float(width) / EMU_IN
    height = float(height) / EMU_IN
    usable = max(0.01, width - 0.14)
    lines = 0
    for line in str(text).replace("\\n", "\n").split("\n"):
        lines += max(1, math.ceil(estimate_width(line, float(size)) / usable))
    estimated = lines * float(size) / 72 * 1.22 + 0.14
    if estimated > height + max(0.04, height * 0.10):
        return [f"{label} 文本容量不足：估算高度 {estimated:.2f}in 超过文本框 {height:.2f}in"]
    return []


def _image_size(path):
    if Image is not None:
        with Image.open(path) as image:
            return image.size
    data = Path(path).read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:2] == b"\xff\xd8":
        pos = 2
        while pos + 9 < len(data):
            if data[pos] != 0xFF:
                pos += 1
                continue
            marker = data[pos + 1]
            pos += 2
            if marker in {0xD8, 0xD9}:
                continue
            if pos + 2 > len(data):
                break
            length = struct.unpack(">H", data[pos:pos + 2])[0]
            if marker in set(range(0xC0, 0xC4)) | set(range(0xC5, 0xC8)) | set(range(0xC9, 0xCC)) | set(range(0xCD, 0xD0)):
                if pos + 7 <= len(data):
                    return struct.unpack(">HH", data[pos + 3:pos + 7])[::-1]
                break
            pos += length
    raise ValueError("仅支持可读取尺寸的 PNG/JPEG 图片")


def check_image(path, shape, label):
    """Return a hard problem when a replacement image is enlarged too far."""
    if Path(path).suffix.lower() == ".svg":
        return []  # Vector sources do not lose detail through pixel scaling.
    try:
        native_w, native_h = _image_size(path)
    except Exception as exc:
        return [f"{label} 图片无法读取：{exc}"]
    display_w = max(0.01, _inches(shape.width)) * PX_PER_IN
    display_h = max(0.01, _inches(shape.height)) * PX_PER_IN
    ratio_w = native_w / display_w
    ratio_h = native_h / display_h
    if min(ratio_w, ratio_h) < MIN_SOURCE_RATIO:
        return [f"{label} 图片源分辨率不足：{native_w}×{native_h}px 显示在约 {display_w:.0f}×{display_h:.0f}px，放大倍率过高"]
    return []
