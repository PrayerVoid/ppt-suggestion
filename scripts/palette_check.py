"""Small Python fallback for palette checks when the Node validator is unavailable."""
from __future__ import annotations

import argparse
import colorsys
import re
import sys


def _rgb(value):
    value = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", value):
        raise ValueError(f"无效颜色：{value!r}")
    return tuple(int(value[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _luminance(rgb):
    channels = []
    for c in rgb:
        channels.append(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def contrast(a, b):
    light, dark = sorted((_luminance(a), _luminance(b)), reverse=True)
    return (light + 0.05) / (dark + 0.05)


def check(colors, mode="light"):
    parsed = [_rgb(item) for item in colors]
    if not parsed:
        return ["色板为空"]
    issues = []
    surface = _rgb("FFFFFF" if mode == "light" else "111111")
    for index, color in enumerate(parsed):
        _, lightness, _ = colorsys.rgb_to_hls(*color)
        if mode == "light" and lightness > 0.82:
            issues.append(f"第 {index + 1} 个颜色在浅色背景上过浅")
        if mode == "dark" and lightness < 0.18:
            issues.append(f"第 {index + 1} 个颜色在深色背景上过暗")
        if contrast(color, surface) < 3:
            issues.append(f"第 {index + 1} 个颜色与背景对比度低于 3:1")
    for index, (left, right) in enumerate(zip(parsed, parsed[1:]), 1):
        if contrast(left, right) < 1.25:
            issues.append(f"相邻颜色 {index}/{index + 1} 对比不足，需增加标签或纹理")
    return issues


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("colors", help="逗号分隔的 HEX 色值")
    ap.add_argument("--mode", choices=("light", "dark"), default="light")
    args = ap.parse_args()
    try:
        issues = check(args.colors.split(","), args.mode)
    except ValueError as exc:
        print(exc, file=sys.stderr)
        return 2
    if issues:
        print("\n".join(issues))
        return 1
    print(f"palette_check: {args.mode} 通过（Python fallback）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
