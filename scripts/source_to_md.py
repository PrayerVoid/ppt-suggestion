#!/usr/bin/env python3
"""Lightweight source-to-Markdown converter for ppt-suggestion.

Handles: HTML files, web URLs, plain text / Markdown.
PDF / DOCX / XLSX: use the Read tool (see references/source-intake.md).

Usage:
  python source_to_md.py <file_or_url> [-o out.md] [-t html|text]

Exit codes:
  0 success (prints output path)
  2 input unreadable / unsupported
"""

from __future__ import annotations

import argparse
import sys
import urllib.request
from pathlib import Path


def fetch_text(src: str) -> tuple[str, str]:
    """Return (content, kind) where kind is html|text."""
    if src.startswith("http://") or src.startswith("https://"):
        with urllib.request.urlopen(src, timeout=30) as resp:
            data = resp.read()
        head = data[:2048].decode("utf-8", "ignore")
        return data.decode("utf-8", "ignore"), ("html" if "<html" in head.lower() else "text")
    p = Path(src)
    if not p.exists():
        raise FileNotFoundError(src)
    data = p.read_bytes()
    head = data[:2048]
    if b"<html" in head.lower() or b"<!doctype html" in head.lower() or p.suffix.lower() in (".html", ".htm"):
        return data.decode("utf-8", "ignore"), "html"
    return data.decode("utf-8", "ignore"), "text"


def html_to_markdown(html: str) -> str:
    """Block-level HTML → Markdown using lxml (keeps tables/headings/lists)."""
    try:
        from lxml import html as lh
        from lxml import etree
    except ImportError:
        sys.stderr.write("需要 lxml（venv python 通常已带）\n")
        raise SystemExit(2)

    root = lh.fromstring(html)
    # drop script/style/nav
    for tag in ("script", "style", "nav", "header", "footer", "form", "aside"):
        for el in root.iter(tag):
            el.drop_tree()

    lines: list[str] = []
    body = root.body if root.body is not None else root

    def walk(el):
        for child in el.iterchildren():
            tag = child.tag
            if tag is etree.Comment:
                continue
            tag = (tag or "").lower()
            text = " ".join((child.text_content() or "").split())
            if tag in ("h1", "h2", "h3", "h4"):
                lines.append(("#" * min(int(tag[1]), 4)) + " " + text)
                lines.append("")
            elif tag == "p":
                if text:
                    lines.append(text)
                    lines.append("")
            elif tag in ("li",):
                lines.append("- " + text)
            elif tag in ("ul", "ol", "blockquote", "div", "section", "article",
                         "main", "body", "html"):
                walk(child)
            elif tag == "table":
                rows = []
                for tr in child.iter("tr"):
                    cells = [" ".join((td.text_content() or "").split())
                             for td in tr.iter("td") or tr.iter("th")]
                    if not cells:
                        cells = [" ".join((th.text_content() or "").split())
                                 for th in tr.iter("th")]
                    rows.append(cells)
                if rows:
                    lines.append("")
                    header = rows[0]
                    lines.append("| " + " | ".join(header) + " |")
                    lines.append("|" + "---|" * len(header))
                    for r in rows[1:]:
                        r = (r + [""] * len(header))[:len(header)]
                        lines.append("| " + " | ".join(r) + " |")
                    lines.append("")
            elif tag == "img":
                src_attr = child.get("src") or ""
                alt = child.get("alt") or ""
                if src_attr:
                    lines.append(f"![{alt}]({src_attr})")
                    lines.append("")
            elif tag in ("br",):
                continue
            elif text and tag not in ("td", "th", "tr", "span", "a", "strong", "em",
                                      "b", "i", "figure", "figcaption"):
                lines.append(text)
                lines.append("")

    walk(body)
    md = "\n".join(lines)
    md = "\n".join(filter(None, (l.strip() for l in md.splitlines())))
    return md


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("-o", "--output")
    ap.add_argument("-t", "--type", choices=["html", "text"], default=None)
    args = ap.parse_args()

    try:
        content, kind = fetch_text(args.source)
    except Exception as e:
        sys.stderr.write(f"读取失败: {e}\n")
        sys.stderr.write("提示：PDF/DOCX/XLSX 请用 Read 工具；仍失败则如实列缺口问用户。\n")
        return 2
    if args.type:
        kind = args.type

    md = html_to_markdown(content) if kind == "html" else content
    if not md.strip():
        sys.stderr.write("转换结果为空，请检查源文件\n")
        return 2

    if args.output:
        out = Path(args.output).expanduser()
        if not out.is_absolute():
            out = Path.cwd() / out
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(md, encoding="utf-8")
        print(str(out))
    else:
        print(md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
