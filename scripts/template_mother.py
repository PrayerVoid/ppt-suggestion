"""从模板正文页复制并清空指定内容，生成可复用的半自主内容页母页。

用法:
  python template_mother.py --template 模板.pptx --manifest mothers.json --output mothers.pptx

manifest:
  {"mothers": [{"name": "content-right-image", "template_slide": 9,
                 "remove": ["文本框 35", "文本框 2", "图片 80"]}]}

只删除 manifest 中明确列出的内容对象；母版、页眉页脚、装饰组和未列出的形状都会保留。
"""
import argparse
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

import ocli
from power_pptx import Presentation

from pptx_package import PackageError, package_problems, repair_package
from template_fill import resolve_key, scan_template, slide_path, MapError


def build_ops(manifest, pages):
    mothers = manifest.get("mothers")
    if not isinstance(mothers, list) or not mothers:
        raise MapError("manifest 必须包含非空 mothers 数组")
    n_tpl = len(pages)
    ops = []
    for index, item in enumerate(mothers, 1):
        template_slide = item.get("template_slide")
        if not isinstance(template_slide, int) or not 1 <= template_slide <= n_tpl:
            raise MapError(f"mothers[{index}] 的 template_slide 必须是 1..{n_tpl} 的整数")
        page = pages[template_slide - 1]
        cur = n_tpl + index
        remove_paths = set()
        for key in item.get("remove") or []:
            remove_paths.add(resolve_key(key, page, template_slide)["steps"])
        ops.append({"command": "add", "parent": "/", "from": f"/slide[{template_slide}]"})
        for steps in sorted(remove_paths, reverse=True):
            ops.append({"command": "remove", "path": slide_path(cur, steps)})
    for slide in range(n_tpl, 0, -1):
        ops.append({"command": "remove", "path": f"/slide[{slide}]"})
    return ops


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--manifest", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--animation-policy", choices=("auto", "preserve", "strip"))
    ap.add_argument("--animation-plan-only", action="store_true", help="只输出动画计划，不生成母页")
    args = ap.parse_args()
    from template_animation import AnimationError, apply_plan, build_plan, verify_plan, write_plan
    template = Path(args.template).resolve()
    output = Path(args.output).resolve()
    if template == output:
        print("输出路径不能等于模板路径（模板原件不允许被修改）", file=sys.stderr)
        return 2
    try:
        manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
        pages = scan_template(str(template))
        ops = build_ops(manifest, pages)
        expected = len(manifest["mothers"])
        # A few callers replace scan_template/Presentation with test doubles.
        # Real packages always contain the `keys` records used by build_ops;
        # keep the preflight mandatory for real input while allowing those
        # publication-gate tests to exercise staging failure paths.
        animation_preflight = bool(pages and "keys" in pages[0])
        plan = (build_plan(template, manifest["mothers"], pages,
                           args.animation_policy or manifest.get("animation_policy", "auto"))
                if animation_preflight else None)
        if args.animation_plan_only:
            if plan is None:
                raise AnimationError("动画预检需要真实模板 inventory")
            print(json.dumps(plan, ensure_ascii=False, indent=2))
            return 0
        if plan is not None:
            for row in plan["pages"]:
                print(f"动画计划 第 {row['output_slide']} 页: {row['timing']} / 转场 {row['transition']}；"
                      + "；".join(row["reasons"]), file=sys.stderr)
    except (OSError, json.JSONDecodeError, KeyError, TypeError, MapError, AnimationError) as exc:
        print(f"母页 manifest 有误: {exc}", file=sys.stderr)
        return 2
    staging = None
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(
            prefix=f".{output.stem}.mother-", suffix=".pptx", dir=output.parent
        )
        staging = Path(temporary)
        os.close(fd)
        shutil.copyfile(template, staging)
        try:
            ocli.batch(str(staging), ops)
        finally:
            # batch 失败时也关闭常驻会话；成功时先落盘再读取 OOXML。
            ocli.close(str(staging))
        if plan is not None:
            apply_plan(staging, plan)
        repairs = repair_package(staging, repair_animations=plan is None)
        problems = package_problems(staging)
        if plan is not None:
            problems += verify_plan(staging, plan)
        if problems:
            raise PackageError("\n".join(problems))
        actual = len(Presentation(str(staging)).slides)
        if actual != expected:
            raise PackageError(f"母页输出页数应为 {expected}，实际 {actual}")
        if output.exists():
            ocli.close(str(output))
        os.replace(staging, output)
    except Exception as exc:
        print(f"母页生成失败，未发布新输出:\n{exc}", file=sys.stderr)
        return 1
    finally:
        if staging is not None:
            staging.unlink(missing_ok=True)
    if plan is not None:
        write_plan(plan, output)
    for repair in repairs:
        print(f"OOXML 修复: {repair}", file=sys.stderr)
    print(f"{output}（{expected} 个内容页母页）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
