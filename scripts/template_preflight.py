"""Read-only selected-page preflight: classify inherited defects and proposed capacity.

Usage: python template_preflight.py --template template.pptx --map map.json --output report.json
No defect is exempted here; clean/remove/replace slots or select another source page before filling.
"""
import argparse
import json
from pathlib import Path

from power_pptx import Presentation
from template_fill import MapError, build_ops, scan_template
from validate_pptx import _inches, check_slide, collect_baseline


def inspect(template, mapping, map_dir):
    prs = Presentation(str(template))
    pages = scan_template(str(template))
    baseline = collect_baseline(prs)
    report = {"template": str(Path(template).resolve()), "pages": [], "capacity_ready": True}
    for number, selected in enumerate(mapping["slides"], 1):
        capacity = []
        try:
            build_ops([selected], pages, map_dir)
        except (MapError, ValueError, KeyError) as exc:
            capacity.append(str(exc))
        inherited, soft, info = [], [], []
        source = selected.get("template_slide")
        if isinstance(source, int) and 1 <= source <= len(prs.slides):
            check_slide(prs.slides[source-1], source, None, inherited, soft,
                        _inches(prs.slide_width), _inches(prs.slide_height), baseline, info,
                        check_formulas=False)
        row = {"output_slide": number, "source_slide": source,
               "proposed_capacity_errors": capacity,
               "template_content_defects": inherited,
               "template_geometry_review": info, "soft": soft,
               "declared_removals": selected.get("remove", []),
               "resolution_order": ["重组真实语义项，保留论据、条件、机制与边界", "换容量更大的原页、拆页或增页，不缩字号",
                                    "全遵循保护全部槽位及几何，禁止静默删除卡片或连线", "半自主三层清理正文画布，预览后注入自主布局",
                                    "保留上次成功输出；返回规划，不跳过Hard"]}
        report["pages"].append(row)
        # Template defects are diagnostics before declared removals; staging
        # final checks remain the publication gate after those operations.
        if capacity:
            report["capacity_ready"] = False
    return report


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--template", required=True)
    ap.add_argument("--map", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    mapping_path = Path(args.map).resolve()
    report = inspect(args.template, json.loads(mapping_path.read_text()), mapping_path.parent)
    Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report["capacity_ready"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
