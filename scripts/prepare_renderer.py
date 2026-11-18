"""Copy a complete autonomous renderer runtime, without overwriting local edits by default."""
import argparse
import shutil
from pathlib import Path

FILES = ("render_pptx.py", "text_measure.py", "deck_check.py", "layout_capacity.py",
         "speaker_notes.py", "pptx_package.py")


def prepare(output, *, force=False):
    source = Path(__file__).resolve().parent
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    conflicts = [name for name in FILES if (output / name).exists()
                 and (output / name).read_bytes() != (source / name).read_bytes()]
    if conflicts and not force:
        raise ValueError("目标含定制修改，未覆盖：" + ", ".join(conflicts) + "；先审阅差异，确需替换才使用 --force")
    for name in FILES:
        shutil.copyfile(source / name, output / name)
    return output


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("output")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    try:
        print(prepare(args.output, force=args.force))
    except (OSError, ValueError) as exc:
        ap.error(str(exc))


if __name__ == "__main__":
    main()
