"""officecli 命令行封装：定位二进制、执行命令、批量操作、落盘。

被 formula.py、render_preview.py、template_fill.py 共用。
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


class OfficeCliError(RuntimeError):
    pass


def find_officecli():
    """先查 PATH，再查官方安装器的默认目录；都没有则报错并给出安装提示。"""
    exe = shutil.which("officecli")
    if exe:
        return exe
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        cand = Path(base) / "OfficeCLI" / "officecli.exe" if base else None
    else:
        cand = Path.home() / ".local" / "bin" / "officecli"
    if cand and cand.exists():
        return str(cand)
    raise OfficeCliError(
        "找不到 officecli。先 pip install -r requirements.txt，再执行 python -m officecli install；"
        "下载失败时手动下载对应平台的二进制，放到 %LOCALAPPDATA%\\OfficeCLI\\officecli.exe（Windows）"
        "或 ~/.local/bin/officecli。"
    )


def run(args, check=True, timeout=180):
    """执行 officecli，返回 CompletedProcess（文本模式，UTF-8）。"""
    cmd = [find_officecli()] + [str(a) for a in args]
    env = dict(os.environ, PYTHONUTF8="1")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=timeout, env=env)
    if check and r.returncode != 0:
        raise OfficeCliError(f"officecli {' '.join(map(str, args))} 失败(退出码 {r.returncode}):\n{r.stdout}{r.stderr}")
    return r


def run_json(args, timeout=180):
    """执行带 --json 的命令并解析结果；success=false 时抛错。"""
    r = run(list(args) + ["--json"], check=False, timeout=timeout)
    try:
        data = json.loads(r.stdout)
    except json.JSONDecodeError:
        raise OfficeCliError(f"officecli 输出不是 JSON:\n{r.stdout}{r.stderr}")
    if not data.get("success", False):
        raise OfficeCliError(f"officecli 报错: {json.dumps(data.get('error', data), ensure_ascii=False)}")
    return data


def batch(pptx, ops, timeout=300):
    """用 batch 原子执行一组操作（失败整批回滚）。ops 是 dict 列表，见 officecli batch 说明。"""
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump(ops, f, ensure_ascii=False)
        tmp = f.name
    try:
        return run_json(["batch", pptx, "--input", tmp], timeout=timeout)
    finally:
        os.unlink(tmp)


def close(pptx):
    """officecli 常驻进程会延迟落盘；其他程序读文件前必须调用。"""
    run(["close", pptx], check=False)


def slide_count(pptx):
    return int(run_json(["view", pptx, "stats"])["data"]["slides"])


def screenshot(pptx, out_png, start=None, end=None, render="html", width=None, grid=None):
    args = ["view", pptx, "screenshot", "-o", out_png, "--render", render]
    if start:
        args += ["--start", start, "--end", end or start]
    if width:
        args += ["--screenshot-width", width]
    if grid:
        args += ["--grid", grid]
    return run(args)


if __name__ == "__main__":
    print(find_officecli())
    sys.exit(0)
