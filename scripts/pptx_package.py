"""检查和保守修复 PPTX 包；不依赖 PowerPoint 或 officecli。

package_problems(path) 返回可读的问题列表（空列表表示通过）。
repair_package(path) 返回修复说明列表；无法修复时抛出 PackageError，文件保持不变。
调用方应传入输出副本，不能传入需要保留的模板原件。

仅在页内动画/构建 spid 悬空时移除整页 p:timing 及 p:bldLst：通常
p:bldLst 已包含在 p:timing 内，也兼顾独立的构建列表。合法页面不重写。
其余仅更新 app.xml 的 Slides/Notes；缺失关系和重复 shape id 不自动修复。
"""
import os
import posixpath
import stat
import tempfile
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS = {"p": P}
APP_PART = "docProps/app.xml"


class PackageError(ValueError):
    """包存在无法保守修复的问题。"""


def _xml(data):
    return etree.fromstring(
        data, etree.XMLParser(resolve_entities=False, no_network=True)
    )


def _parts(names, folder):
    return sorted(
        name for name in names
        if posixpath.dirname(name) == folder and name.endswith(".xml")
    )


def _relationship_source(name):
    if name == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(name)
    if posixpath.basename(directory) != "_rels" or not filename.endswith(".rels"):
        return None
    return posixpath.join(posixpath.dirname(directory), filename[:-5])


def _target_part(source, target):
    # OPC 关系 URI 相对于源部件，而不是 .rels 文件。片段不属于 ZIP 路径。
    uri = urlsplit(target)
    path = unquote(uri.path)
    if not path:
        return source
    if path.startswith("/"):
        return posixpath.normpath(path.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), path))


def _dangling_spids(root):
    ids = set(root.xpath(".//p:cNvPr/@id", namespaces=NS))
    refs = root.xpath(".//p:timing//@spid | .//p:bldLst//@spid", namespaces=NS)
    return [spid for spid in refs if spid not in ids]


def _problems(package):
    names = set(package.namelist())
    problems = []
    for name, count in Counter(package.namelist()).items():
        if count > 1:
            problems.append(f"{name}: ZIP 内存在 {count} 个同名条目")

    roots = {}
    for name in sorted(names):
        if name.endswith((".xml", ".rels")):
            try:
                roots[name] = _xml(package.read(name))
            except etree.XMLSyntaxError as exc:
                problems.append(f"{name}: XML 无法解析: {exc}")

    relationships = {}
    for name, root in roots.items():
        source = _relationship_source(name)
        if source is None:
            continue
        if source and source not in names:
            problems.append(f"{name}: 关系源部件不存在: {source}")
        rels = {}
        for rel in root.findall(f"{{{REL}}}Relationship"):
            rid = rel.get("Id", "")
            if not rid or rid in rels:
                problems.append(f"{name}: 关系 Id 为空或重复: {rid!r}")
            rels[rid] = rel
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            try:
                resolved = _target_part(source, target)
            except ValueError:
                resolved = ""
            if not target or resolved not in names:
                problems.append(f"{name}: 内部关系 {rid} 的目标不存在: {target!r}")
        relationships[source] = rels

    # 同时检查 XML 中 r:id/r:embed 等引用，避免目标存在但关系条目缺失。
    for name, root in roots.items():
        if not name.endswith(".xml"):
            continue
        rels = relationships.get(name, {})
        for element in root.iter():
            for attr, rid in element.attrib.items():
                if attr.startswith(f"{{{R}}}") and rid not in rels:
                    problems.append(f"{name}: 引用的关系不存在: {rid!r}")

    animation_parts = sum((_parts(names, folder) for folder in (
        "ppt/slides", "ppt/slideLayouts", "ppt/slideMasters")), [])
    for name in animation_parts:
        root = roots.get(name)
        if root is None:
            continue
        ids = root.xpath(".//p:cNvPr/@id", namespaces=NS)
        for shape_id, count in Counter(ids).items():
            if count > 1:
                problems.append(f"{name}: 页内 shape id {shape_id} 重复 {count} 次")
        for timing in root.iter(f"{{{P}}}timing"):
            time_ids = timing.xpath(".//p:cTn/@id", namespaces=NS)
            if len(time_ids) != len(set(time_ids)):
                problems.append(f"{name}: 动画时间节点 cTn id 重复")
            for ref in timing.xpath(".//p:tn/@val", namespaces=NS):
                if ref not in time_ids:
                    problems.append(f"{name}: 动画时间依赖 tn={ref} 不存在")
        dangling = _dangling_spids(root)
        if dangling:
            problems.append(
                f"{name}: 动画/构建存在 {len(dangling)} 个悬空 spid: "
                + ", ".join(sorted(set(dangling)))
            )

    app = roots.get(APP_PART)
    if APP_PART not in names:
        problems.append(f"缺少 {APP_PART}，无法核对 Slides/Notes")
    elif app is not None:
        for tag, folder in (("Slides", "ppt/slides"), ("Notes", "ppt/notesSlides")):
            expected = len(_parts(names, folder))
            node = app.find(f"{{{APP}}}{tag}")
            value = node.text if node is not None else None
            if value != str(expected):
                problems.append(f"{APP_PART}: {tag}={value!r}，实际应为 {expected}")
    return problems


def package_problems(path):
    """返回内部关系、页内 shape id、动画/构建引用及页数元数据问题。"""
    try:
        with ZipFile(path) as package:
            return _problems(package)
    except (OSError, BadZipFile, RuntimeError, NotImplementedError) as exc:
        return [f"无法读取 PPTX 包 {path}: {exc}"]


def _serialize(root):
    return etree.tostring(root.getroottree(), encoding="UTF-8", xml_declaration=True)


def repair_package(path, *, repair_animations=True):
    """原子修复输出副本并复查；任何检查/写入失败都不替换原文件。

    repair_animations=False 时只修页数元数据，动画错误会阻断。
    repair_animations=True（默认）时清除含悬空 spid 页面的 timing/bldLst，并更新 app.xml。
    没有修改时也执行检查；已通过且无需修复的文件不会重写。
    """
    path = Path(path).resolve()
    staging = None
    repairs = []
    try:
        with ZipFile(path) as package:
            names = set(package.namelist())
            replacements = {}
            for name in _parts(names, "ppt/slides") if repair_animations else []:
                root = _xml(package.read(name))
                dangling = _dangling_spids(root)
                if not dangling:
                    continue
                for node in root.xpath(".//p:timing | .//p:bldLst", namespaces=NS):
                    node.getparent().remove(node)
                replacements[name] = _serialize(root)
                repairs.append(
                    f"{name}: 移除整页 p:timing 及关联 p:bldLst"
                    f"（{len(dangling)} 个悬空 spid）"
                )

            if APP_PART in names:
                app = _xml(package.read(APP_PART))
                for tag, folder in (("Slides", "ppt/slides"), ("Notes", "ppt/notesSlides")):
                    expected = str(len(_parts(names, folder)))
                    node = app.find(f"{{{APP}}}{tag}")
                    if node is None:
                        node = etree.SubElement(app, f"{{{APP}}}{tag}")
                    if node.text != expected:
                        node.text = expected
                        replacements[APP_PART] = _serialize(app)
                        repairs.append(f"{APP_PART}: 更新 {tag}={expected}")

            if not replacements:
                problems = _problems(package)
                if problems:
                    raise PackageError("\n".join(problems))
                return []

            fd, temporary = tempfile.mkstemp(
                prefix=f".{path.stem}.repair-", suffix=".pptx", dir=path.parent
            )
            staging = Path(temporary)
            os.close(fd)
            with ZipFile(staging, "w") as repaired:
                repaired.comment = package.comment
                for info in package.infolist():
                    data = replacements.get(info.filename)
                    repaired.writestr(info, package.read(info) if data is None else data)

        problems = package_problems(staging)
        if problems:
            raise PackageError("\n".join(problems))
        os.chmod(staging, stat.S_IMODE(path.stat().st_mode))
        os.replace(staging, path)
        return repairs
    except (OSError, BadZipFile, etree.XMLSyntaxError, RuntimeError, NotImplementedError) as exc:
        raise PackageError(f"修复 PPTX 包失败: {exc}") from exc
    finally:
        if staging is not None:
            staging.unlink(missing_ok=True)
