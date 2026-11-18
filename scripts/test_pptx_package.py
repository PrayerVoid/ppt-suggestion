"""运行: python -I -m unittest discover -s scripts -p 'test_pptx_package.py' -v"""
import contextlib
import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree

import pptx_package
from pptx_package import APP, P, R, REL, PackageError, package_problems, repair_package


def _slide(shape_id="1", target="1", standalone_build=False):
    build = f'<p:bldLst><p:bldP spid="{target}" grpId="0"/></p:bldLst>'
    return (
        f'<p:sld xmlns:p="{P}"><p:cSld><p:spTree><p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{shape_id}" name="shape"/></p:nvSpPr></p:sp></p:spTree>'
        f'</p:cSld><p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst>'
        f'<p:set><p:cBhvr><p:tgtEl><p:spTgt spid="{target}"/></p:tgtEl>'
        f'</p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'{"" if standalone_build else build}</p:timing>'
        f'{build if standalone_build else ""}</p:sld>'
    ).encode()


def _fixture(path, *, bad=False, missing=False, stale=False, standalone_build=False):
    parts = {
        "_rels/.rels": (
            f'<Relationships xmlns="{REL}"><Relationship Id="rId1" '
            f'Type="{R}/officeDocument" Target="ppt/presentation.xml"/>'
            '</Relationships>'
        ).encode(),
        "ppt/presentation.xml": (
            f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst>'
            '<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>'
            '</p:sldIdLst></p:presentation>'
        ).encode(),
        "ppt/_rels/presentation.xml.rels": (
            f'<Relationships xmlns="{REL}"><Relationship Id="rId1" '
            f'Type="{R}/slide" Target="slides/{"absent" if missing else "slide1"}.xml"/>'
            f'<Relationship Id="rId2" Type="{R}/slide" Target="/ppt/slides/slide2.xml"/>'
            f'<Relationship Id="external" Type="{R}/hyperlink" '
            'Target="https://example.org/not-in-package" TargetMode="External"/>'
            '</Relationships>'
        ).encode(),
        "ppt/slides/slide1.xml": _slide(
            target="99" if bad else "1", standalone_build=standalone_build
        ),
        "ppt/slides/slide2.xml": _slide(),
        "ppt/notesSlides/notesSlide1.xml": f'<p:notes xmlns:p="{P}"/>'.encode(),
        "docProps/app.xml": (
            f'<Properties xmlns="{APP}"><Application>fixture</Application>'
            f'<Slides>{24 if stale else 2}</Slides><Notes>{24 if stale else 1}</Notes>'
            '</Properties>'
        ).encode(),
        "ppt/media/unchanged.bin": b"\x00\xffopaque payload\n",
    }
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as package:
        package.comment = b"preserve archive comment"
        for name, data in parts.items():
            info = ZipInfo(name, date_time=(2024, 2, 3, 4, 5, 6))
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o100640 << 16
            info.comment = b"preserve entry comment"
            package.writestr(info, data)
    return parts


def _replace_part(path, name, data):
    with ZipFile(path) as package:
        comment = package.comment
        parts = [(info, package.read(info)) for info in package.infolist()]
    with ZipFile(path, "w") as package:
        package.comment = comment
        for info, old in parts:
            package.writestr(info, data if info.filename == name else old)


class PackageTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name).resolve()
        self.path = self.root / "output.pptx"

    def test_bad_animation_detected_repaired_and_other_zip_contents_preserved(self):
        before = _fixture(self.path, bad=True, stale=True)
        problems = package_problems(self.path)
        self.assertTrue(any("悬空 spid" in problem for problem in problems))
        self.assertTrue(any("Slides" in problem for problem in problems))
        with ZipFile(self.path) as package:
            metadata = [(i.filename, i.date_time, i.compress_type, i.external_attr, i.comment)
                        for i in package.infolist()]
        repairs = repair_package(self.path)
        self.assertTrue(any("p:timing" in repair for repair in repairs))
        self.assertEqual(package_problems(self.path), [])
        with ZipFile(self.path) as package:
            root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
            self.assertEqual(root.findall(f"{{{P}}}timing"), [])
            self.assertEqual(root.findall(f".//{{{P}}}bldLst"), [])
            self.assertEqual(root.find(f".//{{{P}}}cNvPr").get("id"), "1")
            for name, data in before.items():
                if name not in ("ppt/slides/slide1.xml", "docProps/app.xml"):
                    self.assertEqual(package.read(name), data, name)
            app = etree.fromstring(package.read("docProps/app.xml"))
            self.assertEqual(app.find(f"{{{APP}}}Slides").text, "2")
            self.assertEqual(app.find(f"{{{APP}}}Notes").text, "1")
            self.assertEqual(package.comment, b"preserve archive comment")
            self.assertEqual(
                [(i.filename, i.date_time, i.compress_type, i.external_attr, i.comment)
                 for i in package.infolist()], metadata
            )
        self.assertEqual(list(self.root.iterdir()), [self.path])
        after = self.path.read_bytes()
        self.assertEqual(repair_package(self.path), [])
        self.assertEqual(self.path.read_bytes(), after)

    def test_valid_animations_remain_byte_identical(self):
        _fixture(self.path)
        before = self.path.read_bytes()
        self.assertEqual(package_problems(self.path), [])
        self.assertEqual(repair_package(self.path), [])
        self.assertEqual(self.path.read_bytes(), before)

    def test_standalone_build_list_is_removed_with_bad_timing(self):
        _fixture(self.path, bad=True, standalone_build=True)
        repair_package(self.path)
        with ZipFile(self.path) as package:
            root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
        self.assertEqual(root.findall(f"{{{P}}}timing"), [])
        self.assertEqual(root.findall(f"{{{P}}}bldLst"), [])
        self.assertEqual(package_problems(self.path), [])

    def test_bad_build_reference_alone_triggers_removal(self):
        parts = _fixture(self.path)
        slide = parts["ppt/slides/slide1.xml"].replace(b'<p:bldP spid="1"', b'<p:bldP spid="99"')
        _replace_part(self.path, "ppt/slides/slide1.xml", slide)
        self.assertTrue(any("悬空 spid" in p for p in package_problems(self.path)))
        repair_package(self.path)
        self.assertEqual(package_problems(self.path), [])

    def test_missing_internal_relationship_blocks_repair_and_preserves_input(self):
        _fixture(self.path, bad=True, stale=True, missing=True)
        before = self.path.read_bytes()
        self.assertTrue(any("目标不存在" in p for p in package_problems(self.path)))
        with self.assertRaisesRegex(PackageError, "目标不存在"):
            repair_package(self.path)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(list(self.root.iterdir()), [self.path])

    def test_missing_relationship_id_is_reported(self):
        parts = _fixture(self.path)
        data = parts["ppt/presentation.xml"].replace(b'r:id="rId2"', b'r:id="missing"')
        _replace_part(self.path, "ppt/presentation.xml", data)
        self.assertTrue(any("引用的关系不存在" in p for p in package_problems(self.path)))
        with self.assertRaises(PackageError):
            repair_package(self.path)

    def test_percent_encoded_relative_relationship_target_exists(self):
        parts = _fixture(self.path)
        data = parts["ppt/_rels/presentation.xml.rels"].replace(
            b'Target="slides/slide1.xml"', b'Target="slides/../slides/%73lide1.xml#fragment"'
        )
        _replace_part(self.path, "ppt/_rels/presentation.xml.rels", data)
        self.assertEqual(package_problems(self.path), [])

    def test_duplicate_shape_ids_are_page_local_and_block_repair(self):
        parts = _fixture(self.path)
        # 两页各自使用 id=1 合法；同页重复才应阻断。
        self.assertEqual(package_problems(self.path), [])
        duplicate = parts["ppt/slides/slide1.xml"].replace(
            b'</p:spTree>', b'<p:sp><p:nvSpPr><p:cNvPr id="1" name="duplicate"/>'
            b'</p:nvSpPr></p:sp></p:spTree>'
        )
        _replace_part(self.path, "ppt/slides/slide1.xml", duplicate)
        before = self.path.read_bytes()
        self.assertTrue(any("shape id 1 重复" in p for p in package_problems(self.path)))
        with self.assertRaisesRegex(PackageError, "shape id 1 重复"):
            repair_package(self.path)
        self.assertEqual(self.path.read_bytes(), before)

    def test_replace_failure_preserves_input_and_cleans_temporary(self):
        _fixture(self.path, bad=True)
        before = self.path.read_bytes()
        with patch.object(pptx_package.os, "replace", side_effect=OSError("replace failed")):
            with self.assertRaisesRegex(PackageError, "replace failed"):
                repair_package(self.path)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(list(self.root.iterdir()), [self.path])

    def test_unreadable_or_malformed_package_is_reported(self):
        self.path.write_bytes(b"not a zip")
        self.assertTrue(package_problems(self.path))
        with self.assertRaises(PackageError):
            repair_package(self.path)
        _fixture(self.path)
        _replace_part(self.path, "ppt/slides/slide1.xml", b"<broken")
        self.assertTrue(any("XML 无法解析" in p for p in package_problems(self.path)))
        with self.assertRaises(PackageError):
            repair_package(self.path)


class MotherPublicationTests(unittest.TestCase):
    def setUp(self):
        import template_mother
        self.mother = template_mother
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name).resolve()
        self.template = self.root / "template.pptx"
        self.output = self.root / "output.pptx"
        self.manifest = self.root / "mothers.json"
        self.manifest.write_text(json.dumps({"mothers": [{"template_slide": 1}]}))
        _fixture(self.template, bad=True, stale=True)
        self.original = self.template.read_bytes()
        self.output.write_bytes(b"previous successful output")
        self.argv = ["template_mother.py", "--template", str(self.template),
                     "--manifest", str(self.manifest), "--output", str(self.output)]

    def _run(self, *, batch_error=None, missing=False, count=1, close_error=None):
        def batch(path, ops):
            self.assertNotEqual(Path(path), self.output)
            self.assertEqual(Path(path).parent, self.output.parent)
            if batch_error:
                raise batch_error
            if missing:
                _fixture(path, bad=True, missing=True, stale=True)

        with contextlib.ExitStack() as stack:
            stack.enter_context(patch("sys.argv", self.argv))
            stack.enter_context(contextlib.redirect_stdout(io.StringIO()))
            self.stderr = io.StringIO()
            stack.enter_context(contextlib.redirect_stderr(self.stderr))
            stack.enter_context(patch.object(self.mother, "scan_template", return_value=[{}]))
            # These lightweight OPC fixtures isolate publication, not the clone parser.
            stack.enter_context(patch('slide_clone.clone_selected',
                side_effect=lambda template, selected, output: shutil.copyfile(template, output)))
            stack.enter_context(patch.object(self.mother.ocli, "batch", side_effect=batch))
            closed = stack.enter_context(patch.object(self.mother.ocli, "close", side_effect=close_error))
            presentation = stack.enter_context(patch.object(
                self.mother, "Presentation", return_value=SimpleNamespace(slides=[None] * count)
            ))
            result = self.mother.main()
        self.assertEqual(self.template.read_bytes(), self.original)
        self.assertEqual(set(self.root.iterdir()), {self.template, self.output, self.manifest})
        return result, closed, presentation

    def test_publish_repaired_staging_only_after_all_checks(self):
        result, closed, presentation = self._run()
        self.assertEqual(result, 0, self.stderr.getvalue())
        self.assertEqual(package_problems(self.output), [])
        self.assertEqual(closed.call_count, 2)
        self.assertEqual(closed.call_args_list[-1].args, (str(self.output),))
        self.assertNotEqual(presentation.call_args.args[0], str(self.output))

    def test_batch_failure_keeps_existing_output_and_closes_staging(self):
        result, closed, presentation = self._run(batch_error=self.mother.ocli.OfficeCliError("batch failed"))
        self.assertEqual(result, 1)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")
        closed.assert_called_once()
        presentation.assert_not_called()

    def test_missing_relationship_blocks_publication(self):
        result, closed, presentation = self._run(missing=True)
        self.assertEqual(result, 1)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")
        closed.assert_called_once()
        presentation.assert_not_called()

    def test_slide_count_failure_keeps_existing_output(self):
        result, closed, _ = self._run(count=2)
        self.assertEqual(result, 1)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")
        closed.assert_called_once()

    def test_close_failure_keeps_existing_output_and_cleans_staging(self):
        result, closed, presentation = self._run(close_error=RuntimeError("close failed"))
        self.assertEqual(result, 1)
        self.assertEqual(self.output.read_bytes(), b"previous successful output")
        closed.assert_called_once()
        presentation.assert_not_called()


if __name__ == "__main__":
    unittest.main()
