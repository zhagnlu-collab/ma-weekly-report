#!/usr/bin/env python3
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from materialize_package import pdf_page_count, resolve_office_pdfs


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def pdf(path: Path, pages: int = 1) -> None:
    """生成指定页数的测试 PDF；优先 pypdf，缺失时回退 pymupdf。"""
    try:
        from pypdf import PdfWriter  # type: ignore
    except ImportError:
        import fitz  # type: ignore
        doc = fitz.open()
        for _ in range(pages):
            doc.new_page(width=72, height=72)
        doc.save(str(path))
        doc.close()
        return
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=72, height=72)
    with path.open("wb") as stream:
        writer.write(stream)


class OfficePdfGateTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.sources = {
            "weekly_docx": self.root / "5、周报.docx",
            "dynamic_docx": self.root / "A股市场并购重组动态.docx",
            "market_ppt": self.root / "并购重组市场周报.pptx",
        }
        for path in self.sources.values():
            path.write_bytes(b"native-office-placeholder")
        self.pdf_dir = self.root / "pdf"
        self.pdf_dir.mkdir()
        self.records = []
        for role, source in self.sources.items():
            target = self.pdf_dir / f"{source.stem}.pdf"
            pdf(target, 2)
            self.records.append(
                {
                    "file_role": role,
                    "source": source.name,
                    "source_sha256": sha(source),
                    "pdf": target.name,
                    "pdf_sha256": sha(target),
                    "pdf_pages": 2,
                    "native_pages": 2,
                    "application": "Microsoft PowerPoint" if role == "market_ppt" else "Microsoft Word",
                    "method": "best_for_printing",
                }
            )
        self.manifest = self.pdf_dir / "office-export-manifest.json"
        self.write_manifest()

    def write_manifest(self):
        self.manifest.write_text(
            json.dumps({"status": "COMPLETE", "exports": self.records}, ensure_ascii=False),
            encoding="utf-8",
        )

    def tearDown(self):
        self.temp.cleanup()

    def test_exact_office_exports_pass(self):
        resolved = resolve_office_pdfs(self.sources, self.pdf_dir)
        self.assertEqual(set(resolved), set(self.sources) | {"__manifest__"})

    def test_missing_formal_pdf_is_rejected(self):
        (self.pdf_dir / "A股市场并购重组动态.pdf").unlink()
        with self.assertRaisesRegex(ValueError, "缺少正式PDF"):
            resolve_office_pdfs(self.sources, self.pdf_dir)

    def test_native_and_pdf_page_mismatch_is_rejected(self):
        self.records[0]["native_pages"] = 3
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, "原生页/幻灯片数"):
            resolve_office_pdfs(self.sources, self.pdf_dir)

    def test_source_change_after_export_is_rejected(self):
        self.sources["weekly_docx"].write_bytes(b"changed-after-export")
        with self.assertRaisesRegex(ValueError, "导出后源文件已变化"):
            resolve_office_pdfs(self.sources, self.pdf_dir)

    def test_pdf_page_count_works_without_pypdf(self):
        target = self.pdf_dir / "pagecount.pdf"
        pdf(target, 3)
        self.assertEqual(pdf_page_count(target), 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
