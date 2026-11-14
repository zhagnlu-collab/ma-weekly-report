#!/usr/bin/env python3
"""登记一份由 Microsoft Office 电脑操控导出的发布 PDF。

登记时复算 PDF 页数，并要求其等于电脑操控读到的 Word 总页数或 PPT 幻灯片数。
三份记录齐全后，manifest 才标记 COMPLETE，供 materialize_package.py 硬门禁使用。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

def pdf_page_count(path: Path) -> int:
    """优先 pymupdf（主链路渲染器），回退 pypdf，最后用 PDF 对象计数兜底。

    主入口用同一个解释器调用所有脚本，因此这里不能硬依赖只在另一个
    Python 环境里存在的 pypdf。
    """
    try:
        import pymupdf
        with pymupdf.open(path) as doc:
            return len(doc)
    except ImportError:
        pass
    try:
        from pypdf import PdfReader
        return len(PdfReader(path).pages)
    except ImportError:
        pass
    import re
    data = path.read_bytes()
    counts = [int(n) for n in re.findall(rb'/Type\s*/Pages[^>]*?/Count\s+(\d+)', data, re.S)]
    return max(counts) if counts else len(re.findall(rb'/Type\s*/Page\b', data))


ROLE_SUFFIX = {
    "weekly_docx": ".docx",
    "dynamic_docx": ".docx",
    "market_ppt": ".pptx",
}
ROLE_APP = {
    "weekly_docx": "Microsoft Word",
    "dynamic_docx": "Microsoft Word",
    "market_ppt": "Microsoft PowerPoint",
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pdf-dir", required=True, type=Path)
    parser.add_argument("--file-role", required=True, choices=sorted(ROLE_SUFFIX))
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--native-pages", required=True, type=int, help="Word总页数或PPT幻灯片数")
    parser.add_argument("--reset", action="store_true", help="清空旧清单后登记")
    args = parser.parse_args()

    role = args.file_role
    source = args.source.resolve()
    pdf = args.pdf.resolve()
    if not source.is_file() or source.suffix.lower() != ROLE_SUFFIX[role]:
        raise ValueError(f"源文件类型与角色不符: {source}")
    if not pdf.is_file() or pdf.suffix.lower() != ".pdf":
        raise ValueError(f"PDF不存在或类型错误: {pdf}")
    if source.stem != pdf.stem:
        raise ValueError(f"PDF必须与源文件同名: {source.stem}.pdf；实际为{pdf.name}")
    pdf_pages = pdf_page_count(pdf)
    if pdf_pages <= 0:
        raise ValueError("PDF为0页")
    if args.native_pages <= 0:
        raise ValueError("--native-pages 必须大于0")
    if pdf_pages != args.native_pages:
        raise ValueError(
            f"PDF页数{pdf_pages}不等于{ROLE_APP[role]}原生页数/幻灯片数{args.native_pages}，必须重新导出"
        )

    args.pdf_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = args.pdf_dir / "office-export-manifest.json"
    manifest = {"version": "1.0", "exports": []}
    if manifest_path.is_file() and not args.reset:
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception as exc:
            raise ValueError(f"既有导出清单无法读取: {exc}") from exc
    records = {
        item.get("file_role"): item
        for item in manifest.get("exports", [])
        if item.get("file_role") in ROLE_SUFFIX
    }
    records[role] = {
        "file_role": role,
        "source": source.name,
        "source_sha256": sha256(source),
        "pdf": pdf.name,
        "pdf_sha256": sha256(pdf),
        "pdf_pages": pdf_pages,
        "native_pages": args.native_pages,
        "application": ROLE_APP[role],
        "method": "best_for_printing",
        "recorded_by": "computer-use",
        "recorded_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
    }
    exports = [records[key] for key in sorted(records)]
    manifest = {
        "version": "1.0",
        "status": "COMPLETE" if set(records) == set(ROLE_SUFFIX) else "INCOMPLETE",
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "exports": exports,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "status": manifest["status"],
                "recorded": role,
                "native_pages": args.native_pages,
                "pdf_pages": pdf_pages,
                "manifest": str(manifest_path),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(json.dumps({"status": "FAIL", "error": str(exc)}, ensure_ascii=False))
        sys.exit(1)
