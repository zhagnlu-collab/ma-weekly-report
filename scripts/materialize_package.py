#!/usr/bin/env python3
"""将已核实的周报内容装配为完整交付包。

正式装配只接收由本机 Microsoft Office 通过电脑操控导出的三份 PDF。
LibreOffice 只可在显式预览模式下生成 preview/，预览结果不进入正式发布 PDF。
本脚本不负责检索、事实核验、Excel 重算或 Office 内容修改。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

import openpyxl


def pdf_page_count(path: Path) -> int:
    """统计 PDF 页数，不把 pypdf 当作硬依赖。

    v4.3.8：本机（WorkBuddy bundled Python 3.13）未安装 pypdf 但装有
    pymupdf(fitz)；原模块顶层 `from pypdf import PdfReader` 会让
    materialize_package.py 与 validate_package.py 在导入阶段直接失败，
    进而打断「日期进、成品出」的交付装配链。此处与
    record_office_export.py 保持同一套多路径兜底口径。
    """
    try:
        from pypdf import PdfReader  # type: ignore
        return len(PdfReader(str(path)).pages)
    except ImportError:
        pass
    try:
        import fitz  # type: ignore
        with fitz.open(str(path)) as doc:
            return doc.page_count
    except ImportError:
        pass
    data = path.read_bytes()
    return len(re.findall(rb'/Type\s*/Page(?!s)\b', data))


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def role(path: Path) -> str:
    name = path.name
    if name.startswith("并购重组市场周报") and path.suffix.lower() == ".pptx":
        return "market_ppt"
    if name.startswith("A股市场并购重组动态") and path.suffix.lower() == ".docx":
        return "dynamic_docx"
    match = re.match(r"^(\d+)、", name)
    if match and path.suffix.lower() in (".xlsx", ".xlsm"):
        return "xlsx_" + match.group(1)
    if match and match.group(1) == "5" and path.suffix.lower() == ".docx":
        return "weekly_docx"
    raise ValueError(f"无法识别文件角色: {name}")


REQUIRED = {f"xlsx_{i}" for i in range(14) if i != 5} | {
    "weekly_docx",
    "dynamic_docx",
    "market_ppt",
}
OFFICE_ROLES = {"weekly_docx", "dynamic_docx", "market_ppt"}


def inventory(directory: Path) -> dict[str, Path]:
    rows: dict[str, Path] = {}
    for path in directory.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in (".xlsx", ".xlsm", ".docx", ".pptx"):
            continue
        try:
            file_role = role(path)
        except ValueError:
            continue
        if file_role in rows:
            raise ValueError(f"角色重复: {file_role}: {rows[file_role].name}, {path.name}")
        rows[file_role] = path
    return rows


def validate_content(files: dict[str, Path]) -> None:
    missing = sorted(REQUIRED - set(files))
    extra = sorted(set(files) - REQUIRED)
    if missing:
        raise ValueError("缺少完整成果角色: " + ", ".join(missing))
    if extra:
        raise ValueError("存在未识别角色: " + ", ".join(extra))
    for file_role, path in files.items():
        with ZipFile(path) as archive:
            broken = archive.testzip()
            if broken:
                raise ValueError(f"{path.name} ZIP损坏: {broken}")
        if file_role.startswith("xlsx_"):
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=False)
            try:
                if not workbook.sheetnames:
                    raise ValueError(f"{path.name} 无工作表")
            finally:
                workbook.close()


def resolve_office_pdfs(files: dict[str, Path], source_dir: Path) -> dict[str, Path]:
    """核对 Office 电脑操控导出清单，并按原生文件名匹配三份 PDF。"""
    expected = {file_role: files[file_role].stem + ".pdf" for file_role in OFFICE_ROLES}
    resolved: dict[str, Path] = {}
    missing: list[str] = []
    invalid: list[str] = []
    manifest_path = source_dir / "office-export-manifest.json"
    if not manifest_path.is_file():
        raise ValueError("缺少 office-export-manifest.json，不能证明三份PDF由Microsoft Office导出")
    try:
        export_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ValueError(f"office-export-manifest.json 无法读取: {exc}") from exc
    if export_manifest.get("status") != "COMPLETE":
        raise ValueError("office-export-manifest.json 状态不是 COMPLETE")
    records = {
        item.get("file_role"): item
        for item in export_manifest.get("exports", [])
        if item.get("file_role")
    }
    if set(records) != set(OFFICE_ROLES) or len(export_manifest.get("exports", [])) != 3:
        raise ValueError("office-export-manifest.json 必须且只能登记周报、动态和PPT三份导出记录")
    for file_role, filename in expected.items():
        path = source_dir / filename
        record = records[file_role]
        if not path.is_file():
            missing.append(filename)
            continue
        if path.stat().st_size == 0:
            invalid.append(f"{filename}: 空文件")
            continue
        try:
            pages = pdf_page_count(path)
        except Exception as exc:
            invalid.append(f"{filename}: {type(exc).__name__}: {exc}")
            continue
        if pages == 0:
            invalid.append(f"{filename}: 0页")
            continue
        expected_app = "Microsoft PowerPoint" if file_role == "market_ppt" else "Microsoft Word"
        if record.get("pdf") != filename:
            invalid.append(f"{filename}: 导出清单PDF名不一致")
        if record.get("source") != files[file_role].name:
            invalid.append(f"{filename}: 导出清单源文件不一致")
        if record.get("source_sha256") != digest(files[file_role]):
            invalid.append(f"{filename}: 导出后源文件已变化")
        if record.get("pdf_sha256") != digest(path):
            invalid.append(f"{filename}: 导出后PDF已变化")
        if record.get("application") != expected_app:
            invalid.append(f"{filename}: 导出应用应为{expected_app}")
        if record.get("method") != "best_for_printing":
            invalid.append(f"{filename}: 未记录“最适合打印”导出")
        try:
            if int(record.get("native_pages")) != pages:
                invalid.append(f"{filename}: 原生页/幻灯片数 {record.get('native_pages')} != PDF页数 {pages}")
        except (TypeError, ValueError):
            invalid.append(f"{filename}: 导出清单缺少有效 native_pages")
        try:
            recorded_pdf_pages = int(record.get("pdf_pages"))
        except (TypeError, ValueError):
            invalid.append(f"{filename}: 导出清单缺少有效pdf_pages")
        else:
            if recorded_pdf_pages != pages:
                invalid.append(f"{filename}: 导出清单pdf_pages与PDF实际页数不一致")
        resolved[file_role] = path
    extras = sorted(
        path.name
        for path in source_dir.glob("*.pdf")
        if path.name not in set(expected.values())
    )
    if missing or invalid or extras:
        parts = []
        if missing:
            parts.append("缺少正式PDF: " + ", ".join(missing))
        if invalid:
            parts.append("PDF不可用: " + "; ".join(invalid))
        if extras:
            parts.append("存在额外PDF: " + ", ".join(extras))
        raise ValueError("；".join(parts))
    # Keep the manifest path discoverable for the caller.
    resolved["__manifest__"] = manifest_path
    return resolved


def convert_preview(soffice: Path, source: Path, outdir: Path, fontconfig: Path | None) -> Path:
    env = os.environ.copy()
    if fontconfig:
        env["FONTCONFIG_FILE"] = str(fontconfig.resolve())
    completed = subprocess.run(
        [str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(source)],
        capture_output=True,
        text=True,
        timeout=180,
        env=env,
    )
    pdf = outdir / (source.stem + ".pdf")
    if completed.returncode or not pdf.is_file() or pdf.stat().st_size == 0:
        raise RuntimeError(
            f"预览PDF生成失败: {source.name}; stdout={completed.stdout[-400:]} stderr={completed.stderr[-400:]}"
        )
    return pdf


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("content_dir", type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--pdf-source-dir", type=Path, help="Microsoft Office 电脑操控导出的三份发布 PDF 所在目录")
    parser.add_argument("--preview-soffice", type=Path, help="仅生成 preview/ 预览 PDF；不得用于正式交付")
    parser.add_argument("--fontconfig", type=Path, help="仅用于 LibreOffice 预览")
    args = parser.parse_args()

    if bool(args.pdf_source_dir) == bool(args.preview_soffice):
        parser.error("必须且只能指定 --pdf-source-dir 或 --preview-soffice")

    content = inventory(args.content_dir)
    validate_content(content)
    office_pdfs: dict[str, Path] = {}
    if args.pdf_source_dir:
        office_pdfs = resolve_office_pdfs(content, args.pdf_source_dir)

    out = args.out.resolve()
    if any(out in path.resolve().parents for path in content.values()) or out in {
        path.resolve().parent for path in content.values()
    }:
        parser.error("输出目录不得与内容输入目录相同")
    if out.exists():
        parser.error(f"输出目录已存在，未覆盖: {out}")
    out.mkdir(parents=True)

    copied: list[dict] = []
    for file_role, source in sorted(content.items()):
        destination = out / source.name
        shutil.copy2(source, destination)
        copied.append(
            {
                "role": file_role,
                "source": str(source.resolve()),
                "dest": str(destination),
                "sha256": digest(destination),
            }
        )

    pdfs: list[dict] = []
    for file_role, source in sorted(office_pdfs.items()):
        if file_role == "__manifest__":
            continue
        destination = out / source.name
        shutil.copy2(source, destination)
        pdfs.append(
            {
                "role": file_role,
                "source": str(source.resolve()),
                "dest": str(destination),
                "sha256": digest(destination),
                "pages": len(PdfReader(destination).pages),
                "origin": "microsoft-office-computer-use",
            }
        )

    previews: list[str] = []
    if args.preview_soffice:
        preview_dir = out / "preview"
        preview_dir.mkdir()
        for file_role in sorted(OFFICE_ROLES):
            previews.append(
                str(convert_preview(args.preview_soffice, content[file_role], preview_dir, args.fontconfig))
            )

    export_manifest_copy = None
    if args.pdf_source_dir:
        export_manifest_copy = out / "office-export-manifest.json"
        shutil.copy2(office_pdfs["__manifest__"], export_manifest_copy)

    status = "PASS" if len(pdfs) == 3 else "PREVIEW_ONLY"
    result = {
        "status": status,
        "content_dir": str(args.content_dir.resolve()),
        "out": str(out),
        "native_files": len(copied),
        "published_pdf_files": len(pdfs),
        "preview_pdf_files": len(previews),
        "errors": [],
        "files": copied,
        "pdfs": pdfs,
        "office_export_manifest": str(export_manifest_copy) if export_manifest_copy else None,
        "previews": previews,
        "scope_note": (
            "仅装配已核实内容；正式PDF由Microsoft Office电脑操控导出。"
            if status == "PASS"
            else "仅完成原生文件装配和LibreOffice预览；preview目录不是正式发布PDF。"
        ),
    }
    (out / "materialize-manifest.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 2


if __name__ == "__main__":
    sys.exit(main())
