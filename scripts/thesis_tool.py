#!/usr/bin/env python3
"""Utilities for repeatable thesis DOCX/PDF production.

The script intentionally handles deterministic chores only:
version copying, placeholder filling, DOCX text/caption extraction, Microsoft Word PDF export,
PDF page mapping, and front-matter page-number patching. Prose edits and fragile layout moves
still belong to the agent, using OOXML carefully.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

try:
    from lxml import etree
except Exception as exc:  # pragma: no cover
    raise SystemExit("lxml is required for DOCX OOXML operations") from exc


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def paragraph_text(p: etree._Element) -> str:
    parts: list[str] = []
    for node in p.iter():
        tag = node.tag
        if tag == f"{W}t" and node.text:
            parts.append(node.text)
        elif tag == f"{W}tab":
            parts.append("\t")
    return "".join(parts)


def clear_paragraph_content(p: etree._Element) -> None:
    for child in list(p):
        if child.tag != f"{W}pPr":
            p.remove(child)


def append_run(p: etree._Element, text: str) -> None:
    r = etree.SubElement(p, f"{W}r")
    segments = text.split("\t")
    for i, segment in enumerate(segments):
        if segment:
            t = etree.SubElement(r, f"{W}t")
            if segment.startswith(" ") or segment.endswith(" "):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = segment
        if i < len(segments) - 1:
            etree.SubElement(r, f"{W}tab")


def set_paragraph_text(p: etree._Element, text: str) -> None:
    clear_paragraph_content(p)
    append_run(p, text)


def read_document_xml(docx: Path) -> etree._Element:
    with zipfile.ZipFile(docx) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def write_docx_from_xml(src_docx: Path, out_docx: Path, document_xml: etree._Element) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(src_docx) as zf:
            zf.extractall(tmp)
        etree.ElementTree(document_xml).write(
            str(tmp / "word/document.xml"),
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )
        if out_docx.exists():
            out_docx.unlink()
        with zipfile.ZipFile(out_docx, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(tmp.rglob("*")):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp).as_posix())


def fill_placeholders_in_docx(src_docx: Path, out_docx: Path, fields: dict[str, Any]) -> int:
    root = read_document_xml(src_docx)
    replacements = {f"{{{{{key}}}}}": str(value) for key, value in fields.items()}
    updates = 0
    for t in root.xpath("//w:t", namespaces=NS):
        if not t.text:
            continue
        new_text = t.text
        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, value)
        if new_text != t.text:
            t.text = new_text
            updates += 1
    write_docx_from_xml(src_docx, out_docx, root)
    return updates


def next_version_name(source: Path, out_dir: Path) -> Path:
    match = re.search(r"Ver(\d+)", source.stem, flags=re.IGNORECASE)
    if not match:
        raise SystemExit(f"Could not find VerN in source filename: {source.name}")
    current = int(match.group(1))
    prefix = source.stem[: match.start(1)]
    suffix = source.stem[match.end(1) :]

    used = [current]
    for candidate in out_dir.glob("*.docx"):
        m = re.search(r"Ver(\d+)", candidate.stem, flags=re.IGNORECASE)
        if m and candidate.stem.startswith(prefix):
            used.append(int(m.group(1)))
    next_n = max(used) + 1
    return out_dir / f"{prefix}{next_n}{suffix}{source.suffix}"


def cmd_next_version(args: argparse.Namespace) -> None:
    source = Path(args.source).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = next_version_name(source, out_dir)
    if args.copy:
        shutil.copy2(source, dest)
    print(dest)


def cmd_fill_placeholders(args: argparse.Namespace) -> None:
    src = Path(args.docx).expanduser().resolve()
    out = Path(args.out).expanduser().resolve()
    fields = json.loads(Path(args.fields).expanduser().read_text(encoding="utf-8"))
    updates = fill_placeholders_in_docx(src, out, fields)
    print(json.dumps({"out": str(out), "updates": updates}, ensure_ascii=False))


def extract_docx(docx: Path) -> dict[str, Any]:
    root = read_document_xml(docx)
    paragraphs = [paragraph_text(p) for p in root.xpath("//w:p", namespaces=NS)]
    figures: list[dict[str, Any]] = []
    tables: list[dict[str, Any]] = []
    toc_entries: list[dict[str, Any]] = []
    front_matter = True

    for idx, text in enumerate(paragraphs):
        stripped = text.strip()
        if stripped in {"1. 서   론", "1. 서론", "1. Introduction"}:
            front_matter = False
        if not stripped:
            continue
        m = re.match(r"Figure\s+(\d+)\.\s*(.*)", stripped)
        if m:
            figures.append({"paragraph": idx, "number": int(m.group(1)), "text": stripped})
        m = re.match(r"Table\s+(\d+)\.\s*(.*)", stripped)
        if m:
            tables.append({"paragraph": idx, "number": int(m.group(1)), "text": stripped})
        if front_matter:
            m = re.fullmatch(r"(?:\t)?(.+?)\t(\d+)", text)
            if m:
                toc_entries.append({"paragraph": idx, "title": m.group(1).strip(), "page": int(m.group(2))})

    return {"docx": str(docx), "toc": toc_entries, "figures": figures, "tables": tables}


def cmd_extract_docx(args: argparse.Namespace) -> None:
    data = extract_docx(Path(args.docx).expanduser().resolve())
    print(json.dumps(data, ensure_ascii=False, indent=2))


def extract_pdf_pages(pdf: Path) -> list[str]:
    try:
        from pypdf import PdfReader
    except Exception as exc:  # pragma: no cover
        raise SystemExit("pypdf is required for pdf-map") from exc

    reader = PdfReader(str(pdf))
    return [" ".join((page.extract_text() or "").split()) for page in reader.pages]


def printed_page_from_text(text: str, physical_page: int, offset: int | None) -> int:
    if offset is not None:
        return physical_page - offset
    m = re.match(r"^(\d+)\s+", text)
    if m:
        return int(m.group(1))
    m = re.search(r"\s(\d+)$", text)
    if m:
        return int(m.group(1))
    return physical_page


def detect_body_start(pages: list[str]) -> int:
    for physical, text in enumerate(pages, start=1):
        compact = re.sub(r"\s+", "", text)
        frontmatter_markers = ["목차", "표목차", "그림목차", "초록", "Abstract"]
        if any(marker in compact for marker in frontmatter_markers):
            continue
        if "1.서론" in compact or "1.서론" in compact.replace(" ", ""):
            return physical
        if "1.Introduction" in compact:
            return physical
    return 1


def cmd_pdf_map(args: argparse.Namespace) -> None:
    pdf = Path(args.pdf).expanduser().resolve()
    pages = extract_pdf_pages(pdf)
    offset = args.offset
    body_start = args.body_start_physical or detect_body_start(pages)
    result: dict[str, Any] = {"pdf": str(pdf), "page_count": len(pages), "figures": {}, "tables": {}, "headings": {}}

    major_heading_pattern = r"\b([1-9]\.\s+.+?)(?=\s+[1-9]\.\d+\s|\s+Figure\s+\d+\.|\s+Table\s+\d+\.|$)"
    heading_patterns = [
        r"[1-9]\.\d+\.\s+[^.]{1,100}",
        r"[1-9]\.\d+\.\d+\.\s+[^.]{1,120}",
    ]

    for physical, text in enumerate(pages, start=1):
        if physical < body_start:
            continue
        printed = printed_page_from_text(text, physical, offset)
        for kind, key in [("Figure", "figures"), ("Table", "tables")]:
            for match in re.finditer(rf"\b{kind}\s+(\d+)\.", text):
                n = match.group(1)
                result[key].setdefault(n, {"physical_page": physical, "printed_page": printed})
        for match in re.finditer(major_heading_pattern, text):
            title = " ".join(match.group(1).split())
            result["headings"].setdefault(title, {"physical_page": physical, "printed_page": printed})
        if text.startswith("참고문헌") or re.search(r"(^|\s)참고문헌\s+\[\d+\]", text):
            result["headings"].setdefault("참고문헌", {"physical_page": physical, "printed_page": printed})
        if text.startswith("Abstract ") or text == "Abstract" or re.match(r"^\d+\s+Abstract\b", text):
            result["headings"].setdefault("Abstract", {"physical_page": physical, "printed_page": printed})
        for pattern in heading_patterns:
            for match in re.finditer(pattern, text):
                title = " ".join(match.group(0).split())
                result["headings"].setdefault(title, {"physical_page": physical, "printed_page": printed})

    if args.out:
        Path(args.out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


def load_page_map(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return data


def page_value(entry: Any) -> int:
    if isinstance(entry, int):
        return entry
    if isinstance(entry, dict):
        if "printed_page" in entry:
            return int(entry["printed_page"])
        if "page" in entry:
            return int(entry["page"])
    raise ValueError(f"Cannot derive page value from {entry!r}")


def patch_frontmatter(docx: Path, out_docx: Path, page_map: dict[str, Any]) -> int:
    root = read_document_xml(docx)
    body = root.find("w:body", NS)
    children = list(body)
    body_start = len(children)
    for idx, child in enumerate(children):
        if child.tag == f"{W}p" and paragraph_text(child).strip() in {"1. 서   론", "1. 서론", "1. Introduction"}:
            body_start = idx
            break

    toc_map = {k: page_value(v) for k, v in page_map.get("toc", {}).items()}
    heading_map = {k: page_value(v) for k, v in page_map.get("headings", {}).items()}
    figure_map = {str(k): page_value(v) for k, v in page_map.get("figures", {}).items()}
    table_map = {str(k): page_value(v) for k, v in page_map.get("tables", {}).items()}
    updates = 0

    for child in children[:body_start]:
        if child.tag != f"{W}p":
            continue
        text = paragraph_text(child)
        if not text.strip():
            continue

        fig_match = re.match(r"Figure\s+(\d+)\.\s", text)
        if fig_match and re.search(r"\s\d+$", text):
            num = fig_match.group(1)
            if num in figure_map:
                updated = re.sub(r"\s\d+$", f" {figure_map[num]}", text)
                if updated != text:
                    set_paragraph_text(child, updated)
                    updates += 1
            continue

        table_match = re.match(r"Table\s+(\d+)\.\s", text)
        if table_match and re.search(r"\s\d+$", text):
            num = table_match.group(1)
            if num in table_map:
                updated = re.sub(r"\s\d+$", f" {table_map[num]}", text)
                if updated != text:
                    set_paragraph_text(child, updated)
                    updates += 1
            continue

        toc_match = re.fullmatch(r"(?:\t)?(.+?)\t(\d+)", text)
        if toc_match:
            title = toc_match.group(1).strip()
            new_page = toc_map.get(title) or heading_map.get(title)
            if new_page is not None and int(toc_match.group(2)) != new_page:
                set_paragraph_text(child, f"\t{title}\t{new_page}")
                updates += 1
            continue

    write_docx_from_xml(docx, out_docx, root)
    return updates


def cmd_patch_frontmatter(args: argparse.Namespace) -> None:
    docx = Path(args.docx).expanduser().resolve()
    out_docx = Path(args.out).expanduser().resolve()
    page_map = load_page_map(Path(args.map).expanduser().resolve())
    updates = patch_frontmatter(docx, out_docx, page_map)
    print(json.dumps({"out": str(out_docx), "updates": updates}, ensure_ascii=False))


WORD_EXPORT_APPLESCRIPT = r'''
on run argv
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set outputHfsPath to (POSIX file outputPath) as text
    set startupDisk to path to startup disk
    tell application id "com.microsoft.Word"
        try
            close startupDisk
        end try
    end tell
    tell application "Microsoft Word"
        set oldAlerts to display alerts
        set display alerts to alerts none
        open POSIX file inputPath
        repeat while (count of documents) is 0
            delay 0.5
        end repeat
        set theDoc to document 1
        save as theDoc file name outputHfsPath file format format PDF
        close window 1 saving no
        set display alerts to oldAlerts
    end tell
end run
'''


def microsoft_word_available() -> bool:
    if not shutil.which("osascript"):
        return False
    result = subprocess.run(
        ["osascript", "-e", 'id of application "Microsoft Word"'],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.returncode == 0


def cmd_render_pdf(args: argparse.Namespace) -> None:
    docx = Path(args.docx).expanduser().resolve()
    out_pdf = Path(args.out).expanduser().resolve()
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    if not microsoft_word_available():
        raise SystemExit("Microsoft Word is required for authoritative thesis PDF export.")
    if out_pdf.exists():
        out_pdf.unlink()
    with tempfile.TemporaryDirectory() as tmpdir:
        script_path = Path(tmpdir) / "export_word_pdf.applescript"
        script_path.write_text(WORD_EXPORT_APPLESCRIPT, encoding="utf-8")
        result = subprocess.run(
            ["osascript", str(script_path), str(docx), str(out_pdf)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if result.returncode != 0:
            details = "\n".join(
                part
                for part in [
                    f"osascript exited with status {result.returncode}",
                    f"stdout: {result.stdout.strip()}" if result.stdout.strip() else "",
                    f"stderr: {result.stderr.strip()}" if result.stderr.strip() else "",
                ]
                if part
            )
            raise SystemExit(details)
    if not out_pdf.exists():
        raise SystemExit(f"Microsoft Word did not generate PDF at {out_pdf}")
    print(out_pdf)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("next-version", help="Compute or copy the next VerN DOCX path")
    p.add_argument("--source", required=True)
    p.add_argument("--out-dir", required=True)
    p.add_argument("--copy", action="store_true")
    p.set_defaults(func=cmd_next_version)

    p = sub.add_parser("extract-docx", help="Extract TOC, figure, and table text from DOCX")
    p.add_argument("docx")
    p.set_defaults(func=cmd_extract_docx)

    p = sub.add_parser("fill-placeholders", help="Fill {{PLACEHOLDER}} slots in a DOCX template")
    p.add_argument("docx")
    p.add_argument("--fields", required=True, help="JSON object mapping placeholder names to values")
    p.add_argument("--out", required=True)
    p.set_defaults(func=cmd_fill_placeholders)

    p = sub.add_parser("pdf-map", help="Extract detected figure/table/heading page map from PDF")
    p.add_argument("pdf")
    p.add_argument("--out")
    p.add_argument("--offset", type=int, help="printed page = physical PDF page - offset")
    p.add_argument("--body-start-physical", type=int, help="first physical PDF page of the body; auto-detected when omitted")
    p.set_defaults(func=cmd_pdf_map)

    p = sub.add_parser("patch-frontmatter", help="Patch TOC/list page numbers from a JSON page map")
    p.add_argument("docx")
    p.add_argument("--map", required=True)
    p.add_argument("--out", required=True)
    p.set_defaults(func=cmd_patch_frontmatter)

    p = sub.add_parser("render-pdf", help="Export DOCX to PDF through Microsoft Word")
    p.add_argument("docx")
    p.add_argument("--out", required=True)
    p.set_defaults(func=cmd_render_pdf)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
