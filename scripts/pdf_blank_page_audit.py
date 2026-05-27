#!/usr/bin/env python3
"""Find blank pages in a rendered thesis PDF or its page PNGs."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


def page_number(path: Path, fallback: int) -> int:
    numbers = re.findall(r"\d+", path.stem)
    return int(numbers[-1]) if numbers else fallback


def image_ink_ratio(path: Path, threshold: int, crop_top: float, crop_bottom: float) -> float:
    try:
        from PIL import Image
    except ImportError as exc:
        raise SystemExit("Pillow가 필요합니다. page PNG 기반 빈 페이지 감사에는 PIL/Pillow를 설치하세요.") from exc

    with Image.open(path) as img:
        gray = img.convert("L")
        width, height = gray.size
        top = int(height * crop_top)
        bottom = int(height * (1.0 - crop_bottom))
        if bottom <= top:
            bottom = height
            top = 0
        cropped = gray.crop((0, top, width, bottom))
        hist = cropped.histogram()
        dark_pixels = sum(hist[:threshold])
        return dark_pixels / max(cropped.size[0] * cropped.size[1], 1)


def audit_rendered_pages(
    rendered_dir: Path,
    glob_pattern: str,
    threshold: int,
    min_ink_ratio: float,
    crop_top: float,
    crop_bottom: float,
) -> dict:
    paths = sorted(rendered_dir.glob(glob_pattern), key=lambda p: (page_number(p, 0), p.name))
    pages = []
    blank_pages = []
    for fallback, path in enumerate(paths, start=1):
        number = page_number(path, fallback)
        ratio = image_ink_ratio(path, threshold, crop_top, crop_bottom)
        item = {"page": number, "image": str(path), "ink_ratio": ratio}
        pages.append(item)
        if ratio < min_ink_ratio:
            blank_pages.append(item)
    return {
        "mode": "rendered_images",
        "rendered_dir": str(rendered_dir),
        "page_count": len(paths),
        "min_ink_ratio": min_ink_ratio,
        "crop_top_ratio": crop_top,
        "crop_bottom_ratio": crop_bottom,
        "ok": not blank_pages,
        "blank_pages": blank_pages,
        "pages": pages,
    }


def audit_pdf_text(pdf_path: Path) -> dict:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise SystemExit("pypdf가 필요합니다. PDF 텍스트 기반 감사에는 pypdf를 설치하세요.") from exc

    reader = PdfReader(str(pdf_path))
    blank_pages = []
    for index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if not text.strip():
            blank_pages.append({"page": index, "reason": "no extracted text"})
    return {
        "mode": "pdf_text",
        "pdf": str(pdf_path),
        "page_count": len(reader.pages),
        "ok": not blank_pages,
        "blank_pages": blank_pages,
        "warning": "Text-only mode can miss pages that contain only page numbers. Prefer --rendered-dir.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("target", type=Path, help="PDF path, or a rendered page directory when --rendered-dir is not supplied.")
    parser.add_argument("--rendered-dir", type=Path, help="Directory containing rendered page PNGs.")
    parser.add_argument("--page-glob", default="*.png")
    parser.add_argument("--ink-threshold", type=int, default=245)
    parser.add_argument("--min-ink-ratio", type=float, default=0.00045)
    parser.add_argument("--crop-top-ratio", type=float, default=0.0)
    parser.add_argument("--crop-bottom-ratio", type=float, default=0.08)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    rendered_dir = args.rendered_dir or (args.target if args.target.is_dir() else None)
    if rendered_dir:
        report = audit_rendered_pages(
            rendered_dir,
            args.page_glob,
            args.ink_threshold,
            args.min_ink_ratio,
            args.crop_top_ratio,
            args.crop_bottom_ratio,
        )
    else:
        report = audit_pdf_text(args.target)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    elif report["ok"]:
        print(f"OK: blank page audit passed ({report['page_count']} pages).")
    else:
        pages = ", ".join(str(item["page"]) for item in report["blank_pages"])
        print(f"FAIL: blank page candidates found on page(s): {pages}")
    return 0 if report["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
