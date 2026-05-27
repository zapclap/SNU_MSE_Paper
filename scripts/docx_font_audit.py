#!/usr/bin/env python3
"""Audit DOCX font slots for the SNU thesis workflow.

The check is intentionally deterministic: Korean/East Asian font slots must be
Batang, Latin slots must be Times New Roman, and explicit text colors must be
black or automatic. It inspects actual content parts plus styles used by the
document, avoiding failures from unused Word built-in styles.
"""

from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def wq(name: str) -> str:
    return f"{{{W_NS}}}{name}"


EXPECTED_FONTS = {
    "ascii": "Times New Roman",
    "hAnsi": "Times New Roman",
    "cs": "Times New Roman",
    "eastAsia": "Batang",
}

CONTENT_PREFIXES = (
    "word/document.xml",
    "word/header",
    "word/footer",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
)


def parse_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def content_part_names(zf: zipfile.ZipFile) -> list[str]:
    names: list[str] = []
    for name in zf.namelist():
        if name.endswith(".xml") and any(name.startswith(prefix) for prefix in CONTENT_PREFIXES):
            names.append(name)
    return sorted(names)


def rfont_values(r_fonts: ET.Element | None) -> dict[str, str | None]:
    if r_fonts is None:
        return {slot: None for slot in EXPECTED_FONTS}
    return {slot: r_fonts.get(wq(slot)) for slot in EXPECTED_FONTS}


def check_font_slots(values: dict[str, str | None], location: str, violations: list[dict]) -> None:
    for slot, expected in EXPECTED_FONTS.items():
        actual = values.get(slot)
        if actual is not None and actual != expected:
            violations.append(
                {
                    "type": "font",
                    "location": location,
                    "slot": slot,
                    "expected": expected,
                    "actual": actual,
                }
            )


def check_color(color: ET.Element | None, location: str, violations: list[dict]) -> None:
    if color is None:
        return
    value = color.get(wq("val"))
    if value and value.lower() not in {"000000", "auto"}:
        violations.append(
            {
                "type": "color",
                "location": location,
                "expected": "000000 or auto",
                "actual": value,
            }
        )


def collect_used_styles(roots: list[ET.Element]) -> set[str]:
    used = {"Normal"}
    for root in roots:
        for style_ref in root.findall(".//w:pStyle", NS) + root.findall(".//w:rStyle", NS):
            value = style_ref.get(wq("val"))
            if value:
                used.add(value)
    return used


def style_map(styles_root: ET.Element | None) -> dict[str, ET.Element]:
    if styles_root is None:
        return {}
    styles = {}
    for style in styles_root.findall(".//w:style", NS):
        style_id = style.get(wq("styleId"))
        if style_id:
            styles[style_id] = style
    return styles


def check_doc_defaults(styles_root: ET.Element | None, violations: list[dict]) -> None:
    if styles_root is None:
        violations.append({"type": "font", "location": "word/styles.xml", "issue": "styles.xml missing"})
        return
    r_fonts = styles_root.find("./w:docDefaults/w:rPrDefault/w:rPr/w:rFonts", NS)
    if r_fonts is None:
        violations.append(
            {
                "type": "font",
                "location": "word/styles.xml docDefaults",
                "issue": "missing docDefaults rFonts",
            }
        )
        return
    values = rfont_values(r_fonts)
    for slot, expected in EXPECTED_FONTS.items():
        actual = values.get(slot)
        if actual != expected:
            violations.append(
                {
                    "type": "font",
                    "location": "word/styles.xml docDefaults",
                    "slot": slot,
                    "expected": expected,
                    "actual": actual,
                }
            )


def check_used_styles(styles_root: ET.Element | None, used_styles: set[str], violations: list[dict]) -> None:
    styles = style_map(styles_root)
    for style_id in sorted(used_styles):
        style = styles.get(style_id)
        if style is None:
            continue
        location = f"word/styles.xml style:{style_id}"
        r_fonts = style.find("./w:rPr/w:rFonts", NS)
        check_font_slots(rfont_values(r_fonts), location, violations)
        check_color(style.find("./w:rPr/w:color", NS), location, violations)


def run_audit(docx_path: Path) -> dict:
    violations: list[dict] = []
    with zipfile.ZipFile(docx_path) as zf:
        roots = []
        for name in content_part_names(zf):
            root = parse_xml(zf, name)
            if root is not None:
                roots.append((name, root))

        styles_root = parse_xml(zf, "word/styles.xml")
        check_doc_defaults(styles_root, violations)
        check_used_styles(styles_root, collect_used_styles([root for _, root in roots]), violations)

        for part_name, root in roots:
            for index, run in enumerate(root.findall(".//w:r", NS), start=1):
                text = "".join(t.text or "" for t in run.findall(".//w:t", NS))
                if not text.strip():
                    continue
                location = f"{part_name} run:{index}"
                r_pr = run.find("./w:rPr", NS)
                if r_pr is None:
                    continue
                check_font_slots(rfont_values(r_pr.find("./w:rFonts", NS)), location, violations)
                check_color(r_pr.find("./w:color", NS), location, violations)

    return {
        "docx": str(docx_path),
        "expected": EXPECTED_FONTS,
        "ok": not violations,
        "violations": violations,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    args = parser.parse_args()

    report = run_audit(args.docx)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    elif report["ok"]:
        print(f"OK: {args.docx} uses Batang/Times New Roman font slots and black text.")
    else:
        print(f"FAIL: {args.docx} has {len(report['violations'])} font/color violation(s).")
        for item in report["violations"][:40]:
            print(json.dumps(item, ensure_ascii=False))
        if len(report["violations"]) > 40:
            print(f"... {len(report['violations']) - 40} more")
    return 0 if report["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
