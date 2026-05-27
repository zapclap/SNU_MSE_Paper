#!/usr/bin/env python3
"""Scaffold an isolated review-thesis project folder.

Reusable skill files stay in the skill root. Manuscript-specific configs,
ledgers, source figures, generated figures, QA images, page maps, and builder
scripts belong inside the project folder this script creates.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def slugify(text: str) -> str:
    slug = re.sub(r"[^\w가-힣.-]+", "-", text.strip(), flags=re.UNICODE).strip("-._")
    return slug or f"review-project-{date.today().isoformat()}"


def write_if_missing(path: Path, content: str, force: bool = False) -> None:
    if path.exists() and not force:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def q(value: object) -> str:
    return json.dumps(str(value), ensure_ascii=False)


def review_config(args: argparse.Namespace, project_dir: Path) -> str:
    return f'''# Project-local review configuration. Keep manuscript-specific metadata here.
review_project:
  name: {q(args.name)}
  topic_request: {q(args.topic_request or '')}
  selected_topic: ""
  topic_rationale: ""
  target_pages: {args.target_pages}
  page_count_policy:
    exact_target_required: null
    allow_over_target: null
    confirm_before_generation: true
  language: {q(args.language)}
  citation_style: {q(args.citation_style)}
  image_mode: {q(args.image_mode)}
  source_figure_typography_policy: "ask_if_strict_typography"
  placeholder_draft_approved: false

metadata:
  author: {q(args.author or '')}
  korean_title: ""
  english_title: ""
  degree_name: ""
  school_name: "서울대학교 대학원"
  department: ""
  major: ""
  advisor: ""
  submission_month: ""
  approval_month: ""
  acknowledgements: false
  committee:
    chair: ""
    vice_chair: ""
    members: []

layout_rules:
  front_matter:
    page_1: "cover"
    page_2: "approval"
    page_3: "korean_abstract"
    omit_separate_submission_page: true
    omit_acknowledgements_by_default: true
  typography:
    korean_font: "Batang"
    english_font: "Times New Roman"
    generated_figure_korean_font: "Batang"
    generated_figure_english_font: "Times New Roman"
    font_color: "000000"
    allow_colored_text: false
    fail_if_required_font_missing: true
    font_audit_command: "scripts/docx_font_audit.py <docx>"
  figures_and_tables:
    isolated_pages: true
    allowed_content_on_item_page: ["figure_or_table", "caption"]
    default_items_per_page: 1
    max_items_per_page_when_small: 2
    never_split_across_pages: true
    never_clip_or_crop: true
    center_single_item_group: true
    center_two_item_group: true
    two_item_gap_min_pt: 12
    pdf_center_tolerance_ratio: 0.08
    pdf_edge_margin_min_ratio: 0.02

paths:
  project_dir: {q(project_dir)}
  docx_dir: {q(project_dir / 'output' / 'docx')}
  pdf_dir: {q(project_dir / 'output' / 'pdf')}
  figures_dir: {q(project_dir / 'output' / 'figures')}
  source_figures_dir: {q(project_dir / 'output' / 'source_figures')}
  qa_dir: {q(project_dir / 'output' / 'qa')}
  page_maps_dir: {q(project_dir / 'output' / 'page_maps')}
  font_audit_dir: {q(project_dir / 'output' / 'qa' / 'font')}
  blank_page_audit_dir: {q(project_dir / 'output' / 'qa' / 'blank_pages')}

final_audits:
  no_blank_pages: true
  blank_page_audit_command: "scripts/pdf_blank_page_audit.py <pdf> --rendered-dir <all-page-png-dir>"
  render_all_pages_before_blank_audit: true
  font_audit_required: true
  blank_page_audit_required: true
'''


def project_config(args: argparse.Namespace, project_dir: Path) -> str:
    return f'''project:
  name: {q(args.name)}
  project_dir: {q(project_dir)}
  output_dir: {q(project_dir / 'output')}
  latest_source_docx: ""
  version_pattern: "Ver{{n}}"
  make_pdf_for_every_version: true

institution:
  official_rules_file: {q(ROOT / 'references' / 'snu-thesis-rules.md')}
  school_name: "서울대학교 대학원"
  degree_name: ""
  department: ""
  major: ""
  author: {q(args.author or '')}
  student_id: ""
  advisor: ""
  committee:
    chair: ""
    vice_chair: ""
    members: []
  submission_month: ""
  approval_month: ""
  korean_title: ""
  english_title: ""

layout_rules:
  front_matter_order: ["cover", "approval", "korean_abstract", "toc", "list_of_tables", "list_of_figures", "body", "references", "appendix", "english_abstract"]
  approval_page_must_end_on_page: 2
  omit_acknowledgements: true
  figure_table_isolated_pages: true
  max_items_per_figure_table_page: 2
  figure_table_never_split_across_pages: true
  figure_table_never_clip_or_crop: true
  figure_table_center_single_item_group: true
  figure_table_center_two_item_group: true
  figure_table_two_item_gap_min_pt: 12
  figure_table_pdf_center_tolerance_ratio: 0.08
  figure_table_pdf_edge_margin_min_ratio: 0.02
  font_color: "000000"
  korean_font: "Batang"
  english_font: "Times New Roman"
  generated_figure_korean_font: "Batang"
  generated_figure_english_font: "Times New Roman"
  fail_if_required_font_missing: true
  font_audit_command: "scripts/docx_font_audit.py <docx>"
  no_blank_pages: true
  blank_page_audit_command: "scripts/pdf_blank_page_audit.py <pdf> --rendered-dir <all-page-png-dir>"
'''


def evidence_ledger(args: argparse.Namespace) -> str:
    return f'''date_created: "{date.today().isoformat()}"
topic_request: {q(args.topic_request or '')}
selected_topic: ""
topic_rationale: ""
search_queries: []
sources: []
unresolved_checks: []
'''


def image_ledger() -> str:
    return f'''date_created: "{date.today().isoformat()}"
image_mode: ""
imported_images: []
generated_images: []
unresolved_permissions: []
'''


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("name", help="Human-readable project name.")
    parser.add_argument("--root", default=str(ROOT / "projects"), help="Directory that contains project folders.")
    parser.add_argument("--slug", help="Folder name. Defaults to a sanitized project name.")
    parser.add_argument("--author", default="")
    parser.add_argument("--topic-request", default="")
    parser.add_argument("--target-pages", type=int, default=50)
    parser.add_argument("--language", default="ko")
    parser.add_argument("--citation-style", default="numeric")
    parser.add_argument("--image-mode", default="source_figures_with_references")
    parser.add_argument("--force", action="store_true", help="Overwrite scaffold files if they already exist.")
    args = parser.parse_args()

    project_dir = Path(args.root).expanduser().resolve() / (args.slug or slugify(args.name))
    for subdir in [
        "config",
        "ledgers",
        "notes",
        "scripts",
        "output/docx",
        "output/pdf",
        "output/figures",
        "output/source_figures",
        "output/qa",
        "output/page_maps",
    ]:
        (project_dir / subdir).mkdir(parents=True, exist_ok=True)

    write_if_missing(project_dir / "config" / "review-config.yaml", review_config(args, project_dir), args.force)
    write_if_missing(project_dir / "config" / "project-config.yaml", project_config(args, project_dir), args.force)
    write_if_missing(project_dir / "ledgers" / "evidence-ledger.yaml", evidence_ledger(args), args.force)
    write_if_missing(project_dir / "ledgers" / "image-ledger.yaml", image_ledger(), args.force)
    write_if_missing(
        project_dir / "notes" / "README.txt",
        "Put project-specific markdown notes, outlines, and drafting scratch files here.\n",
        args.force,
    )
    write_if_missing(
        project_dir / "scripts" / "README.txt",
        "Put project-specific manuscript builders or data-cleaning scripts here. Do not put them in the skill root.\n",
        args.force,
    )

    print(project_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
