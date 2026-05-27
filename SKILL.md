---
name: snu-thesis-production
description: Produce or revise Seoul National University graduate thesis DOCX/PDF deliverables, especially materials science and engineering theses. Use when working on SNU thesis drafts, Korean/English thesis formatting, versioned DOCX/PDF outputs, figure/table/TOC renumbering, render-based QA, or repeatable thesis production workflows.
---

# SNU Thesis Production

Use this skill as a production pipeline, not as a prose-only editing prompt. Always combine it with the Documents skill for DOCX rendering and visual QA.

For review papers, literature-review theses, survey papers, or "find the hottest topic and write a review" requests, also read `references/review-paper-workflow.md`. Start each manuscript in its own project folder with `scripts/init_review_project.py`; do not keep manuscript-specific config, ledgers, scripts, images, or outputs in the skill root.

If the user is new, opened this downloaded folder as a Codex project, or does not know what information to provide, first read `references/new-user-onboarding.md` and run `python scripts/first_run_questions.py`. Ask the generated questions in plain language before creating a manuscript.

## Core Contract

1. Preserve the latest user-edited source. Never overwrite it.
2. Create a new `VerN` DOCX for every edit batch and create the matching PDF.
3. Treat the official SNU thesis rules and the project config as higher priority than local style habits.
4. Preserve the user's research meaning, evidence, data, quotations, terminology, names, dates, calculations, and references. Do not silently put `[확인 필요]` into a final draft; run the metadata preflight and ask the user for missing required values first. Use placeholders only when the user explicitly approves a placeholder draft.
5. Do not change intentionally rotated images, user-arranged figure layouts, or unrelated sections unless explicitly requested.
6. Preserve the front-matter layout template. For the cover/title/approval pages, change only metadata slots such as title, author, advisor, committee, department, degree label, and dates.
7. Export the DOCX to PDF through Microsoft Word, inspect page images, then fix and re-export until the requested area is clean.
8. After figure/table moves, synchronize body references, captions, lists of figures/tables, and PDF page numbers.
9. Use Batang for all Korean text, Times New Roman for all English/Latin text, and black font color throughout the document unless the user explicitly requests a different style.
10. Keep project-specific artifacts isolated: each thesis/review project gets its own folder containing its configs, ledgers, generated scripts, source images, QA images, DOCX/PDF versions, and page maps.
11. When the user asks to keep this workflow synchronized with GitHub, use `references/github-sync-workflow.md` and `scripts/github_sync.py`. If remote URL, git identity, or credentials are missing, prepare the local repo and ask for only the missing values instead of guessing.
12. For every figure/table page, follow `references/visual-item-layout-policy.md`: no clipping, no page break continuation, no body text, centered layout in the rendered PDF.

## Required Project Inputs

Before major work, locate or create a project folder. For a new review/thesis project, run:

```bash
python scripts/init_review_project.py "Project Name" --root projects
```

Then work only inside that project folder for manuscript-specific files. Locate or create a project config from `references/project-config-template.yaml`; it must identify:

- latest source DOCX and output folder
- versioning rule
- SNU metadata: degree type, title, author, department, advisor, committee, submission/approval month
- fixed terminology
- project-specific domain rules, interpretation limits, and citation rules. For materials science/engineering theses, start from `references/materials-engineering-defaults.yaml` and override per project.
- files or regions not to touch
- render/audit expectations

If the user gives instructions in chat that supersede the config, follow the newest user instruction and update the next working assumptions.

Before generating a new thesis/review DOCX, run the project metadata preflight. If degree type, department/major, advisor, submission/approval month, author, committee roles, title, or page-count strictness are missing, ask the user for those values in a concise checklist and pause document generation until they answer, unless they explicitly ask for a placeholder draft.

## Workflow

### Review-Paper Add-On

When the task is a review paper:

1. Read `references/review-paper-workflow.md`.
2. Create or select a project folder; for new work use `scripts/init_review_project.py`.
3. Run `scripts/review_preflight.py <project-folder>/config/review-config.yaml` or perform the same checklist manually before drafting; ask for missing required fields instead of inserting `[확인 필요]`.
4. Browse current literature when the user asks for a recent, hot, latest, or field-discovery topic.
5. Create or update the project-local review config and evidence ledger before drafting.
6. If the user asks to use source images, keep the project-local image ledger with source URL, figure number, license/permission status, and caption credit.
7. Any topic-specific builder script belongs in `<project-folder>/scripts/`, not in the skill root.
8. If the user corrects the workflow for future runs, update the relevant skill reference or reusable script first, then continue the document work.
9. If GitHub sync is enabled for the skill, commit and push the reusable workflow update after validation.

1. **Orient**
   - Find the latest `VerN` DOCX in the output folder unless the user names a source.
   - Read the current request and list the affected sections, figures, tables, captions, and front matter.
   - Use `scripts/thesis_tool.py extract-docx <docx>` to inspect captions and front-matter entries when numbering may change.

2. **Version**
   - Use `scripts/thesis_tool.py next-version --source <docx> --out-dir <folder> --copy` to create the next DOCX.
   - Keep all edits in the new version only.

3. **Front Matter**
   - For this review-paper workflow, default front matter is exactly: page 1 cover, page 2 approval page, page 3 Korean abstract. Do not create a separate submission/title page unless the user asks for it.
   - Use the existing front matter as a locked layout, or start from `assets/snu-materials-frontmatter-template.docx` for a new thesis when the official full SNU template is required.
   - Fill only placeholder slots listed in `references/frontmatter-fields-template.json`.
   - Use `scripts/thesis_tool.py fill-placeholders <template.docx> --fields <fields.json> --out <docx>` when generating from the placeholder template.
   - Keep title line breaks deliberate. If a title needs more or fewer lines, adjust only the title slot paragraphs while preserving the page's overall spacing and alignment.

4. **Edit**
   - Prefer minimal OOXML edits for existing DOCX files so layout and images survive.
   - For figure moves, move the image paragraph(s), caption paragraph, nearby explanatory paragraphs, and list entry together where appropriate.
   - Update body prose so each figure/table is mentioned before insertion.
   - Keep captions concise; put interpretation in the body.
   - Figures and tables must be isolated on their own page: no body text on that page, only the figure/table and its caption. Use one item per page by default; two small related figures/tables may share one page.
   - A figure/table must never split, crop, overflow, or continue onto another page. One item is centered as a single group on the page; two items are centered as one balanced group with clear spacing.
   - Omit acknowledgements unless the user explicitly requests them.

5. **Render**
   - Export the final PDF through Microsoft Word, not LibreOffice, Pandoc, Preview, or a generic DOCX renderer. Word pagination is authoritative.
   - Use `scripts/thesis_tool.py render-pdf <docx> --out <pdf>` on macOS. This opens the DOCX in Microsoft Word and saves it as PDF.
   - Use a PDF-page rasterizer on the Word-exported PDF for visual QA. If page-image rendering fails, manually inspect the Word-exported PDF and state the limitation.
   - Never use a non-Word PDF export for final page counts, TOC/list page-number audits, or final deliverables unless the user explicitly approves a fallback.

6. **Audit Page Numbers**
   - Run `scripts/thesis_tool.py pdf-map <pdf> --out <map.json>` to get detected heading, figure, and table pages.
   - If the front matter uses printed page numbers rather than physical PDF page indexes, use the script's `printed_page` values or pass `--offset`.
   - Patch front matter with `scripts/thesis_tool.py patch-frontmatter <docx> --map <map.json> --out <docx>`.
   - Re-export the PDF through Microsoft Word after patching.

7. **Visual QA**
   - Inspect the table of contents, list of tables, list of figures, and every page touched by the change.
   - Check for wrong numbering, stale page numbers, broken captions, image overflow, clipping, table row/page splitting, excessive blank space, failed centering, text overlap, and accidental layout changes.
   - For rendered figure/table page PNGs, use `scripts/visual_page_qa.py <page-images>` as a screening check when available, then visually inspect the Word-exported PDF.
   - For moved figures, inspect the page before, the figure page, and the page after.

8. **Final Response**
   - Report only the new DOCX/PDF paths, the high-level changes, and what was verified.
   - Mention any render limitation or unresolved preflight/evidence-ledger items.

## Reusable Resources

- `references/project-config-template.yaml`: copy into a thesis project and fill in once.
- `references/new-user-onboarding.md`: first-use question flow for users who downloaded the folder and do not know the workflow.
- `references/review-paper-workflow.md`: workflow for topic discovery, evidence-ledger construction, sourced images, page-count control, and reusable revision learning for review papers.
- `references/review-paper-config-template.yaml`: starting config for review-thesis projects.
- `references/github-sync-workflow.md`: local git/GitHub synchronization workflow for preserving reusable skill and project changes.
- `references/visual-item-layout-policy.md`: figure/table page isolation, centering, fit, and PDF QA rules.
- `scripts/init_review_project.py`: scaffolds a clean project folder so manuscript-specific files never pollute the skill root.
- `scripts/first_run_questions.py`: prints the beginner-friendly starting checklist before project setup.
- `scripts/review_preflight.py`: checks a review config for missing required metadata and prints the exact user questions to ask before drafting.
- `scripts/github_sync.py`: initializes git, commits changes, and pushes to GitHub when a remote and credentials are configured.
- `scripts/visual_page_qa.py`: screens rendered page images for content centering and edge-clipping risk.
- `references/snu-thesis-rules.md`: compact SNU thesis production rules and scientific-writing guardrails.
- `references/materials-engineering-defaults.yaml`: optional baseline terminology and evidence rules for materials science/engineering theses.
- `references/frontmatter-fields-template.json`: metadata slots for the locked front-matter template.
- `assets/snu-materials-frontmatter-template.docx`: placeholder DOCX template for SNU materials-engineering thesis cover/title/approval pages.
- `scripts/thesis_tool.py`: versioning, DOCX extraction, PDF page mapping, and front-matter page patching utilities.

## Decision Rules

- If a change can alter pagination, always re-export PDF through Microsoft Word and refresh lists.
- If a caption becomes too long, shorten it and move detailed interpretation into body prose.
- If evidence was not directly established, write the statement as an interpretation, limitation, or possibility.
- If a source/citation cannot be verified, do not invent bibliographic data.
- If a user says a figure rotation/layout is intentional, treat it as locked.
- If front-matter metadata changes, edit only the corresponding slot. Do not restyle or reflow the entire cover/title/approval layout.
- If required thesis metadata is missing, ask before generation instead of hiding it as `[확인 필요]` in the manuscript.
