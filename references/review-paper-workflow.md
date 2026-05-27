# Review Paper Thesis Workflow

Use this reference when the user asks for a review paper, literature-review thesis, survey paper, or topic-discovery review in SNU thesis form. The goal is not only to produce one document, but to preserve the workflow so future revisions can be repeated.

## Operating Principles

1. Treat the review as an evidence pipeline, not a prose-only task.
2. Browse for current literature whenever the topic asks for "latest", "hot", "recent", or "current".
3. Prefer primary sources, review articles, official reports, and publisher pages with DOI metadata.
4. Never invent bibliographic metadata. Do not put `[확인 필요]` directly into the final manuscript as a substitute for asking the user; first ask for missing user/project metadata, and keep unresolved source metadata in the evidence ledger.
5. Keep a project-local evidence ledger: topic rationale, search queries, selected sources, figures, licenses, and unresolved checks.
6. Make a new `VerN` DOCX/PDF for every edit batch.
7. After the user gives a workflow preference, update this reference or the relevant script before continuing.
8. Use Batang for all Korean text, Times New Roman for all English/Latin text, and black document text only.
9. Keep every manuscript in its own project folder. Never place topic-specific configs, source images, generated figures, builder scripts, ledgers, QA images, or DOCX/PDF outputs in the skill root.
10. If GitHub sync is enabled, commit reusable workflow changes with `scripts/github_sync.py` after validation.
11. For visual pages, apply `references/visual-item-layout-policy.md` and verify centering from the Word-exported PDF.

## Project Folder Isolation

For a first-time user who has only downloaded the folder and opened it in Codex, do not expect them to know the config structure. Read `references/new-user-onboarding.md`, run `python scripts/first_run_questions.py`, ask the checklist in plain language, and then create the project folder.

For every new review paper, create or select a project folder first:

```bash
python scripts/init_review_project.py "Project Name" --root projects
```

Use this structure:

```text
project-folder/
  config/
    project-config.yaml
    review-config.yaml
  ledgers/
    evidence-ledger.yaml
    image-ledger.yaml
  notes/
    *.md
  scripts/
    build_<project>.py
  output/
    docx/
    pdf/
    figures/
    source_figures/
    qa/
    page_maps/
```

Reusable skill scripts stay in the skill `scripts/` folder. Topic-specific generation scripts, markdown notes/drafts, and all manuscript data stay in the project `scripts/`, `notes/`, `config/`, `ledgers/`, and `output/` folders. When revising an existing project, locate its folder first and do not create or edit root-level manuscript files.

## Preflight Before Drafting

Before writing the DOCX, run:

```bash
python scripts/review_preflight.py <project-folder>/config/review-config.yaml
```

If required fields are missing, ask the user the generated questions and pause manuscript generation. Do not generate a final DOCX/PDF with `[확인 필요]` in the front matter. If the user explicitly asks for a rough placeholder draft, record `placeholder_draft_approved: true` in the config and list unresolved fields in the final response.

Required or inferable:
- author name
- Korean title
- English title
- degree label
- department and major
- advisor
- submission month and approval month
- committee names and roles
- target page count
- whether the page count must be exact, or whether the manuscript may exceed the target when figure/table isolation and centering require more pages
- language
- citation style

Optional:
- student ID
- acknowledgements; omit by default

## Topic Discovery

When the user asks Codex to find the hottest topic:

1. Search recent literature with at least three query angles:
   - field + "review" + current years
   - field + application bottleneck
   - field + material/process/mechanism keywords
2. Prefer sources from the last 24 months, but include foundational references if needed.
3. Choose the topic with a short rationale based on:
   - publication clustering
   - industrial relevance
   - unresolved mechanism or design gap
   - fit to the user's domain
4. Save the rationale in the project-local config or evidence ledger.

## Evidence Ledger

For each selected source, record:

```yaml
- id: R01
  title: ""
  authors: ""
  year: ""
  venue: ""
  doi: ""
  url: ""
  source_type: primary|review|official_report|other
  used_for: topic|mechanism|figure|table|roadmap|background
  confidence: verified|partial|needs_check
  notes: ""
```

## Image Policy For Review Papers

If the user wants images copied from papers:

1. Use open-access or clearly licensed figures first.
2. For every imported figure, save an image ledger entry:

```yaml
- local_path: ""
  original_figure: "Figure 3"
  source_ref: "R01"
  article_url: ""
  image_url: ""
  license: ""
  caption_credit: "Reproduced/adapted from ..."
  needs_permission: false
```

3. Captions must include the source reference and whether the image is reproduced or adapted.
4. If the license is unclear, either:
   - use the image only in an internal draft and mark `needs_permission: true`, or
   - redraw it as an original schematic and cite the source as inspiration.
5. Do not leave imported images uncited.

## Outline Construction

Build the review around a thesis-shaped argument:

1. Background and why the topic is timely.
2. Fundamental mechanisms and definitions.
3. State of the art grouped by mechanism/application, not by one-paper summaries.
4. Comparative tables that make the literature searchable.
5. Figures that explain mechanisms, taxonomies, and research gaps.
6. Limitations of current knowledge.
7. Roadmap and conclusion.

Avoid a pure annotated bibliography. Every cited source should support a synthesis claim.

## Drafting Rules

- Mention every figure/table in body text before insertion.
- Keep figure captions concise and move interpretation into prose.
- Separate verified evidence from interpretation.
- Missing user/project metadata must be collected through preflight questions before generation. Do not leave `[확인 필요]` in the final manuscript unless the user explicitly approves a placeholder draft.
- Preserve exact user-provided names and committee roles.
- Keep imported figure credits in captions and references.
- Use only black font color in the document. No colored headings, colored table text, or colored emphasis.
- Use Batang for all Korean text and Times New Roman for all English/Latin text.
- Omit acknowledgements unless the user asks to include them.

## Front Matter Rules

Default review-paper front matter:

1. Page 1: cover.
2. Page 2: approval page. The approval content must fit on this single page.
3. Page 3: Korean abstract.
4. Then table of contents, list of tables, list of figures, body, references, appendix if any, English abstract.

Do not create a separate submission/title page in between the cover and approval page unless the user explicitly requests the full official front-matter set.

Do not reduce the line spacing of the table of contents, list of tables, or list of figures merely to fit them on one page. These list pages may continue to additional pages when needed.

## Figure And Table Layout

- Every figure or table page is isolated from body prose.
- A figure/table page may contain only the figure/table and its caption.
- Use one figure/table per page by default.
- Use two items on one page only when both are small and tightly related.
- Never allow a figure/table to be cropped, clipped, split across pages, or continued onto the next page.
- One item must be centered horizontally and vertically as a single group including the caption.
- Two items must be centered as one combined group with generous spacing between them.
- If a table cannot fit on one page after reasonable scaling, split it intentionally into separate labeled tables and put each table on its own isolated page.
- Body text must mention the figure/table before the isolated page.
- After every layout change, re-export through Microsoft Word, inspect the rendered PDF pages, and refresh the table/figure lists.

## Page-Count Control

Before drafting, ask whether the target page count is strict or flexible. If strict, tune prose density, appendix depth, and item grouping after PDF export. If flexible, preserve clean figure/table isolation and readable synthesis even when the final page count exceeds the initial target.

1. Generate the DOCX.
2. Export through Microsoft Word using `scripts/thesis_tool.py render-pdf`.
3. Count pages with `scripts/thesis_tool.py pdf-map` or `pypdf`.
4. If short:
   - expand comparison tables,
   - add literature synthesis subsections,
   - add an appendix with source-by-source evidence notes.
5. If long:
   - compress repeated prose,
   - move detailed source notes to appendix,
   - reduce image heights before deleting content.
6. Re-export until the target page count is reached or the remaining difference is explicitly reported.

## Revision Learning Loop

When the user says "next time do X", "change the workflow", or corrects the process:

1. Decide whether the correction is project-specific or reusable.
2. Project-specific: update files inside the active project folder only, such as `config/project-config.yaml`, `config/review-config.yaml`, or `ledgers/evidence-ledger.yaml`.
3. Reusable: update this file, `SKILL.md`, or a script.
4. Keep the update concise and operational.
5. If GitHub sync has been configured, run `python scripts/github_sync.py --message "<concise change summary>"`.
6. Continue the document revision after the workflow change.

## Deliverables

For each version, produce:

- `VerN.docx`
- matching `VerN.pdf`
- project-local evidence ledger (`ledgers/evidence-ledger.yaml` or JSON)
- project-local imported image ledger if source figures are used
- project-local page map JSON after PDF export

Final response should report paths, page count, verification performed, and unresolved preflight/evidence-ledger items.
For the preferred workflow, there should be no unresolved manuscript placeholders; unresolved metadata belongs in the config/evidence ledger until the user answers.
