# SNU Thesis Production Rules

## Priority

1. User's newest explicit instruction
2. Official SNU thesis-writing rules
3. Project config
4. Existing document style
5. General academic style

## Document Order

Use the SNU order unless the user explicitly removes optional pages:

외표지, 간지, 속표지, 인준지, 국문초록, 목차, 표 목차, 그림 목차, 본문, 참고문헌, 부록, Abstract, 감사문, 간지, 뒷표지.

## Front Matter Template

For the cover/title/approval pages, preserve the locked template layout. Change only metadata slots: degree label, Korean/English title lines, degree-award month, school, department, advisor, author, submission/approval months, and committee names. Do not restyle, reflow, or redesign these pages unless the user explicitly requests it.

## Scientific Writing Guardrails

- Preserve research meaning and all supplied evidence, data, quotations, names, dates, values, and references.
- Do not introduce unprovided data, sources, measurements, cases, calculations, quotations, or references.
- Use objective thesis prose.
- Avoid question-style, oral-presentation, or casual phrasing.
- Separate directly established evidence, calculation/analysis results, and interpretation.
- Use `[확인 필요]` for uncertain content.

## Figures and Tables

- Mention every figure/table in body text before insertion.
- Put every figure/table on an isolated page unless the user explicitly requests another layout.
- Do not let a figure/table clip, crop, overflow, split, or continue onto the next page.
- Center one figure/table as one visual group on the page; when two small related items share a page, center the combined group with clear spacing.
- Keep Figure/Table numbering sequential.
- Captions should identify what is shown, not repeat full interpretation.
- When moving/deleting/adding a figure, update:
  - body references
  - caption number
  - list of figures entry
  - actual PDF page number
  - surrounding prose
- Preserve user-locked rotations and layouts.

## Render QA

Always create a Microsoft Word-exported PDF for each DOCX version. Word pagination is authoritative; do not use LibreOffice, Pandoc, Preview, or a generic renderer for final PDFs unless the user explicitly approves a fallback. Inspect:

- front matter pages
- TOC/list pages
- changed figure/table pages
- page before and after moved content
- final page count

Do not trust DOCX XML or non-Word conversion output alone. Use the Microsoft Word-exported PDF and its rendered page images.
