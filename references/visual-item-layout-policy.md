# Figure And Table Page Layout Policy

Use this reference whenever a thesis/review manuscript contains figure pages, table pages, source-paper images, generated schematics, or large comparison tables.

## Non-Negotiable Rules

1. A figure/table page may contain only the visual item(s) and caption(s). No body prose, interpretation paragraph, bullet list, or transition sentence may appear on that page.
2. A figure or table must never be clipped, cropped, split across pages, or allowed to continue onto the next page.
3. One item per page is the default. Two items may share one page only when both are small and tightly related.
4. One item must be centered horizontally and vertically as a single visual group, including its caption.
5. Two items must be treated as one balanced group: keep generous spacing between them, align their visual centers, and center the combined group on the page.
6. If an item does not fit, reduce size, font size, row spacing, or column width while preserving readability. If it still cannot fit, split it intentionally into separate labeled items, such as Table 3A/Table 3B, and place each on its own page.
7. Do not crop a figure to force fit. Scaling down is allowed; destructive cropping is not.

## DOCX Construction Rules

- Insert a hard page break before and after every isolated visual page.
- Keep the item and its caption together with Word paragraph/table keep rules.
- Use centered paragraph alignment for image paragraphs and captions.
- For one item, build a single centered block and place it near the physical center of the text area.
- For two items, use either:
  - a stacked layout with a clear vertical gap when items are wide, or
  - a two-column borderless table when items are narrow.
- The two-item group must be centered as a group, not merely each item independently.
- For tables, disable row splitting across pages. In OOXML this means adding `w:cantSplit` to table rows when possible.
- Avoid floating anchors for core figures/tables unless an existing template requires them. Inline or table-contained items are easier to keep stable across Word-to-PDF export.

## Sizing Rules

Use the page text area, not the physical paper edge, as the available box. Reserve space for caption(s) and a small breathing margin.

Recommended starting limits:

- one image: maximum width 95% of text width; maximum height 75% of text height before caption
- one table: maximum width 100% of text width; maximum height 80% of text height including caption
- two stacked items: each item maximum height 38% of text height, with at least 12 pt gap
- two side-by-side items: each item maximum width 47% of text width, with at least 12 pt gutter

If the page uses page numbers or running headers, keep the visual group clear of those areas.

## PDF QA Rules

After Microsoft Word PDF export:

1. Rasterize each changed figure/table page.
2. Confirm the page contains only item(s), caption(s), and normal page furniture.
3. Confirm no item touches or crosses the content boundary.
4. Confirm there is no repeated header row or orphaned table continuation on the next page unless it was intentionally split and renamed.
5. Confirm visual centering from the rendered PDF, not from DOCX assumptions.
6. For one item, the detected content-group center should be close to the page center.
7. For two items, the combined detected content-group center should be close to the page center and the gap between items should look intentional.

Use `scripts/visual_page_qa.py` on rendered page PNGs when available. Treat it as a screening tool; final acceptance still requires visual inspection of the Word-exported PDF pages.

