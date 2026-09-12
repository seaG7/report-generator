---
name: build-report
description: Patch title fields and replace body of a copied sample.docx, then render pages for visual QA. Use when composing the Word report after work is done, or when the user runs /build-report.
---

# Сборка report.docx

Preconditions: `title-fields.yaml` has `confirmed: true`; `work.yaml` exists; `report.md` exists; `work/checklist.md` has no unchecked `- [ ]` items; every `screenshots/` image referenced in `report.md` has `ok: true` in `manifest.json`. Do not compose if any assignment checklist item is still open.

```
python tools/patch_title.py --docx "<work>/report.docx" --title-fields "subjects/<subject>/title-fields.yaml" --student student.yaml --work "<work>/work.yaml"
python tools/replace_body.py --docx "<work>/report.docx" --md "<work>/report.md" --title-fields "subjects/<subject>/title-fields.yaml" --work-dir "<work>"
python tools/render_docx.py "<work>/report.docx" --out-dir "<work>/preview"
```

Never write `sample.docx`. Never create a Word document from scratch. After render, open every `preview/page-*.png` and inspect against the sample: title intact; body first-line indent and justify; figure captions size/bold/center like sample; images full text width; no invented list formatting. On failure, fix the copy and re-render. After 3 failed inspect loops, stop and list defects. Do not declare the report ready without a passing inspect.
