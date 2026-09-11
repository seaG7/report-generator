---
name: pdf-to-md
description: Convert an assignment PDF to Markdown with images in document order. Use when ingesting a lab PDF, converting PDF to md, or when /report needs assignment.md. Use when the user runs /pdf-to-md.
---

# PDF → Markdown

From repo root:

```
python tools/pdf_to_md.py "<path-to.pdf>" --out "subjects/<subject>/lab-NN-<slug>"
```

Read `assignment.md`. Do not invent figures that were not extracted. If the command exits 1 with «Нужен текстовый PDF или Marker», stop and tell the user.
