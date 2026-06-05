---
name: docx-bibliography-crossrefs
description: Use when the user asks to modify DOCX bibliography/reference numbering, convert hand-written [1] [2] references into Word/WPS numbered references, or change正文参考文献引用/文献引用/引用编号 into clickable cross-references.
---

# DOCX Bibliography Crossrefs

## Overview

Convert DOCX bibliography entries written as literal `[1] ...`, `[2] ...` into real Word/WPS automatic numbering and replace body citations with clickable superscript bracketed `REF` cross-reference fields.

Use the bundled script for the fragile OOXML work. Do not hand-build this flow from memory.

## Workflow

1. Confirm the target file is a `.docx` and identify the bibliography heading, usually `References`.
2. Preserve safety: do not overwrite the source file. Use `--backup` and write a new output beside the input.
3. Run:

```powershell
python "<skill-dir>\scripts\bibliography_crossrefs.py" "C:\path\paper.docx" --backup --audit
```

Use `--out "C:\path\paper_crossref.docx"` when a specific output path is needed.

Body citations are superscripted by default while preserving brackets, e.g. superscript `[1]`. Use one shared bracket pair for grouped citations such as superscript `[6, 11, 12]`, not separate `[6], [11], [12]` markers. Ranged citations with hyphen, en dash, or em dash such as superscript `[37–39]` keep the visible range; each visible number is a separate `REF` field while punctuation and brackets remain superscript text.

After delivering a superscript output, mention that a normal baseline version is also available. Generate that version with `--no-superscript-body-citations` if the user asks for citations that are not superscripted.

If the heading is not `References`, pass the exact heading:

```powershell
python "<skill-dir>\scripts\bibliography_crossrefs.py" "C:\path\paper.docx" --heading "参考文献" --backup --audit
```

4. Verify the script output:
   - `reference_paragraphs_numbered` equals the number of bibliography entries.
   - `body_ref_fields_inserted` is nonzero when body citations exist.
   - `ref_fields` equals `ref_fields_with_w_h`.
   - `superscript_digit_runs` is nonzero unless `--no-superscript-body-citations` was used.
   - `hyperlink_wrappers=0`.
   - `numbering_relationship=True` and `numbering_content_type=True`.
5. Open the output in WPS/Word and check:
   - References display as automatic `[1]`, `[2]`, ...
   - Body citations display as superscript bracketed citations, unless `--no-superscript-body-citations` was used.
   - A body citation Ctrl+click jumps to the matching References entry.

## Important Details

- The correct cross-reference field is `REF Ref_Bib_001 \w \h`.
- The visible body citation should keep one shared pair of brackets as superscript text around the REF fields: superscript `[` + REF number(s) and separators + superscript `]`.
- For `[6, 11, 12]`, create three `REF` fields inside one bracket pair. For `[37–39]`, create `REF` fields for the visible `37` and `39` and keep the dash as superscript text.
- `\w` is required for "paragraph number full context/full number" behavior.
- `\h` is required for Ctrl+click jump behavior.
- Do not wrap the REF field in `w:hyperlink`. WPS may treat that as "open specified file" and show `无法打开指定的文件`.
- Add `word/numbering.xml` plus both package links when the source document did not already have numbering:
  - `word/_rels/document.xml.rels`
  - `[Content_Types].xml`
- The script only converts bibliography entries that start with continuous literal `[1]`, `[2]`, ... immediately after the bibliography heading. Stop and inspect manually if numbering is missing, duplicated, reordered, or not continuous.

## Safety Rules

- Never overwrite the original DOCX.
- Keep the backup beside the input as `<stem>.before_crossref.docx`.
- If the DOCX is open in WPS and Windows locks the output path, write a new versioned filename instead of forcing close or deleting files.
- Report exactly which final file should be used; ignore intermediate attempts.

## Verification Limits

If LibreOffice rendering is unavailable, structural audit plus WPS visual testing is acceptable. Still disclose that full PNG render QA was skipped.
