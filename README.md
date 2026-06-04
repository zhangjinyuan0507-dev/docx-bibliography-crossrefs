# DOCX Bibliography Crossrefs

Codex skill and utility script for converting hand-written DOCX bibliography citations into real Word/WPS cross-references.

It is designed for papers whose bibliography entries look like:

```text
References
[1] Author A. Title...
[2] Author B. Title...
```

The script converts the bibliography into Word/WPS automatic numbering (`[1]`, `[2]`, ...) and replaces body citations such as `[1]`, `[1, 7, 15]`, or `[62-65]` with `REF` fields that use:

```text
REF Ref_Bib_001 \w \h
```

That means:

- `\w`: use the full paragraph number, so the visible citation remains `[1]`.
- `\h`: make the cross-reference clickable with Ctrl+click in Word/WPS.

## What It Fixes

Plain text bibliography citations cannot update or jump to the reference list. This skill makes them behave like real cross-references while preserving the familiar bracketed citation style.

## Usage

Run the script on a `.docx` file:

```powershell
python scripts\bibliography_crossrefs.py "C:\path\paper.docx" --backup --audit
```

By default, the output is written beside the input as:

```text
paper_crossref.docx
```

Use a custom output path:

```powershell
python scripts\bibliography_crossrefs.py "C:\path\paper.docx" --out "C:\path\paper_crossref.docx" --backup --audit
```

If the bibliography heading is not `References`, pass it explicitly:

```powershell
python scripts\bibliography_crossrefs.py "C:\path\paper.docx" --heading "参考文献" --backup --audit
```

## Expected Input

- DOCX file.
- Bibliography heading defaults to `References`.
- Bibliography entries immediately after the heading must start with continuous literal numbers:
  - `[1] ...`
  - `[2] ...`
  - `[3] ...`
- Body citations should use bracketed numeric citations, for example:
  - `[1]`
  - `[1, 7, 15, 24]`
  - `[62-65]`

If the reference list is not continuous or uses a different format, inspect manually before running.

## Verification

The `--audit` output should show:

```text
numbering_part=True
numbering_relationship=True
numbering_content_type=True
ref_fields=<nonzero>
ref_fields_with_w_h=<same as ref_fields>
hyperlink_wrappers=0
```

Then open the output in Word or WPS and test Ctrl+click on a body citation. It should jump to the matching bibliography entry.

## Codex Skill

The repository is also a Codex skill. Install it under your Codex skills directory, for example:

```text
~/.codex/skills/docx-bibliography-crossrefs
```

Then ask Codex things like:

- "修改一下参考文献编号以及正文参考文献引用"
- "把正文的参考文献引用改成交叉引用"
- "让 DOCX 的文献编号可以 Ctrl 点击跳转"

Codex should use the bundled script instead of manually editing fragile OOXML.

## Safety

- The script refuses to overwrite the input file.
- Use `--backup` to create `<stem>.before_crossref.docx`.
- If Word/WPS has the output open and locks it, write a new versioned filename.

## Known Notes

Do not wrap the `REF` field in a separate OOXML `w:hyperlink`. WPS may interpret that as a file hyperlink and show `无法打开指定的文件`. The correct jump behavior comes from the `\h` switch inside the `REF` field.
