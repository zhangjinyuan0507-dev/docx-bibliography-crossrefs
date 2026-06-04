from __future__ import annotations

import argparse
import copy
import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("", REL_NS)


def qn(tag: str) -> str:
    prefix, name = tag.split(":", 1)
    if prefix == "w":
        return f"{{{W_NS}}}{name}"
    raise ValueError(tag)


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))


def next_int_attr(root: ET.Element, path: str, attr: str, start: int = 1) -> int:
    values = []
    for node in root.findall(path, NS):
        raw = node.get(qn(attr))
        if raw and raw.isdigit():
            values.append(int(raw))
    return max(values, default=start - 1) + 1


def add_numbering_def(numbering_root: ET.Element) -> str:
    abstract_num_id = str(next_int_attr(numbering_root, "w:abstractNum", "w:abstractNumId", 100))
    num_id = str(next_int_attr(numbering_root, "w:num", "w:numId", 100))

    abstract = ET.Element(qn("w:abstractNum"), {qn("w:abstractNumId"): abstract_num_id})
    lvl = ET.SubElement(abstract, qn("w:lvl"), {qn("w:ilvl"): "0"})
    ET.SubElement(lvl, qn("w:start"), {qn("w:val"): "1"})
    ET.SubElement(lvl, qn("w:numFmt"), {qn("w:val"): "decimal"})
    ET.SubElement(lvl, qn("w:lvlText"), {qn("w:val"): "[%1]"})
    ET.SubElement(lvl, qn("w:lvlJc"), {qn("w:val"): "left"})
    ppr = ET.SubElement(lvl, qn("w:pPr"))
    ET.SubElement(ppr, qn("w:ind"), {qn("w:left"): "360", qn("w:hanging"): "360"})

    num = ET.Element(qn("w:num"), {qn("w:numId"): num_id})
    ET.SubElement(num, qn("w:abstractNumId"), {qn("w:val"): abstract_num_id})

    numbering_root.append(abstract)
    numbering_root.append(num)
    return num_id


def ensure_numbering_part(files: dict[str, bytes]) -> ET.Element:
    path = "word/numbering.xml"
    if path in files:
        return ET.fromstring(files[path])
    return ET.Element(qn("w:numbering"))


def ensure_numbering_package_links(files: dict[str, bytes]) -> None:
    rels_path = "word/_rels/document.xml.rels"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    if rels_path in files:
        rels_root = ET.fromstring(files[rels_path])
    else:
        rels_root = ET.Element(f"{{{REL_NS}}}Relationships")

    has_numbering_rel = any(
        rel.get("Type") == rel_type or rel.get("Target") == "numbering.xml"
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship")
    )
    if not has_numbering_rel:
        used = []
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.get("Id", "")
            if rid.startswith("rId") and rid[3:].isdigit():
                used.append(int(rid[3:]))
        ET.SubElement(
            rels_root,
            f"{{{REL_NS}}}Relationship",
            {
                "Id": f"rId{max(used, default=0) + 1}",
                "Type": rel_type,
                "Target": "numbering.xml",
            },
        )
    files[rels_path] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    ct_path = "[Content_Types].xml"
    ct_root = ET.fromstring(files[ct_path])
    has_override = any(
        node.get("PartName") == "/word/numbering.xml"
        for node in ct_root.findall(f"{{{CT_NS}}}Override")
    )
    if not has_override:
        ET.SubElement(
            ct_root,
            f"{{{CT_NS}}}Override",
            {
                "PartName": "/word/numbering.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            },
        )
    files[ct_path] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)


def set_reference_numbering(p: ET.Element, num_id: str) -> None:
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("w:pPr"))
        p.insert(0, ppr)

    numpr = ppr.find("w:numPr", NS)
    if numpr is not None:
        ppr.remove(numpr)
    numpr = ET.Element(qn("w:numPr"))
    pstyle = ppr.find("w:pStyle", NS)
    insert_pos = list(ppr).index(pstyle) + 1 if pstyle is not None else 0
    ppr.insert(insert_pos, numpr)

    ET.SubElement(numpr, qn("w:ilvl"), {qn("w:val"): "0"})
    ET.SubElement(numpr, qn("w:numId"), {qn("w:val"): num_id})

    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:left"), "360")
    ind.set(qn("w:hanging"), "360")


def remove_leading_literal_number(p: ET.Element, number: int) -> None:
    target = f"[{number}]"
    remaining = len(target)
    seen = ""
    for t in p.findall(".//w:t", NS):
        if remaining <= 0:
            break
        text = t.text or ""
        take = min(len(text), remaining)
        seen += text[:take]
        t.text = text[take:]
        remaining -= take
    if seen != target:
        raise ValueError(f"Could not remove literal number {target!r}; saw {seen!r}")

    for t in p.findall(".//w:t", NS):
        text = t.text or ""
        if text:
            if text.startswith(" "):
                t.text = text[1:]
            return


def add_bookmark(p: ET.Element, name: str, bookmark_id: int) -> None:
    start = ET.Element(qn("w:bookmarkStart"), {qn("w:id"): str(bookmark_id), qn("w:name"): name})
    end = ET.Element(qn("w:bookmarkEnd"), {qn("w:id"): str(bookmark_id)})
    insert_at = 1 if p.find("w:pPr", NS) is not None else 0
    p.insert(insert_at, start)
    p.append(end)


def clone_rpr(template_rpr: ET.Element | None = None, superscript: bool = False) -> ET.Element | None:
    rpr = copy.deepcopy(template_rpr) if template_rpr is not None else None
    if superscript:
        if rpr is None:
            rpr = ET.Element(qn("w:rPr"))
        vert_align = rpr.find("w:vertAlign", NS)
        if vert_align is None:
            vert_align = ET.SubElement(rpr, qn("w:vertAlign"))
        vert_align.set(qn("w:val"), "superscript")
    return rpr


def make_run_with_text(text: str, template_rpr: ET.Element | None = None, superscript: bool = False) -> ET.Element:
    r = ET.Element(qn("w:r"))
    rpr = clone_rpr(template_rpr, superscript=superscript)
    if rpr is not None:
        r.append(rpr)
    t = ET.SubElement(r, qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


def make_ref_field(
    bookmark: str,
    display: str,
    template_rpr: ET.Element | None = None,
    superscript: bool = False,
) -> list[ET.Element]:
    def fld_char(kind: str) -> ET.Element:
        r = ET.Element(qn("w:r"))
        rpr = clone_rpr(template_rpr, superscript=superscript)
        if rpr is not None:
            r.append(rpr)
        ET.SubElement(r, qn("w:fldChar"), {qn("w:fldCharType"): kind})
        return r

    instr = ET.Element(qn("w:r"))
    instr_rpr = clone_rpr(template_rpr, superscript=superscript)
    if instr_rpr is not None:
        instr.append(instr_rpr)
    instr_t = ET.SubElement(instr, qn("w:instrText"))
    instr_t.set(f"{{{XML_NS}}}space", "preserve")
    instr_t.text = f" REF {bookmark} \\w \\h "

    return [
        fld_char("begin"),
        instr,
        fld_char("separate"),
        make_run_with_text(display, template_rpr, superscript=superscript),
        fld_char("end"),
    ]


BODY_REF_RE = re.compile(r"\[(\d+(?:\s*[-,;]\s*\d+)*)\]")


def replace_body_refs_in_paragraph(p: ET.Element, bookmarks: dict[int, str], superscript: bool) -> int:
    replaced = 0
    new_children: list[ET.Element] = []
    for child in list(p):
        if child.tag != qn("w:r"):
            new_children.append(child)
            continue

        text_nodes = child.findall("w:t", NS)
        if len(text_nodes) != 1:
            new_children.append(child)
            continue

        text = text_nodes[0].text or ""
        matches = list(BODY_REF_RE.finditer(text))
        if not matches:
            new_children.append(child)
            continue

        rpr = child.find("w:rPr", NS)
        pos = 0
        changed = False
        for m in matches:
            nums = [int(n) for n in re.findall(r"\d+", m.group(1))]
            if not nums or any(n not in bookmarks for n in nums):
                continue
            if m.start() > pos:
                new_children.append(make_run_with_text(text[pos : m.start()], rpr))
            new_children.append(make_run_with_text("[", rpr, superscript=superscript))
            for part in re.split(r"(\d+)", m.group(1)):
                if not part:
                    continue
                if part.isdigit() and int(part) in bookmarks:
                    new_children.extend(make_ref_field(bookmarks[int(part)], part, rpr, superscript=superscript))
                    replaced += 1
                else:
                    new_children.append(make_run_with_text(part, rpr, superscript=superscript))
            new_children.append(make_run_with_text("]", rpr, superscript=superscript))
            pos = m.end()
            changed = True
        if changed and pos < len(text):
            new_children.append(make_run_with_text(text[pos:], rpr))
        elif not changed:
            new_children.append(child)

    if replaced:
        p[:] = new_children
    return replaced


def find_references_heading(paragraphs: list[ET.Element], heading: str) -> int:
    wanted = heading.strip().lower()
    for idx, p in enumerate(paragraphs):
        if paragraph_text(p).strip().lower() == wanted:
            return idx
    raise ValueError(f"References heading not found: {heading!r}")


def patch_docx(
    src: Path,
    out: Path,
    heading: str = "References",
    superscript_body_citations: bool = True,
) -> dict[str, int]:
    with zipfile.ZipFile(src, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    document_root = ET.fromstring(files["word/document.xml"])
    body = document_root.find("w:body", NS)
    if body is None:
        raise ValueError("document body not found")
    paragraphs = body.findall("w:p", NS)

    ref_heading_idx = find_references_heading(paragraphs, heading)
    ref_paragraphs: list[tuple[int, int, ET.Element]] = []
    for idx in range(ref_heading_idx + 1, len(paragraphs)):
        text = paragraph_text(paragraphs[idx]).strip()
        match = re.match(r"^\[(\d+)\]\s+", text)
        if not match:
            break
        ref_paragraphs.append((idx, int(match.group(1)), paragraphs[idx]))

    if not ref_paragraphs:
        raise ValueError("No literal [n] reference paragraphs found after the References heading")

    actual = [number for _, number, _ in ref_paragraphs]
    expected = list(range(1, len(ref_paragraphs) + 1))
    if actual != expected:
        raise ValueError(f"Reference numbering is not continuous from 1: first/last={actual[:5]}/{actual[-5:]}")

    numbering_root = ensure_numbering_part(files)
    num_id = add_numbering_def(numbering_root)
    ensure_numbering_package_links(files)

    bookmark_ids = []
    for bm in document_root.findall(".//w:bookmarkStart", NS):
        raw = bm.get(qn("w:id"))
        if raw and raw.isdigit():
            bookmark_ids.append(int(raw))
    next_bookmark_id = max(bookmark_ids, default=0) + 1

    bookmarks: dict[int, str] = {}
    for _, number, p in ref_paragraphs:
        set_reference_numbering(p, num_id)
        remove_leading_literal_number(p, number)
        name = f"Ref_Bib_{number:03d}"
        add_bookmark(p, name, next_bookmark_id)
        next_bookmark_id += 1
        bookmarks[number] = name

    replaced = 0
    for p in paragraphs[:ref_heading_idx]:
        replaced += replace_body_refs_in_paragraph(p, bookmarks, superscript=superscript_body_citations)

    settings_path = "word/settings.xml"
    settings_root = ET.fromstring(files[settings_path]) if settings_path in files else ET.Element(qn("w:settings"))
    update_fields = settings_root.find("w:updateFields", NS)
    if update_fields is None:
        update_fields = ET.SubElement(settings_root, qn("w:updateFields"))
    update_fields.set(qn("w:val"), "true")

    files["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    files["word/numbering.xml"] = ET.tostring(numbering_root, encoding="utf-8", xml_declaration=True)
    files[settings_path] = ET.tostring(settings_root, encoding="utf-8", xml_declaration=True)

    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    return {
        "reference_paragraphs_numbered": len(ref_paragraphs),
        "body_ref_fields_inserted": replaced,
        "bibliography_bookmarks": len(ref_paragraphs),
    }


def audit_docx(path: Path) -> dict[str, int | bool]:
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        doc = ET.fromstring(zin.read("word/document.xml"))
        rels = zin.read("word/_rels/document.xml.rels").decode("utf-8")
        content_types = zin.read("[Content_Types].xml").decode("utf-8")

    instr = [t.text or "" for t in doc.findall(".//w:instrText", NS)]
    ref_result_runs = [
        run
        for run in doc.findall(".//w:r", NS)
        if run.find("w:t", NS) is not None and (run.find("w:t", NS).text or "").isdigit()
    ]
    superscript_digit_runs = [
        run
        for run in ref_result_runs
        if run.find("w:rPr/w:vertAlign", NS) is not None
        and run.find("w:rPr/w:vertAlign", NS).get(qn("w:val")) == "superscript"
    ]
    bookmarks = [
        b.get(qn("w:name")) or ""
        for b in doc.findall(".//w:bookmarkStart", NS)
        if (b.get(qn("w:name")) or "").startswith("Ref_Bib_")
    ]
    return {
        "numbering_part": "word/numbering.xml" in names,
        "numbering_relationship": "relationships/numbering" in rels,
        "numbering_content_type": "numbering+xml" in content_types,
        "bibliography_bookmarks": len(bookmarks),
        "ref_fields": sum(1 for x in instr if "REF Ref_Bib_" in x),
        "ref_fields_with_w_h": sum(1 for x in instr if "REF Ref_Bib_" in x and "\\w" in x and "\\h" in x),
        "superscript_digit_runs": len(superscript_digit_runs),
        "hyperlink_wrappers": len(doc.findall(".//w:hyperlink", NS)),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert literal DOCX bibliography citations to Word/WPS cross-references.")
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("--out", type=Path, help="Output DOCX. Default: <stem>_crossref.docx beside input.")
    parser.add_argument("--backup", action="store_true", help="Create <stem>.before_crossref.docx if it does not exist.")
    parser.add_argument("--heading", default="References", help="Bibliography heading text. Default: References.")
    parser.add_argument(
        "--no-superscript-body-citations",
        action="store_true",
        help="Do not superscript body citation markers. Default is to superscript them.",
    )
    parser.add_argument("--audit", action="store_true", help="Print structural audit after writing.")
    args = parser.parse_args(argv)

    src = args.input_docx
    if not src.exists():
        raise FileNotFoundError(src)
    out = args.out or src.with_name(f"{src.stem}_crossref.docx")
    if out.resolve() == src.resolve():
        raise ValueError("Refusing to overwrite input_docx; choose a different --out path")

    if args.backup:
        backup = src.with_name(f"{src.stem}.before_crossref.docx")
        if not backup.exists():
            shutil.copy2(src, backup)
            print(f"backup={backup}")

    stats = patch_docx(
        src,
        out,
        heading=args.heading,
        superscript_body_citations=not args.no_superscript_body_citations,
    )
    print(f"wrote={out}")
    for key, value in stats.items():
        print(f"{key}={value}")

    if args.audit:
        for key, value in audit_docx(out).items():
            print(f"{key}={value}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"error={exc}", file=sys.stderr)
        raise
