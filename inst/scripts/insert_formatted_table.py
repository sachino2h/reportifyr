import argparse
import copy
import zipfile
import sys

from lxml import etree


def _visible_text(el, ns):
    return "".join(el.xpath(".//w:t/text()", namespaces=ns)).strip()


def _collect_used_style_ids(elements, ns):
    style_ids = set()
    for el in elements:
        style_ids.update(el.xpath(".//w:pStyle/@w:val", namespaces=ns))
        style_ids.update(el.xpath(".//w:rStyle/@w:val", namespaces=ns))
        style_ids.update(el.xpath(".//w:tblStyle/@w:val", namespaces=ns))
    return {sid for sid in style_ids if sid}


def _add_style_and_dependencies(style_id, src_styles_root, tpl_styles_root, ns):
    style_xpath = f"./w:style[@w:styleId='{style_id}']"
    if tpl_styles_root.xpath(style_xpath, namespaces=ns):
        return

    src_style = src_styles_root.xpath(style_xpath, namespaces=ns)
    if not src_style:
        return

    src_style = src_style[0]
    dep_ids = src_style.xpath(
        "./w:basedOn/@w:val | ./w:next/@w:val | ./w:link/@w:val",
        namespaces=ns,
    )
    for dep_id in dep_ids:
        if dep_id:
            _add_style_and_dependencies(dep_id, src_styles_root, tpl_styles_root, ns)

    tpl_styles_root.append(copy.deepcopy(src_style))


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Copy a table from a source DOCX into a template DOCX by replacing "
            "a placeholder paragraph."
        )
    )
    parser.add_argument("-s", "--source", required=True, help="Source DOCX path")
    parser.add_argument(
        "-t", "--template", required=True, help="Template DOCX path"
    )
    parser.add_argument("-o", "--output", required=True, help="Output DOCX path")
    parser.add_argument(
        "-p",
        "--placeholder",
        required=True,
        help="Placeholder text to find and replace",
    )
    parser.add_argument(
        "--table-index",
        type=int,
        default=1,
        help="1-based table index from source DOCX (default: 1)",
    )
    parser.add_argument(
        "--include-footnote",
        action="store_true",
        help=(
            "If set, copy the first non-empty paragraph after the selected "
            "source table and insert it after the table."
        ),
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.table_index < 1:
        print("ERROR: --table-index must be >= 1", file=sys.stderr)
        return 2

    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w_ns}

    with zipfile.ZipFile(args.source, "r") as src_zip:
        try:
            src_xml = src_zip.read("word/document.xml")
        except KeyError:
            print("ERROR: source DOCX missing word/document.xml", file=sys.stderr)
            return 2
        src_styles_xml = src_zip.read("word/styles.xml") if "word/styles.xml" in src_zip.namelist() else None

    src_root = etree.fromstring(src_xml)
    src_body = src_root.xpath("//w:body", namespaces=ns)
    if not src_body:
        print("ERROR: source DOCX has invalid document.xml (missing w:body)", file=sys.stderr)
        return 2
    src_body = src_body[0]

    src_siblings = src_body.xpath("./*", namespaces=ns)
    src_table_positions = [
        idx for idx, el in enumerate(src_siblings) if el.tag == f"{{{w_ns}}}tbl"
    ]
    if len(src_table_positions) < args.table_index:
        print(
            (
                "ERROR: source document has "
                f"{len(src_table_positions)} table(s), but --table-index={args.table_index}"
            ),
            file=sys.stderr,
        )
        return 2

    source_table_pos = src_table_positions[args.table_index - 1]
    table_el = copy.deepcopy(src_siblings[source_table_pos])

    footnote_els = []
    if args.include_footnote:
        footnotes_started = False
        empty_after_start = 0
        for el in src_siblings[source_table_pos + 1 :]:
            # Stop once another table starts.
            if el.tag == f"{{{w_ns}}}tbl":
                break

            el_text = _visible_text(el, ns)
            has_visible_text = bool(el_text)

            if has_visible_text:
                footnotes_started = True
                empty_after_start = 0
                footnote_els.append(copy.deepcopy(el))
            elif footnotes_started:
                # Allow a few blank paragraph lines within notes, but stop
                # once we hit a clear break.
                if el.tag == f"{{{w_ns}}}p":
                    empty_after_start += 1
                    if empty_after_start >= 3:
                        break

    with zipfile.ZipFile(args.template, "r") as tpl_zip:
        try:
            tpl_xml = tpl_zip.read("word/document.xml")
        except KeyError:
            print("ERROR: template DOCX missing word/document.xml", file=sys.stderr)
            return 2
        tpl_entries = {name: tpl_zip.read(name) for name in tpl_zip.namelist()}

    tpl_root = etree.fromstring(tpl_xml)
    body = tpl_root.xpath("//w:body", namespaces=ns)
    if not body:
        print("ERROR: template DOCX has invalid document.xml (missing w:body)", file=sys.stderr)
        return 2
    body = body[0]

    target_para = None
    for para in body.xpath("./w:p", namespaces=ns):
        para_text = "".join(para.xpath(".//w:t/text()", namespaces=ns))
        if args.placeholder in para_text:
            target_para = para
            break

    if target_para is None:
        print(
            f"ERROR: placeholder not found: {args.placeholder}",
            file=sys.stderr,
        )
        return 2

    target_para.addprevious(table_el)
    if footnote_els:
        anchor = table_el
        for footnote_el in footnote_els:
            anchor.addnext(footnote_el)
            anchor = footnote_el
    target_para.getparent().remove(target_para)

    tpl_entries["word/document.xml"] = etree.tostring(
        tpl_root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes",
    )

    # Preserve table/footnote look by copying missing style definitions from
    # source DOCX into template DOCX for any styles used by inserted elements.
    tpl_styles_xml = tpl_entries.get("word/styles.xml")
    if src_styles_xml is not None and tpl_styles_xml is not None:
        src_styles_root = etree.fromstring(src_styles_xml)
        tpl_styles_root = etree.fromstring(tpl_styles_xml)
        inserted_elements = [table_el] + footnote_els
        used_style_ids = _collect_used_style_ids(inserted_elements, ns)
        for style_id in sorted(used_style_ids):
            _add_style_and_dependencies(
                style_id, src_styles_root, tpl_styles_root, ns
            )
        tpl_entries["word/styles.xml"] = etree.tostring(
            tpl_styles_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

    with zipfile.ZipFile(args.output, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, data in tpl_entries.items():
            out_zip.writestr(name, data)

    print(f"Done: table inserted into {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
