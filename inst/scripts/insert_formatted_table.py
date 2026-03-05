import argparse
import copy
import zipfile
import sys

from lxml import etree


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

    src_root = etree.fromstring(src_xml)
    src_tables = src_root.xpath("//w:body/w:tbl", namespaces=ns)
    if len(src_tables) < args.table_index:
        print(
            (
                "ERROR: source document has "
                f"{len(src_tables)} table(s), but --table-index={args.table_index}"
            ),
            file=sys.stderr,
        )
        return 2

    table_el = copy.deepcopy(src_tables[args.table_index - 1])

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
    target_para.getparent().remove(target_para)

    tpl_entries["word/document.xml"] = etree.tostring(
        tpl_root,
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
