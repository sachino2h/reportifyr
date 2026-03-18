import argparse
import copy
import json
import zipfile
import sys

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _visible_text(el, ns):
    return "".join(el.xpath(".//w:t/text()", namespaces=ns)).strip()


def _get_or_create_child(parent, local_name, prepend=False):
    child = parent.find(f"./{_w_tag(local_name)}")
    if child is not None:
        return child

    child = etree.Element(_w_tag(local_name))
    if prepend:
        parent.insert(0, child)
    else:
        parent.append(child)

    return child


def _w_tag(local_name):
    return f"{{{W_NS}}}{local_name}"


def _w_attr(local_name):
    return f"{{{W_NS}}}{local_name}"


def _parse_int(value):
    if value is None:
        return None

    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _get_body_section_props(body, table_pos):
    sect_pr = None

    for el in body[: table_pos + 1]:
        if el.tag != _w_tag("p"):
            continue

        para_sect_pr = el.find(f"./{_w_tag('pPr')}/{_w_tag('sectPr')}")
        if para_sect_pr is not None:
            sect_pr = para_sect_pr

    if sect_pr is None:
        sect_pr = body.find(f"./{_w_tag('sectPr')}")

    return sect_pr


def _get_available_width_twips(sect_pr):
    if sect_pr is None:
        return None

    pg_sz = sect_pr.find(f"./{_w_tag('pgSz')}")
    pg_mar = sect_pr.find(f"./{_w_tag('pgMar')}")
    if pg_sz is None or pg_mar is None:
        return None

    page_width = _parse_int(pg_sz.get(_w_attr("w")))
    margin_left = _parse_int(pg_mar.get(_w_attr("left")))
    margin_right = _parse_int(pg_mar.get(_w_attr("right")))
    if None in (page_width, margin_left, margin_right):
        return None

    available_width = page_width - margin_left - margin_right
    if available_width <= 0:
        return None

    return available_width


def _get_grid_width_twips(table_el):
    grid_cols = table_el.findall(f"./{_w_tag('tblGrid')}/{_w_tag('gridCol')}")
    widths = [_parse_int(col.get(_w_attr("w"))) for col in grid_cols]
    widths = [width for width in widths if width is not None]

    if not widths:
        return None

    total_width = sum(widths)
    if total_width <= 0:
        return None

    return total_width


def _get_or_create_tbl_pr(table_el):
    return _get_or_create_child(table_el, "tblPr", prepend=True)


def _set_fixed_layout(table_el, width_twips):
    if width_twips is None or width_twips <= 0:
        return

    tbl_pr = _get_or_create_tbl_pr(table_el)

    tbl_w = tbl_pr.find(f"./{_w_tag('tblW')}")
    if tbl_w is None:
        tbl_w = etree.Element(_w_tag("tblW"))
        tbl_pr.insert(0, tbl_w)

    tbl_w.set(_w_attr("type"), "dxa")
    tbl_w.set(_w_attr("w"), str(width_twips))

    tbl_layout = tbl_pr.find(f"./{_w_tag('tblLayout')}")
    if tbl_layout is None:
        tbl_layout = etree.Element(_w_tag("tblLayout"))
        tbl_pr.append(tbl_layout)

    tbl_layout.set(_w_attr("type"), "fixed")


def _normalize_table_width(table_el, src_body, source_table_pos):
    tbl_pr = table_el.find(f"./{_w_tag('tblPr')}")
    if tbl_pr is None:
        return

    tbl_w = tbl_pr.find(f"./{_w_tag('tblW')}")
    if tbl_w is None:
        return

    width_type = tbl_w.get(_w_attr("type"))
    width_value = _parse_int(tbl_w.get(_w_attr("w")))

    if width_type == "dxa" and width_value is not None and width_value > 0:
        _set_fixed_layout(table_el, width_value)
        return

    source_sect_pr = _get_body_section_props(src_body, source_table_pos)
    available_width = _get_available_width_twips(source_sect_pr)
    if available_width is None:
        return

    fixed_width = None
    if width_type == "pct" and width_value is not None:
        fixed_width = round((available_width * width_value) / 5000)
    elif width_type == "auto":
        fixed_width = _get_grid_width_twips(table_el)

    if fixed_width is None or fixed_width <= 0:
        return

    _set_fixed_layout(table_el, fixed_width)


def _parse_style_json(style_json):
    if style_json is None:
        return None

    try:
        style = json.loads(style_json)
    except json.JSONDecodeError as exc:
        print(f"ERROR: invalid --style-json value: {exc}", file=sys.stderr)
        return None

    if not isinstance(style, dict):
        print("ERROR: --style-json must decode to an object", file=sys.stderr)
        return None

    return style


def _set_run_font(run_el, font_family=None, font_size=None, bold=None):
    if font_family is None and font_size is None and bold is None:
        return

    r_pr = _get_or_create_child(run_el, "rPr", prepend=True)

    if font_family is not None:
        r_fonts = _get_or_create_child(r_pr, "rFonts")
        for attr_name in ("ascii", "hAnsi", "cs", "eastAsia"):
            r_fonts.set(_w_attr(attr_name), str(font_family))

    if font_size is not None:
        font_size_val = str(int(round(float(font_size) * 2)))
        for tag_name in ("sz", "szCs"):
            size_el = _get_or_create_child(r_pr, tag_name)
            size_el.set(_w_attr("val"), font_size_val)

    if bold is not None:
        bold_val = "1" if bold else "0"
        for tag_name in ("b", "bCs"):
            bold_el = _get_or_create_child(r_pr, tag_name)
            bold_el.set(_w_attr("val"), bold_val)


def _set_cell_fill(cell_el, fill):
    tc_pr = _get_or_create_child(cell_el, "tcPr", prepend=True)
    shd = _get_or_create_child(tc_pr, "shd")
    shd.set(_w_attr("val"), "clear")
    shd.set(_w_attr("color"), "auto")
    shd.set(_w_attr("fill"), fill)


def _apply_table_style(table_el, table_style):
    if not table_style:
        return

    font_family = table_style.get("font_family")
    font_size = table_style.get("font_size")
    header_fill = table_style.get("header_fill")
    header_bold = table_style.get("header_bold")
    header_rows = int(table_style.get("header_rows", 1) or 0)

    rows = table_el.findall(f"./{_w_tag('tr')}")
    for row_index, row_el in enumerate(rows):
        is_header_row = row_index < header_rows

        for cell_el in row_el.findall(f"./{_w_tag('tc')}"):
            if is_header_row and header_fill:
                _set_cell_fill(cell_el, header_fill)

            for run_el in cell_el.xpath(".//w:r", namespaces=NS):
                _set_run_font(
                    run_el,
                    font_family=font_family,
                    font_size=font_size,
                    bold=header_bold if is_header_row else None,
                )


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
    parser.add_argument(
        "--style-json",
        default=None,
        help="Optional JSON object describing DOCX table styling overrides.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.table_index < 1:
        print("ERROR: --table-index must be >= 1", file=sys.stderr)
        return 2

    w_ns = W_NS
    ns = NS
    table_style = _parse_style_json(args.style_json)
    if args.style_json is not None and table_style is None:
        return 2

    with zipfile.ZipFile(args.source, "r") as src_zip:
        try:
            src_xml = src_zip.read("word/document.xml")
        except KeyError:
            print("ERROR: source DOCX missing word/document.xml", file=sys.stderr)
            return 2

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
    _normalize_table_width(table_el, src_body, source_table_pos)
    _apply_table_style(table_el, table_style)

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

    with zipfile.ZipFile(args.output, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, data in tpl_entries.items():
            out_zip.writestr(name, data)

    print(f"Done: table inserted into {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
