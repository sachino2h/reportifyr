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


def _clear_run_properties(run_el):
    r_pr = run_el.find(f"./{_w_tag('rPr')}")
    if r_pr is None:
        return
    run_el.remove(r_pr)


def _set_cell_fill(cell_el, fill):
    tc_pr = _get_or_create_child(cell_el, "tcPr", prepend=True)
    shd = _get_or_create_child(tc_pr, "shd")
    shd.set(_w_attr("val"), "clear")
    shd.set(_w_attr("color"), "auto")
    shd.set(_w_attr("fill"), fill)


def _set_paragraph_style(para_el, style_id, clear_direct_formatting=False):
    if not style_id:
        return

    if clear_direct_formatting:
        old_p_pr = para_el.find(f"./{_w_tag('pPr')}")
        if old_p_pr is not None:
            para_el.remove(old_p_pr)

    p_pr = _get_or_create_child(para_el, "pPr", prepend=True)
    p_style = _get_or_create_child(p_pr, "pStyle", prepend=True)
    p_style.set(_w_attr("val"), style_id)


def _apply_cell_paragraph_style(cell_el, style_id):
    if not style_id:
        return

    for para_el in cell_el.findall(f"./{_w_tag('p')}"):
        _set_paragraph_style(
            para_el,
            style_id,
            clear_direct_formatting=True,
        )


def _get_table_col_count(table_el):
    return len(table_el.findall(f"./{_w_tag('tblGrid')}/{_w_tag('gridCol')}"))


def _row_total_span(row_el):
    total = 0
    for cell_el in row_el.findall(f"./{_w_tag('tc')}"):
        tc_pr = cell_el.find(f"./{_w_tag('tcPr')}")
        if tc_pr is None:
            total += 1
            continue

        grid_span = tc_pr.find(f"./{_w_tag('gridSpan')}")
        if grid_span is None:
            total += 1
            continue

        span_val = _parse_int(grid_span.get(_w_attr("val")))
        total += span_val if span_val is not None and span_val > 0 else 1
    return total


def _row_has_tbl_header(row_el):
    return row_el.find(f"./{_w_tag('trPr')}/{_w_tag('tblHeader')}") is not None


def _compute_row_meta(rows, n_cols, header_rows):
    row_meta = []
    for row_index, row_el in enumerate(rows):
        is_header_by_marker = _row_has_tbl_header(row_el)
        is_header_row = is_header_by_marker or row_index < header_rows
        total_span = _row_total_span(row_el)
        n_cells = len(row_el.findall(f"./{_w_tag('tc')}"))
        is_full_width = n_cells == 1 or (n_cols > 0 and total_span == n_cols)
        row_meta.append(
            {
                "row_index": row_index,
                "is_header_row": is_header_row,
                "is_full_width": is_full_width,
            }
        )

    data_row_indices = [
        r["row_index"] for r in row_meta
        if not r["is_header_row"] and not r["is_full_width"]
    ]
    last_data_row = max(data_row_indices) if data_row_indices else -1
    return row_meta, last_data_row


def _extract_style_lookup(styles_xml):
    lookup = {}
    if not styles_xml:
        return lookup

    try:
        root = etree.fromstring(styles_xml)
    except etree.XMLSyntaxError:
        return lookup

    for style_el in root.findall(f".//{_w_tag('style')}"):
        if style_el.get(_w_attr("type")) != "paragraph":
            continue
        style_id = style_el.get(_w_attr("styleId"))
        if not style_id:
            continue

        lookup[style_id.strip().lower()] = style_id
        name_el = style_el.find(f"./{_w_tag('name')}")
        if name_el is not None:
            name_val = name_el.get(_w_attr("val"))
            if name_val and name_val.strip():
                lookup[name_val.strip().lower()] = style_id

    return lookup


def _resolve_paragraph_style_id(style_value, style_lookup):
    if style_value is None:
        return None
    value = str(style_value).strip()
    if not value:
        return None
    return style_lookup.get(value.lower(), value)


def _apply_table_style(table_el, table_style, style_lookup=None):
    if not table_style:
        return
    if style_lookup is None:
        style_lookup = {}

    font_family = table_style.get("font_family")
    font_size = table_style.get("font_size")
    header_fill = table_style.get("header_fill")
    header_bold = table_style.get("header_bold")
    header_rows = int(table_style.get("header_rows", 1) or 0)
    header_para_style = _resolve_paragraph_style_id(
        table_style.get("header_paragraph_style"),
        style_lookup,
    )
    split_para_style = _resolve_paragraph_style_id(
        table_style.get("split_row_paragraph_style"),
        style_lookup,
    )
    first_col_para_style = _resolve_paragraph_style_id(
        table_style.get("first_column_paragraph_style"),
        style_lookup,
    )
    body_para_style = _resolve_paragraph_style_id(
        table_style.get("body_paragraph_style"),
        style_lookup,
    )
    footnote_para_style = _resolve_paragraph_style_id(
        table_style.get("footnote_paragraph_style"),
        style_lookup,
    )

    rows = table_el.findall(f"./{_w_tag('tr')}")
    n_cols = _get_table_col_count(table_el)
    row_meta, last_data_row = _compute_row_meta(rows, n_cols, header_rows)

    for row_index, row_el in enumerate(rows):
        meta = row_meta[row_index]
        is_header_row = meta["is_header_row"]
        is_full_width = meta["is_full_width"]

        if is_header_row:
            row_para_style = header_para_style
        elif is_full_width:
            row_para_style = footnote_para_style if row_index > last_data_row else split_para_style
        else:
            row_para_style = None

        cells = row_el.findall(f"./{_w_tag('tc')}")
        for cell_index, cell_el in enumerate(cells):
            if is_header_row and header_fill:
                _set_cell_fill(cell_el, header_fill)

            applied_para_style = False
            if row_para_style:
                _apply_cell_paragraph_style(cell_el, row_para_style)
                applied_para_style = True
            elif (not is_header_row) and (not is_full_width):
                if cell_index == 0:
                    _apply_cell_paragraph_style(cell_el, first_col_para_style)
                    applied_para_style = bool(first_col_para_style)
                else:
                    _apply_cell_paragraph_style(cell_el, body_para_style)
                    applied_para_style = bool(body_para_style)

            for run_el in cell_el.xpath(".//w:r", namespaces=NS):
                if applied_para_style:
                    # Ensure paragraph style is not shown with direct run-font
                    # overrides such as "+ Calibri" in Word style inspector.
                    _clear_run_properties(run_el)
                _set_run_font(
                    run_el,
                    font_family=font_family,
                    font_size=font_size,
                    bold=header_bold if is_header_row else None,
                )


def _extract_table_footnote_paragraphs(table_el, table_style, style_lookup=None):
    if not table_style or not bool(table_style.get("extract_table_footnotes", False)):
        return []

    if style_lookup is None:
        style_lookup = {}

    header_rows = int(table_style.get("header_rows", 1) or 0)
    footnote_para_style = _resolve_paragraph_style_id(
        table_style.get("footnote_paragraph_style"),
        style_lookup,
    )

    rows = table_el.findall(f"./{_w_tag('tr')}")
    n_cols = _get_table_col_count(table_el)
    row_meta, last_data_row = _compute_row_meta(rows, n_cols, header_rows)

    footnote_rows = [
        rows[m["row_index"]]
        for m in row_meta
        if (not m["is_header_row"]) and m["is_full_width"] and m["row_index"] > last_data_row
    ]
    if not footnote_rows:
        return []

    out_paragraphs = []
    for row_el in footnote_rows:
        cells = row_el.findall(f"./{_w_tag('tc')}")
        if not cells:
            continue
        foot_cell = cells[0]
        for para_el in foot_cell.findall(f"./{_w_tag('p')}"):
            cloned_para = copy.deepcopy(para_el)
            if footnote_para_style:
                _set_paragraph_style(
                    cloned_para,
                    footnote_para_style,
                    clear_direct_formatting=True,
                )
                for run_el in cloned_para.xpath(".//w:r", namespaces=NS):
                    _clear_run_properties(run_el)
            out_paragraphs.append(cloned_para)

    for row_el in footnote_rows:
        parent = row_el.getparent()
        if parent is not None:
            parent.remove(row_el)

    return out_paragraphs


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
    style_lookup = _extract_style_lookup(tpl_entries.get("word/styles.xml"))
    extracted_table_footnotes = _extract_table_footnote_paragraphs(
        table_el,
        table_style,
        style_lookup=style_lookup,
    )
    _apply_table_style(table_el, table_style, style_lookup=style_lookup)

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
    anchor = table_el
    if extracted_table_footnotes:
        for para_el in extracted_table_footnotes:
            anchor.addnext(para_el)
            anchor = para_el
    if footnote_els:
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
