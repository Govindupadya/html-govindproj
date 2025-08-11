import os
import time
import zipfile
from datetime import date

OUTPUT_PATH = "/workspace/AHT_Productivity_Dashboard.xlsx"


def xml_header():
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"


def build_content_types():
    return xml_header() + (
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
        "</Types>"
    )


def build_root_rels():
    return xml_header() + (
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
        "</Relationships>"
    )


def build_app_props():
    return xml_header() + (
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>Python</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<Company></Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>16.0300</AppVersion>"
        "</Properties>"
    )


def build_core_props():
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    return xml_header() + (
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        "<dc:creator>Generator</dc:creator>"
        f"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>"
        "<cp:revision>1</cp:revision>"
        "</cp:coreProperties>"
    )


def build_workbook():
    return xml_header() + (
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets>"
        "<sheet name=\"RawData\" sheetId=\"1\" r:id=\"rId1\"/>"
        "<sheet name=\"Dashboard\" sheetId=\"2\" r:id=\"rId2\"/>"
        "</sheets>"
        "</workbook>"
    )


def build_workbook_rels():
    return xml_header() + (
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        "</Relationships>"
    )


def build_styles():
    return xml_header() + (
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
        "<fills count=\"2\">"
        "<fill><patternFill patternType=\"none\"/></fill>"
        "<fill><patternFill patternType=\"gray125\"/></fill>"
        "</fills>"
        "<borders count=\"1\"><border/></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "</styleSheet>"
    )


def cell_inline_str(ref: str, text: str) -> str:
    text = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f"<c r=\"{ref}\" t=\"inlineStr\"><is><t>{text}</t></is></c>"


def cell_number(ref: str, number) -> str:
    return f"<c r=\"{ref}\" t=\"n\"><v>{number}</v></c>"


def cell_formula(ref: str, formula: str) -> str:
    formula = formula.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f"<c r=\"{ref}\" t=\"n\"><f>{formula}</f></c>"


def build_sheet1_rawdata():
    # Headers: Date, Agent, Process, Allocated, Processed, AHT (mins)
    today = date.today().isoformat()
    rows = []
    # Row 1 headers
    headers = [
        cell_inline_str("A1", "Date"),
        cell_inline_str("B1", "Agent"),
        cell_inline_str("C1", "Process"),
        cell_inline_str("D1", "Allocated"),
        cell_inline_str("E1", "Processed"),
        cell_inline_str("F1", "AHT (mins)"),
    ]
    rows.append(f"<row r=\"1\">{''.join(headers)}</row>")

    # Sample data rows (2-5)
    samples = [
        (today, "Agent 1", "Process A", 100, 90, 4.5),
        (today, "Agent 1", "Process B", 80, 65, 5.2),
        (today, "Agent 2", "Process A", 120, 110, 4.2),
        (today, "Agent 2", "Process C", 75, 60, 6.1),
    ]
    for idx, (d, agent, proc, alloc, procd, aht) in enumerate(samples, start=2):
        row_cells = [
            cell_inline_str(f"A{idx}", d),
            cell_inline_str(f"B{idx}", agent),
            cell_inline_str(f"C{idx}", proc),
            cell_number(f"D{idx}", alloc),
            cell_number(f"E{idx}", procd),
            cell_number(f"F{idx}", aht),
        ]
        rows.append(f"<row r=\"{idx}\">{''.join(row_cells)}</row>")

    sheet_data = "".join(rows)

    return xml_header() + (
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        f"<dimension ref=\"A1:F1000\"/>"
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/>"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def build_sheet2_dashboard():
    rows = []

    # Title row
    rows.append(f"<row r=\"1\">{cell_inline_str('A1', 'AHT Productivity Dashboard')}</row>")

    # KPIs
    rows.append(
        "<row r=\"3\">"
        f"{cell_inline_str('A3', 'Total Allocated')}"
        f"{cell_formula('B3', 'SUM(RawData!D:D)')}"
        f"{cell_inline_str('D3', 'Overall Productivity%')}"
        f"{cell_formula('E3', 'IFERROR(B4/B3,0)')}"
        "</row>"
    )

    rows.append(
        "<row r=\"4\">"
        f"{cell_inline_str('A4', 'Total Processed')}"
        f"{cell_formula('B4', 'SUM(RawData!E:E)')}"
        f"{cell_inline_str('D4', 'Overall Avg AHT (mins)')}"
        f"{cell_formula('E4', 'IFERROR(AVERAGE(RawData!F:F),0)')}"
        "</row>"
    )

    # Table headers at row 7
    rows.append(
        "<row r=\"7\">"
        f"{cell_inline_str('A7', 'Process')}"
        f"{cell_inline_str('B7', 'Allocated')}"
        f"{cell_inline_str('C7', 'Processed')}"
        f"{cell_inline_str('D7', 'Productivity%')}"
        f"{cell_inline_str('E7', 'Average AHT')}"
        "</row>"
    )

    # Data rows 8..107 (user fills A; formulas in B..E)
    for i in range(8, 108):
        proc_ref = f"A{i}"
        rows.append(
            f"<row r=\"{i}\">"
            f"{cell_inline_str(proc_ref, '')}"
            f"{cell_formula(f'B{i}', f'IF({proc_ref}<>"" , SUMIF(RawData!C:C, {proc_ref}, RawData!D:D), "")')}"
            f"{cell_formula(f'C{i}', f'IF({proc_ref}<>"" , SUMIF(RawData!C:C, {proc_ref}, RawData!E:E), "")')}"
            f"{cell_formula(f'D{i}', f'IFERROR(C{i}/B{i}, 0)')}"
            f"{cell_formula(f'E{i}', f'IF({proc_ref}<>"" , IFERROR(AVERAGEIF(RawData!C:C, {proc_ref}, RawData!F:F), 0), "")')}"
            "</row>"
        )

    sheet_data = "".join(rows)

    return xml_header() + (
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        f"<dimension ref=\"A1:E120\"/>"
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/>"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def generate_xlsx(path: str):
    out_dir = os.path.dirname(path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        # Root parts
        z.writestr('[Content_Types].xml', build_content_types())
        z.writestr('_rels/.rels', build_root_rels())
        z.writestr('docProps/app.xml', build_app_props())
        z.writestr('docProps/core.xml', build_core_props())

        # Workbook and relationships
        z.writestr('xl/workbook.xml', build_workbook())
        z.writestr('xl/_rels/workbook.xml.rels', build_workbook_rels())
        z.writestr('xl/styles.xml', build_styles())

        # Sheets
        z.writestr('xl/worksheets/sheet1.xml', build_sheet1_rawdata())
        z.writestr('xl/worksheets/sheet2.xml', build_sheet2_dashboard())


if __name__ == '__main__':
    generate_xlsx(OUTPUT_PATH)
    print(f"Created: {OUTPUT_PATH}")