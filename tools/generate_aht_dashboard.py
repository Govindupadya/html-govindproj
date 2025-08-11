import os
from datetime import date
import xlsxwriter

OUTPUT_PATH = "/workspace/AHT_Productivity_Dashboard.xlsx"


def write_rawdata_sheet(workbook):
    worksheet = workbook.add_worksheet("RawData")

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9E1F2",
        "border": 1
    })
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})

    headers = [
        "Date",
        "Agent",
        "Process",
        "Allocated",
        "Processed",
        "AHT (mins)",
    ]

    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_format)

    # Sample rows (you can delete these later)
    sample_rows = [
        [date.today().isoformat(), "Agent 1", "Process A", 100, 90, 4.5],
        [date.today().isoformat(), "Agent 1", "Process B", 80, 65, 5.2],
        [date.today().isoformat(), "Agent 2", "Process A", 120, 110, 4.2],
        [date.today().isoformat(), "Agent 2", "Process C", 75, 60, 6.1],
    ]

    for row_idx, row in enumerate(sample_rows, start=1):
        # Date
        worksheet.write_datetime(row_idx, 0, date.fromisoformat(str(row[0])), date_format)
        # Agent, Process
        worksheet.write(row_idx, 1, row[1])
        worksheet.write(row_idx, 2, row[2])
        # Allocated, Processed, AHT (mins)
        worksheet.write_number(row_idx, 3, row[3])
        worksheet.write_number(row_idx, 4, row[4])
        worksheet.write_number(row_idx, 5, row[5])

    # Create an Excel Table for a large input range so you can keep adding rows
    last_row = 1000
    last_col = len(headers) - 1
    worksheet.add_table(0, 0, last_row, last_col, {
        "name": "RawData",
        "style": {"theme": "Table Style Light 11", "show_first_column": False, "show_last_column": False},
        "columns": [{"header": h} for h in headers],
        "banded_rows": True,
        "autofilter": True,
    })

    # Column widths
    worksheet.set_column("A:A", 12)
    worksheet.set_column("B:B", 16)
    worksheet.set_column("C:C", 18)
    worksheet.set_column("D:E", 12)
    worksheet.set_column("F:F", 12)


def write_dashboard_sheet(workbook):
    ws = workbook.add_worksheet("Dashboard")

    title_format = workbook.add_format({"bold": True, "font_size": 16})
    header_format = workbook.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1})
    int_format = workbook.add_format({"num_format": "#,##0"})
    pct_format = workbook.add_format({"num_format": "0.00%"})
    time_format = workbook.add_format({"num_format": "0.0"})

    ws.write("A1", "AHT Productivity Dashboard", title_format)

    # Overall KPI tiles
    ws.write("A3", "Total Allocated", header_format)
    ws.write("B3", "=SUM(RawData[Allocated])", int_format)

    ws.write("A4", "Total Processed", header_format)
    ws.write("B4", "=SUM(RawData[Processed])", int_format)

    ws.write("D3", "Overall Productivity%", header_format)
    ws.write("E3", "=IFERROR(B4/B3,0)", pct_format)

    ws.write("D4", "Overall Avg AHT (mins)", header_format)
    ws.write("E4", "=IFERROR(AVERAGE(RawData[AHT (mins)]),0)", time_format)

    # Table headers for per-Process summary
    start_row = 6  # 1-indexed row 7
    ws.write(start_row - 1, 0, "Process", header_format)
    ws.write(start_row - 1, 1, "Allocated", header_format)
    ws.write(start_row - 1, 2, "Processed", header_format)
    ws.write(start_row - 1, 3, "Productivity%", header_format)
    ws.write(start_row - 1, 4, "Average AHT (mins)", header_format)

    # Build formulas down for up to 100 processes
    max_rows = 100
    for i in range(max_rows):
        row = start_row + i
        proc_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 0)
        alloc_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
        procd_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 2)
        prodp_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 3)
        aht_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 4)

        if i == 0:
            # Unique process list spills from first cell (Excel 365+)
            ws.write_formula(proc_cell, "=SORT(UNIQUE(RawData[Process]))")
        else:
            # Keep format consistency
            ws.write_blank(proc_cell, None)

        ws.write_formula(alloc_cell, f"=IF({proc_cell}<>\"\", SUMIF(RawData[Process], {proc_cell}, RawData[Allocated]), \"\")")
        ws.write_formula(procd_cell, f"=IF({proc_cell}<>\"\", SUMIF(RawData[Process], {proc_cell}, RawData[Processed]), \"\")")
        ws.write_formula(prodp_cell, f"=IFERROR({procd_cell}/{alloc_cell}, 0)")
        ws.write_formula(aht_cell, f"=IF({proc_cell}<>\"\", IFERROR(AVERAGEIF(RawData[Process], {proc_cell}, RawData[AHT (mins)]), 0), \"\")")

        ws.write_blank(row, 1, None, int_format)
        ws.write_blank(row, 2, None, int_format)
        ws.write_blank(row, 3, None, pct_format)
        ws.write_blank(row, 4, None, time_format)

    # Named ranges limited to actual non-empty processes using COUNTA on Process column
    end_row = start_row + max_rows - 1
    workbook.define_name("ProcessNames", f"=Dashboard!$A${start_row}:INDEX(Dashboard!$A${start_row}:$A${end_row}, COUNTA(Dashboard!$A${start_row}:$A${end_row}))")
    workbook.define_name("AllocatedByProcess", f"=Dashboard!$B${start_row}:INDEX(Dashboard!$B${start_row}:$B${end_row}, COUNTA(Dashboard!$A${start_row}:$A${end_row}))")
    workbook.define_name("ProcessedByProcess", f"=Dashboard!$C${start_row}:INDEX(Dashboard!$C${start_row}:$C${end_row}, COUNTA(Dashboard!$A${start_row}:$A${end_row}))")
    workbook.define_name("ProductivityByProcess", f"=Dashboard!$D${start_row}:INDEX(Dashboard!$D${start_row}:$D${end_row}, COUNTA(Dashboard!$A${start_row}:$A${end_row}))")
    workbook.define_name("AvgAHTByProcess", f"=Dashboard!$E${start_row}:INDEX(Dashboard!$E${start_row}:$E${end_row}, COUNTA(Dashboard!$A${start_row}:$A${end_row}))")

    # Conditional formatting for Productivity%
    ws.conditional_format(start_row, 3, end_row, 3, {
        "type": "3_color_scale",
        "min_color": "#F8696B",
        "mid_color": "#FFEB84",
        "max_color": "#63BE7B",
    })

    # Charts
    chart1 = workbook.add_chart({"type": "column"})
    chart1.add_series({
        "name": "Productivity%",
        "categories": "=ProcessNames",
        "values": "=ProductivityByProcess",
        "y2_axis": False,
    })
    chart1.set_title({"name": "Productivity% by Process"})
    chart1.set_y_axis({"name": "Productivity%", "num_format": "0%"})
    chart1.set_legend({"position": "bottom"})
    ws.insert_chart("G6", chart1, {"x_scale": 1.25, "y_scale": 1.25})

    chart2 = workbook.add_chart({"type": "column"})
    chart2.add_series({
        "name": "Average AHT (mins)",
        "categories": "=ProcessNames",
        "values": "=AvgAHTByProcess",
    })
    chart2.set_title({"name": "Average AHT (mins) by Process"})
    chart2.set_y_axis({"name": "Minutes", "num_format": "0.0"})
    chart2.set_legend({"position": "bottom"})
    ws.insert_chart("G22", chart2, {"x_scale": 1.25, "y_scale": 1.25})

    # Column widths for dashboard
    ws.set_column("A:A", 20)
    ws.set_column("B:E", 16)


def write_instructions_sheet(workbook):
    ws = workbook.add_worksheet("Instructions")
    text = (
        "How to use this dashboard:\n\n"
        "1) Go to the 'RawData' sheet and paste or enter your tracker rows.\n"
        "   Required columns: Date, Agent, Process, Allocated, Processed, AHT (mins).\n"
        "   - Date: yyyy-mm-dd\n"
        "   - Allocated/Processed: whole numbers\n"
        "   - AHT (mins): average handle time in minutes per completed item\n\n"
        "2) Return to the 'Dashboard' sheet to see per-process totals, Productivity% (Processed/Allocated), and Average AHT.\n\n"
        "3) The charts update automatically based on the data.\n\n"
        "Notes:\n"
        "- You can add up to 1000 data rows and up to 100 distinct processes by default.\n"
        "- You can delete the sample rows in 'RawData'.\n"
    )
    ws.write("A1", text)
    ws.set_column("A:A", 100)
    ws.set_row(0, 200)


def main():
    # Ensure directory exists
    out_dir = os.path.dirname(OUTPUT_PATH)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    workbook = xlsxwriter.Workbook(OUTPUT_PATH)
    try:
        write_rawdata_sheet(workbook)
        write_dashboard_sheet(workbook)
        write_instructions_sheet(workbook)
    finally:
        workbook.close()


if __name__ == "__main__":
    main()