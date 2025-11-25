# report_generator.py
import io
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Spacer, Paragraph, PageBreak, KeepTogether
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def generate_report(employee_file, clock_name_rows):
    """
    employee_file: file-like Excel (main master file)
    clock_name_rows: list of dicts like:
        [{"clock": "123456", "name": "Test"}, ...]
        'name' can be None/"" if not available
    returns: bytes of PDF
    """

    # Load Excel
    df = pd.read_excel(employee_file)

    # Standardize Employee Code
    if "Employee Code" not in df.columns:
        raise ValueError("Column 'Employee Code' not found in Excel")

    df["Employee Code"] = df["Employee Code"].astype(str).str.upper().str.zfill(6)

    # Build:
    # - full clocks list
    # - manual_missing map (clock -> name) from Gemini
    clocks = []
    manual_missing = {}

    for row in clock_name_rows:
        c = str(row.get("clock", "")).upper().zfill(6)
        name = (row.get("name") or "").strip()
        if not c:
            continue
        clocks.append(c)
        if name:
            manual_missing[c] = name

    matched = []
    not_found = []

    for c in clocks:
        row = df[df["Employee Code"] == c]
        if row.empty:
            name = manual_missing.get(c, "Not Found")
            not_found.append([c, name])
        else:
            matched.append([
                c,
                row.iloc[0]["Employee Name"],
                row.iloc[0]["OO Name"],
            ])

    matched_df = pd.DataFrame(matched, columns=["Clock No", "Employee Name", "OO Name"])
    if not matched_df.empty:
        matched_df = matched_df.sort_values(by=["OO Name", "Employee Name"])

    # SUMMARY
    summary_dict = {}
    if not matched_df.empty:
        summary = matched_df.groupby("OO Name")["Clock No"].count()
        summary_dict.update(summary.to_dict())

    summary_dict["Manual Found"] = sum(1 for c, n in not_found if n != "Not Found")
    summary_dict["Not Found"] = sum(1 for c, n in not_found if n == "Not Found")
    summary_dict["Grand Total"] = len(clocks)

    # Build PDF in memory (BytesIO)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # GROUPED OUTPUT – keep each OO block together
    if not matched_df.empty:
        for oo, group in matched_df.groupby("OO Name"):
            block = []
            block.append(Paragraph(f"<b>{oo}</b>", styles["Heading4"]))
            data = [["Clock No", "Employee Name"]] + group[
                ["Clock No", "Employee Name"]
            ].values.tolist()
            t = Table(data, colWidths=[120, 200])
            t.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                    ]
                )
            )
            block.append(t)
            block.append(Spacer(1, 12))
            elements.append(KeepTogether(block))

    # NOT FOUND SECTION
    if not_found:
        elements.append(PageBreak())
        elements.append(Paragraph("<b>Not Found (With Manual Names)</b>", styles["Heading4"]))
        nf = [["Clock No", "Name"]] + not_found
        t_nf = Table(nf, colWidths=[120, 200])
        t_nf.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.black)]))
        elements.append(t_nf)
        elements.append(Spacer(1, 12))

    # SUMMARY
    elements.append(PageBreak())
    elements.append(Paragraph("<b>SUMMARY (OO WISE TOTAL)</b>", styles["Heading4"]))
    summary_table = [["Description", "Count"]] + [[k, v] for k, v in summary_dict.items()]
    t_sum = Table(summary_table, colWidths=[200, 120])
    t_sum.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.30, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ]
        )
    )
    elements.append(t_sum)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()
