import argparse
import csv
import html
import json
import math
import posixpath
import statistics
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "data" / "Startup_Financial_KPI_Analysis_Dataset.xlsx"
DEFAULT_OUTPUT = ROOT / "deliverables"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS}

MONTH_FIELDS = {"month", "signup_month", "cohort_month", "activity_month"}


def column_index(reference):
    letters = "".join(character for character in reference if character.isalpha())
    index = 0
    for character in letters:
        index = index * 26 + ord(character.upper()) - ord("A") + 1
    return index - 1


def read_xlsx(path):
    """Read values from a simple .xlsx workbook without third-party packages."""
    with zipfile.ZipFile(path) as workbook:
        workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
        relationships_xml = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        relationships = {item.attrib["Id"]: item.attrib["Target"] for item in relationships_xml}
        shared_strings = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            strings_xml = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(text.text or "" for text in item.findall(".//m:t", NS))
                for item in strings_xml.findall("m:si", NS)
            ]

        sheets = {}
        for sheet in workbook_xml.findall("m:sheets/m:sheet", NS):
            target = relationships[sheet.attrib["{" + REL_NS + "}id"]]
            sheet_path = (
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.normpath(posixpath.join("xl", target))
            )
            root = ET.fromstring(workbook.read(sheet_path))
            raw_rows = []
            for row in root.findall(".//m:sheetData/m:row", NS):
                values = {}
                for cell in row.findall("m:c", NS):
                    value_node = cell.find("m:v", NS)
                    value = value_node.text if value_node is not None else None
                    cell_type = cell.attrib.get("t")
                    if cell_type == "s" and value is not None:
                        value = shared_strings[int(value)]
                    elif cell_type == "inlineStr":
                        value = "".join(text.text or "" for text in cell.findall(".//m:t", NS))
                    elif cell_type == "b" and value is not None:
                        value = value == "1"
                    elif value is not None and cell_type not in {"str", "e"}:
                        numeric = float(value)
                        value = int(numeric) if numeric.is_integer() else numeric
                    values[column_index(cell.attrib["r"])] = value
                raw_rows.append(values)

            if not raw_rows:
                sheets[sheet.attrib["name"]] = []
                continue
            width = max(max(row.keys(), default=-1) for row in raw_rows) + 1
            headers = [raw_rows[0].get(index) for index in range(width)]
            records = []
            for raw in raw_rows[1:]:
                record = {}
                for index, field in enumerate(headers):
                    if field is None:
                        continue
                    value = raw.get(index)
                    if field in MONTH_FIELDS and isinstance(value, (int, float)):
                        value = (datetime(1899, 12, 30) + timedelta(days=value)).strftime("%Y-%m")
                    record[field] = value
                if any(value is not None for value in record.values()):
                    records.append(record)
            sheets[sheet.attrib["name"]] = records
        return sheets


def require_sheets(sheets, expected):
    missing = [name for name in expected if name not in sheets or not sheets[name]]
    if missing:
        raise ValueError("Workbook is missing required sheets or data: " + ", ".join(missing))


def number(record, field):
    value = record.get(field)
    if value is None or value == "":
        raise ValueError("Missing numeric value for '" + field + "'")
    return float(value)


def rounded(value, places=2):
    return round(value, places) if value is not None else ""


def monthly_kpis(financials):
    required = {
        "month", "revenue_usd", "total_operating_expenses_usd", "net_burn_usd",
        "marketing_spend_usd", "new_customers", "active_customers",
        "churned_customers_estimate", "cash_balance_start_usd",
        "monthly_arpu_usd", "cash_balance_end_usd", "run_rate_arr_usd",
        "gross_margin_pct",
    }
    if not required.issubset(financials[0]):
        raise ValueError("Monthly_Financials does not contain all expected KPI fields")

    output = []
    previous_revenue = None
    previous_customers = None
    for item in sorted(financials, key=lambda row: row["month"]):
        revenue = number(item, "revenue_usd")
        new_customers = number(item, "new_customers")
        net_burn = number(item, "net_burn_usd")
        ending_cash = number(item, "cash_balance_end_usd")
        marketing = number(item, "marketing_spend_usd")
        churned = number(item, "churned_customers_estimate")
        gross_margin = number(item, "gross_margin_pct")
        churn_rate = churned / previous_customers if previous_customers else None
        ltv = (
            number(item, "monthly_arpu_usd") * (gross_margin / 100) / churn_rate
            if churn_rate and churn_rate > 0 and gross_margin > 0 else None
        )
        cac = marketing / new_customers if new_customers else None
        output.append({
            "Month": item["month"],
            "Revenue": rounded(revenue),
            "Operating Expenses": rounded(number(item, "total_operating_expenses_usd")),
            "Net Burn": rounded(net_burn),
            "Marketing Spend": rounded(marketing),
            "New Customers": int(new_customers),
            "Active Customers": int(number(item, "active_customers")),
            "Churned Customers Estimate": int(churned),
            "Starting Customers Estimate": int(previous_customers) if previous_customers else "",
            "Monthly Churn Estimate %": rounded(churn_rate * 100) if churn_rate is not None else "",
            "Starting Cash": rounded(number(item, "cash_balance_start_usd")),
            "ARPU": rounded(number(item, "monthly_arpu_usd")),
            "Ending Cash": rounded(ending_cash),
            "ARR Run Rate": rounded(number(item, "run_rate_arr_usd")),
            "Gross Margin %": rounded(gross_margin),
            "Blended CAC": rounded(cac) if cac is not None else "",
            "Estimated LTV": rounded(ltv) if ltv is not None else "",
            "Estimated LTV:CAC": rounded(ltv / cac, 2) if ltv is not None and cac else "",
            "Revenue Growth %": (
                rounded((revenue / previous_revenue - 1) * 100)
                if previous_revenue else ""
            ),
            "Runway Months": rounded(ending_cash / net_burn) if net_burn > 0 else "",
        })
        previous_revenue = revenue
        previous_customers = number(item, "active_customers")
    return output


def cohort_retention(cohorts):
    required = {
        "cohort_month", "activity_month", "cohort_age_month",
        "starting_cohort_customers", "active_customers",
        "retention_rate_pct", "cohort_revenue_usd",
    }
    if not required.issubset(cohorts[0]):
        raise ValueError("Cohort_Analysis does not contain all expected fields")
    rows = []
    for item in sorted(cohorts, key=lambda row: (row["cohort_month"], number(row, "cohort_age_month"))):
        rows.append({
            "Cohort Month": item["cohort_month"],
            "Activity Month": item["activity_month"],
            "Cohort Age Month": int(number(item, "cohort_age_month")),
            "Starting Cohort Customers": int(number(item, "starting_cohort_customers")),
            "Active Customers": int(number(item, "active_customers")),
            "Retention Rate %": rounded(number(item, "retention_rate_pct")),
            "Cohort Revenue": rounded(number(item, "cohort_revenue_usd")),
        })
    return rows


def channel_economics(customers, activity, last_month):
    channels = {}
    customer_channel = {}
    for item in customers:
        channel = item["acquisition_channel"]
        details = channels.setdefault(channel, {
            "costs": [], "customers": set(), "revenue": 0.0, "active_latest": set(),
        })
        customer_id = item["customer_id"]
        details["customers"].add(customer_id)
        customer_channel[customer_id] = channel
        details["costs"].append(number(item, "estimated_acquisition_cost_usd"))

    for item in activity:
        channel = customer_channel.get(item["customer_id"], item.get("acquisition_channel"))
        if channel not in channels:
            continue
        details = channels[channel]
        details["revenue"] += number(item, "recognized_revenue_usd")
        if item["month"] == last_month and number(item, "active_customer") > 0:
            details["active_latest"].add(item["customer_id"])

    rows = []
    for channel, details in sorted(channels.items()):
        count = len(details["customers"])
        costs = details["costs"]
        rows.append({
            "Acquisition Channel": channel,
            "Customers": count,
            "Total Estimated CAC": rounded(sum(costs)),
            "Average Estimated CAC": rounded(statistics.mean(costs)),
            "Median Estimated CAC": rounded(statistics.median(costs)),
            "Cumulative Recognized Revenue": rounded(details["revenue"]),
            "Latest Month Active Customers": len(details["active_latest"]),
            "Latest Month Active Rate %": rounded(len(details["active_latest"]) / count * 100)
            if count else "",
        })
    return rows


def expense_detail(expenses):
    required = {"month", "expense_category", "expense_amount_usd"}
    if not required.issubset(expenses[0]):
        raise ValueError("Expense_Details does not contain all expected fields")
    return [
        {
            "Month": item["month"],
            "Expense Category": item["expense_category"],
            "Expense Amount": rounded(number(item, "expense_amount_usd")),
        }
        for item in sorted(expenses, key=lambda row: (row["month"], row["expense_category"]))
    ]


def write_csv(path, rows):
    if not rows:
        raise ValueError("Cannot write an empty CSV: " + str(path))
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def excel_column(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def excel_cell(reference, value, style=0):
    style_attr = ' s="%d"' % style if style else ""
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return '<c r="%s" t="b"%s><v>%d</v></c>' % (reference, style_attr, int(value))
    if isinstance(value, (int, float)) and math.isfinite(value):
        return '<c r="%s"%s><v>%s</v></c>' % (reference, style_attr, value)
    return (
        '<c r="%s" t="inlineStr"%s><is><t xml:space="preserve">%s</t></is></c>'
        % (reference, style_attr, html.escape(str(value), quote=False))
    )


def worksheet_xml(headers, rows):
    last_col = excel_column(len(headers))
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="%s"><sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><sheetData>' % MAIN_NS,
        '<row r="1">' + "".join(
            excel_cell(excel_column(index) + "1", header, 1)
            for index, header in enumerate(headers, 1)
        ) + "</row>",
    ]
    for row_index, record in enumerate(rows, 2):
        cells = []
        for col_index, header in enumerate(headers, 1):
            cells.append(excel_cell(
                excel_column(col_index) + str(row_index), record.get(header)
            ))
        parts.append('<row r="%d">%s</row>' % (row_index, "".join(cells)))
    parts.append(
        "</sheetData><autoFilter ref=\"A1:%s%d\"/></worksheet>"
        % (last_col, len(rows) + 1)
    )
    return "".join(parts)


def write_analysis_xlsx(path, tables):
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    workbook_sheets = []
    relationships = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    sheet_xml = []
    for index, (name, headers, rows) in enumerate(tables, 1):
        content_types.append(
            '<Override PartName="/xl/worksheets/sheet%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            % index
        )
        workbook_sheets.append(
            '<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
            % (html.escape(name, quote=True), index, index)
        )
        relationships.append(
            '<Relationship Id="rId%d" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet%d.xml"/>'
            % (index, index)
        )
        sheet_xml.append(worksheet_xml(headers, rows))
    relationships.append(
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/></Relationships>' % (len(tables) + 1)
    )
    content_types.append("</Types>")
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets>'
        '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
        % (MAIN_NS, REL_NS, "".join(workbook_sheets))
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="%s"><fonts count="2"><font><sz val="11"/>'
        '<name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/>'
        '<sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF164E63"/>'
        '<bgColor indexed="64"/></patternFill></fill></fills><borders count="1">'
        '<border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        '</cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" '
        'borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" '
        'xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1">'
        '<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        % MAIN_NS
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", "".join(content_types))
        package.writestr("_rels/.rels", root_rels)
        package.writestr("xl/workbook.xml", workbook_xml)
        package.writestr("xl/_rels/workbook.xml.rels", "".join(relationships))
        package.writestr("xl/styles.xml", styles_xml)
        for index, content in enumerate(sheet_xml, 1):
            package.writestr("xl/worksheets/sheet%d.xml" % index, content)


def money(value):
    return ("-$" if value < 0 else "$") + format(abs(value), ",.0f")


def pdf_escape(text):
    return str(text).encode("ascii", "replace").decode("ascii").replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def make_pdf(path, monthly, cohorts, channels):
    latest = monthly[-1]
    first = monthly[0]
    growth = (latest["Revenue"] / first["Revenue"] - 1) * 100 if first["Revenue"] else 0
    breakeven = next((row["Month"] for row in monthly if row["Net Burn"] <= 0), "Not reached")
    month12 = [
        row["Retention Rate %"] for row in cohorts
        if row["Cohort Age Month"] == 12
    ]
    median_month12 = statistics.median(month12) if month12 else None
    lines = [
        ("Period", "%s to %s" % (first["Month"], latest["Month"])),
        ("Revenue", "%s to %s (%+.1f%%)" % (money(first["Revenue"]), money(latest["Revenue"]), growth)),
        ("ARR run rate", money(latest["ARR Run Rate"])),
        ("Active customers", "%s to %s" % (first["Active Customers"], latest["Active Customers"])),
        ("Latest blended CAC", money(latest["Blended CAC"])),
        ("Estimated LTV / LTV:CAC", (
            "%s / %.2fx" % (money(latest["Estimated LTV"]), latest["Estimated LTV:CAC"])
            if latest["Estimated LTV:CAC"] != "" else "Unavailable"
        )),
        ("Latest gross margin", "%.2f%% (source field)" % latest["Gross Margin %"]),
        ("Ending cash", money(latest["Ending Cash"])),
        ("Latest monthly net burn", money(latest["Net Burn"])),
        ("First cash-generative month", breakeven),
        ("Median month-12 cohort retention", "%.1f%%" % median_month12 if median_month12 is not None else "Unavailable"),
    ]
    cmds = ["BT /F1 20 Tf 48 748 Td (Startup Financial KPI Analysis) Tj ET"]
    cmds.append("BT /F1 10 Tf 48 725 Td (Dataset: synthetic/demo B2B SaaS data; USD; educational analysis only.) Tj ET")
    y = 693
    for label, value in lines:
        cmds.append("BT /F1 11 Tf 52 %d Td (%s) Tj ET" % (y, pdf_escape(label + ": " + value)))
        y -= 21
    cmds.append("BT /F1 13 Tf 48 455 Td (Monthly revenue trend) Tj ET")

    baseline = 245
    max_height = 170
    maximum = max(row["Revenue"] for row in monthly) or 1
    step = 510 / len(monthly)
    for index, row in enumerate(monthly):
        x = 56 + index * step
        height = row["Revenue"] / maximum * max_height
        cmds.extend([
            "0.09 0.31 0.39 rg",
            "%.2f %d %.2f %.2f re f" % (x, baseline, max(8, step - 5), height),
            "0 0 0 rg",
            "BT /F1 6 Tf %.2f %d Td (%s) Tj ET" % (x - 1, baseline - 12, row["Month"][2:]),
        ])
    cmds.extend([
        "BT /F1 12 Tf 48 192 Td (Acquisition channel snapshot) Tj ET",
        "BT /F1 9 Tf 52 174 Td (Channel | customers | average estimated CAC | cumulative recognized revenue) Tj ET",
    ])
    y = 158
    for channel in channels[:5]:
        line = "%s | %s | %s | %s" % (
            channel["Acquisition Channel"], channel["Customers"],
            money(channel["Average Estimated CAC"]),
            money(channel["Cumulative Recognized Revenue"]),
        )
        cmds.append("BT /F1 9 Tf 52 %d Td (%s) Tj ET" % (y, pdf_escape(line)))
        y -= 14
    cmds.extend([
        "BT /F1 8 Tf 48 55 Td (Definitions and limitations: blended CAC = monthly marketing spend / new customers; channel CAC is customer-level estimated cost.) Tj ET",
        "BT /F1 8 Tf 48 42 Td (Net burn is reported as operating expenses minus revenue; negative values indicate cash generation in this dataset.) Tj ET",
        "BT /F1 8 Tf 48 29 Td (Gross margin and retention are reported from source sheets; synthetic data is not actual company financial information.) Tj ET",
    ])
    stream = "\n".join(cmds).encode("latin-1", "replace")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length %d >>\nstream\n%s\nendstream" % (len(stream), stream),
    ]
    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, obj in enumerate(objects, 1):
        offsets.append(len(pdf))
        pdf.extend(("%d 0 obj\n" % index).encode("ascii") + obj + b"\nendobj\n")
    xref_offset = len(pdf)
    pdf.extend(("xref\n0 %d\n" % (len(objects) + 1)).encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(("%010d 00000 n \n" % offset).encode("ascii"))
    pdf.extend((
        "trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF"
        % (len(objects) + 1, xref_offset)
    ).encode("ascii"))
    path.write_bytes(pdf)


def json_for_script(value):
    return json.dumps(value, separators=(",", ":"), ensure_ascii=True).replace("<", "\\u003c")


def make_dashboard(path, monthly, cohorts, channels, expenses):
    cohort_matrix = defaultdict(dict)
    cohort_months = set()
    max_age = 0
    for row in cohorts:
        cohort_months.add(row["Cohort Month"])
        max_age = max(max_age, row["Cohort Age Month"])
        cohort_matrix[row["Cohort Month"]][row["Cohort Age Month"]] = row["Retention Rate %"]
    cohort_data = [
        {"month": month, "cells": cohort_matrix[month]}
        for month in sorted(cohort_months)
    ]
    last_month_expenses = [
        row for row in expenses if row["Month"] == monthly[-1]["Month"]
    ]
    data = {
        "monthly": monthly,
        "cohorts": cohort_data,
        "channels": channels,
        "expenses": last_month_expenses,
        "maxAge": max_age,
    }
    payload = json_for_script(data)
    page = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Startup Financial KPI Dashboard</title>
<style>
:root{font-family:Inter,Segoe UI,Arial,sans-serif;color:#102a43;background:#f3f7fa}body{margin:0}
header{background:#123b4a;color:#fff;padding:24px max(5vw,24px)}header h1{margin:0 0 7px;font-size:25px}
header p{margin:0;color:#d4e5e9;font-size:14px}.wrap{max-width:1220px;margin:24px auto;padding:0 20px}
.notice{border-left:4px solid #d38b00;background:#fff8e6;padding:12px 16px;margin-bottom:18px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}.card,.panel{background:white;border:1px solid #dce6eb;border-radius:10px;padding:16px;box-shadow:0 2px 6px #17324d0b}
.label{font-size:12px;color:#526777}.value{font-size:23px;font-weight:700;margin:7px 0}.trend{font-size:12px;color:#526777}
.grid{display:grid;grid-template-columns:1.35fr 1fr;gap:14px;margin-top:14px}.panel h2{font-size:16px;margin:0 0 14px}
select{padding:7px;border:1px solid #c9d6dd;border-radius:6px;background:#fff;margin-left:8px}
svg{width:100%;height:auto;overflow:visible}.chart-label{font-size:10px;fill:#526777}.gridline{stroke:#e5edf1;stroke-width:1}
.table-wrap{overflow:auto;max-height:520px}.data{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap}.data th,.data td{padding:7px;text-align:right;border:1px solid #e6edf1}.data th{background:#f4f8fa;position:sticky;top:0}.data th:first-child,.data td:first-child{text-align:left}
.heat td{min-width:42px;text-align:center}.foot{font-size:12px;color:#526777;margin:16px 0 32px}.links a{color:#087e8b;margin-right:18px}
@media(max-width:750px){.grid{grid-template-columns:1fr}}
</style></head>
<body><header><h1>Startup Financial KPI Dashboard</h1><p id="subtitle"></p></header>
<main class="wrap"><div class="notice"><strong>Synthetic/demo data.</strong> The source workbook labels these as learning data, not actual company financials. Metrics reflect the supplied assumptions and are for educational analysis only. Estimated LTV uses monthly ARPU × reported gross margin ÷ estimated logo churn; LTV:CAC is indicative, not a forecast.</div>
<section class="cards" id="cards"></section>
<section class="grid"><div class="panel"><h2>Monthly trend <select id="metric">
<option value="Revenue">Revenue</option><option value="Net Burn">Net burn</option><option value="Active Customers">Active customers</option>
<option value="Blended CAC">Blended CAC</option><option value="Estimated LTV:CAC">Estimated LTV:CAC</option><option value="ARR Run Rate">ARR run rate</option>
<option value="Gross Margin %">Gross margin</option><option value="Ending Cash">Ending cash</option>
</select></h2><svg id="trend" viewBox="0 0 760 310" role="img" aria-label="Monthly KPI trend"></svg></div>
<div class="panel"><h2>Latest-month expenses</h2><svg id="expense-chart" viewBox="0 0 540 310" role="img" aria-label="Latest month expense categories"></svg></div></section>
<section class="panel" style="margin-top:14px"><h2>Cohort retention</h2><div class="table-wrap"><table class="data heat" id="cohort"></table></div></section>
<section class="grid"><div class="panel"><h2>Acquisition channel economics</h2><div class="table-wrap"><table class="data" id="channels"></table></div></div>
<div class="panel"><h2>Monthly KPI detail</h2><div class="table-wrap"><table class="data" id="detail"></table></div></div></section>
<p class="foot links">CSV exports: <a href="Tableau_Monthly_KPIs.csv">monthly KPIs</a><a href="Tableau_Cohort_Retention.csv">cohort retention</a><a href="Tableau_Channel_Economics.csv">channel economics</a><a href="Tableau_Expense_Detail.csv">expense detail</a><br>Workbook: <a href="Startup_KPI_Analysis.xlsx">Excel analysis</a> · Report: <a href="LTV_CAC_Analysis_Report.pdf">PDF summary</a>.</p>
</main><script>
const data=__DATA__,monthly=data.monthly,latest=monthly[monthly.length-1],first=monthly[0];
const esc=v=>String(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money=n=>(n<0?'-$':'$')+Math.round(Math.abs(n)).toLocaleString('en-US');
const pct=n=>Number(n).toFixed(1)+'%';
document.getElementById('subtitle').textContent=`${first.Month} – ${latest.Month} · Source: supplied startup KPI workbook`;
const cashText=latest['Net Burn']<0?money(-latest['Net Burn'])+' generation':money(latest['Net Burn'])+' burn';
const cards=[['Revenue',money(latest.Revenue),latest.Month],['ARR run rate',money(latest['ARR Run Rate']),'Latest monthly revenue × 12'],
['Active customers',latest['Active Customers'].toLocaleString(),'Latest source count'],['Gross margin',pct(latest['Gross Margin %']),'As reported in source'],
['Blended CAC',money(latest['Blended CAC']),'Monthly marketing / new customers'],['Ending cash',money(latest['Ending Cash']),'Reported month-end balance'],
['Net cash flow',cashText,'Negative net burn indicates cash generation'],
['Estimated LTV:CAC',latest['Estimated LTV:CAC']===''?'n/a':Number(latest['Estimated LTV:CAC']).toFixed(2)+'x','ARPU × gross margin / estimated churn']];
document.getElementById('cards').innerHTML=cards.map(x=>`<div class="card"><div class="label">${esc(x[0])}</div><div class="value">${esc(x[1])}</div><div class="trend">${esc(x[2])}</div></div>`).join('');
const currencyKeys=['Revenue','Net Burn','Blended CAC','Estimated LTV','ARR Run Rate','Ending Cash'];
const axisFormat=(key,value)=>currencyKeys.includes(key)?money(value):key==='Gross Margin %'?pct(value):key==='Estimated LTV:CAC'?Number(value).toFixed(1)+'x':Math.round(value).toLocaleString('en-US');
function drawTrend(){const key=document.getElementById('metric').value,svg=document.getElementById('trend');
const points=monthly.map((d,i)=>({x:i,v:d[key]===''||d[key]===null?NaN:Number(d[key])})).filter(p=>Number.isFinite(p.v));
if(!points.length){svg.innerHTML='<text x="20" y="40" class="chart-label">Metric unavailable for this period.</text>';return}
let lo=Math.min(0,...points.map(p=>p.v)),hi=Math.max(0,...points.map(p=>p.v));if(lo===hi){lo-=1;hi+=1}const left=62,right=742,top=18,bottom=258;
const x=i=>left+i*(right-left)/(monthly.length-1||1),y=v=>bottom-(v-lo)/(hi-lo)*(bottom-top);let s='';
for(let j=0;j<5;j++){const v=lo+(hi-lo)*j/4,yy=y(v);s+=`<line class="gridline" x1="${left}" y1="${yy}" x2="${right}" y2="${yy}"/><text class="chart-label" x="2" y="${yy+3}">${esc(axisFormat(key,v))}</text>`}
s+=`<polyline fill="none" stroke="#07808c" stroke-width="3" points="${points.map(p=>`${x(p.x)},${y(p.v)}`).join(' ')}"/>`;
points.forEach(p=>{s+=`<circle cx="${x(p.x)}" cy="${y(p.v)}" r="4" fill="#07808c"><title>${esc(monthly[p.x].Month)}: ${esc(axisFormat(key,p.v))}</title></circle>`;if(p.x%2===0||p.x===monthly.length-1)s+=`<text class="chart-label" text-anchor="middle" x="${x(p.x)}" y="280">${esc(monthly[p.x].Month.slice(2))}</text>`});
svg.innerHTML=s}
document.getElementById('metric').addEventListener('change',drawTrend);drawTrend();
const expenseSvg=document.getElementById('expense-chart'),exp=data.expenses,expMax=Math.max(1,...exp.map(x=>x['Expense Amount']));
expenseSvg.innerHTML=exp.map((item,i)=>{const y=24+i*34,w=item['Expense Amount']/expMax*280;return `<text class="chart-label" x="2" y="${y+12}">${esc(item['Expense Category'])}</text><rect x="235" y="${y}" width="${w}" height="18" rx="3" fill="#07808c"/><text class="chart-label" x="${Math.min(520,242+w)}" y="${y+12}">${esc(money(item['Expense Amount']))}</text>`}).join('');
const ageCount=data.maxAge+1,cohort=document.getElementById('cohort');
cohort.innerHTML='<thead><tr><th>Cohort</th>'+Array.from({length:ageCount},(_,i)=>`<th>Month ${i}</th>`).join('')+'</tr></thead><tbody>'+data.cohorts.map(c=>'<tr><th>'+esc(c.month)+'</th>'+Array.from({length:ageCount},(_,i)=>{const v=c.cells[i];return v===undefined?'<td>—</td>':`<td style="background:rgba(7,128,140,${.10+.55*Math.max(0,Math.min(100,v))/100});color:${v>65?'white':'#102a43'}">${esc(pct(v))}</td>`}).join('')+'</tr>').join('')+'</tbody>';
const channelCols=[['Acquisition Channel','Channel'],['Customers','Customers'],['Average Estimated CAC','Avg est. CAC'],['Total Estimated CAC','Total est. CAC'],['Cumulative Recognized Revenue','Recognized revenue'],['Latest Month Active Rate %','Latest active %']];
document.getElementById('channels').innerHTML='<thead><tr>'+channelCols.map(x=>`<th>${esc(x[1])}</th>`).join('')+'</tr></thead><tbody>'+data.channels.map(r=>'<tr>'+channelCols.map(([key])=>`<td>${esc(key==='Acquisition Channel'?r[key]:key==='Customers'?Number(r[key]).toLocaleString('en-US'):key==='Latest Month Active Rate %'?pct(r[key]):money(r[key]))}</td>`).join('')+'</tr>').join('')+'</tbody>';
const cols=['Month','Revenue','Net Burn','Blended CAC','Estimated LTV:CAC','Gross Margin %','Active Customers','ARR Run Rate','Ending Cash'];
document.getElementById('detail').innerHTML='<thead><tr>'+cols.map(c=>`<th>${esc(c)}</th>`).join('')+'</tr></thead><tbody>'+monthly.map(r=>'<tr>'+cols.map(c=>`<td>${esc(c==='Month'?r[c]:c==='Gross Margin %'?pct(r[c]):c==='Active Customers'?Number(r[c]).toLocaleString('en-US'):c==='Estimated LTV:CAC'?(r[c]===''?'n/a':Number(r[c]).toFixed(2)+'x'):money(r[c]))}</td>`).join('')+'</tr>').join('')+'</tbody>';
</script></body></html>'''
    path.write_text(page.replace("__DATA__", payload), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Build startup KPI outputs from the supplied Excel dataset.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Path to the source .xlsx workbook.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="Directory for generated deliverables.")
    args = parser.parse_args()
    if not args.input.is_file():
        raise FileNotFoundError("Source workbook not found: " + str(args.input))

    sheets = read_xlsx(args.input)
    require_sheets(sheets, [
        "Monthly_Financials", "Customer_Data", "Customer_Monthly_Activity",
        "Expense_Details", "Cohort_Analysis",
    ])
    monthly = monthly_kpis(sheets["Monthly_Financials"])
    cohorts = cohort_retention(sheets["Cohort_Analysis"])
    expenses = expense_detail(sheets["Expense_Details"])
    channels = channel_economics(
        sheets["Customer_Data"], sheets["Customer_Monthly_Activity"], monthly[-1]["Month"]
    )

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(args.output_dir / "Tableau_Monthly_KPIs.csv", monthly)
    write_csv(args.output_dir / "Tableau_Cohort_Retention.csv", cohorts)
    write_csv(args.output_dir / "Tableau_Channel_Economics.csv", channels)
    write_csv(args.output_dir / "Tableau_Expense_Detail.csv", expenses)
    write_analysis_xlsx(args.output_dir / "Startup_KPI_Analysis.xlsx", [
        ("Monthly KPIs", list(monthly[0]), monthly),
        ("Cohort Retention", list(cohorts[0]), cohorts),
        ("Channel Economics", list(channels[0]), channels),
        ("Expense Detail", list(expenses[0]), expenses),
    ])
    make_pdf(args.output_dir / "LTV_CAC_Analysis_Report.pdf", monthly, cohorts, channels)
    make_dashboard(args.output_dir / "Startup_KPI_Dashboard.html", monthly, cohorts, channels, expenses)
    print(
        "Generated dataset-based KPI deliverables from %s (%d months, %d cohorts, %d channels)."
        % (args.input, len(monthly), len({row["Cohort Month"] for row in cohorts}), len(channels))
    )
    print("Output directory: " + str(args.output_dir))


if __name__ == "__main__":
    main()
