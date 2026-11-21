"""Build the stock count reconciliation workbook.

Inputs (project root):
  input/entered_counts.csv   counts entered from the count-sheet photos
  input/system_export.xlsx   stock-on-hand export from the stock system
Output:
  output/Stock_Count_Reconciliation.xlsx

Every comparison in the workbook is a live Excel formula, so the client can
see how each number was worked out and corrections flow through automatically.

    python3 scripts/build_reconciliation.py
"""

import csv
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output" / "Stock_Count_Reconciliation.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUS_COLOURS = {
    "Under": ("FFC7CE", "9C0006"),  # counted less than system
    "Over": ("DDEBF7", "1F4E79"),  # counted more than system
    "Not counted": ("E4DFEC", "5B3A8C"),  # in system, missing from count sheet
    "Not in system": ("FCE4D6", "843C0C"),  # counted, but not in system export
}


def load_inputs():
    with open(ROOT / "input" / "entered_counts.csv", newline="") as fh:
        counts = list(csv.DictReader(fh))
    ws = load_workbook(ROOT / "input" / "system_export.xlsx", read_only=True).active
    system = [
        dict(zip(("store", "store_name", "sku", "description", "qty"), r))
        for r in ws.iter_rows(min_row=2, values_only=True)
        if r[0]
    ]
    return counts, system


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BOX
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws.row_dimensions[row].height = 32


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build():
    counts, system = load_inputs()
    store_names = {s["store"]: s["store_name"] for s in system}
    stores = sorted({r["store"] for r in counts} | set(store_names))
    checks = [r for r in counts if r["reading"] == "Check"]

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    recon = wb.create_sheet("Reconciliation")
    cnt = wb.create_sheet("Counts (Entered)")
    nc = wb.create_sheet("Needs Checking")
    sysws = wb.create_sheet("System Export")

    # ---------- System Export (copied as-is) ----------
    sysws.append(["Store", "Store Name", "SKU", "Description", "On Hand Qty"])
    style_header(sysws, 1, 5)
    for s in system:
        sysws.append(
            [s["store"], s["store_name"], s["sku"], s["description"], s["qty"]]
        )
    set_widths(sysws, [8, 14, 10, 30, 12])
    sysws.freeze_panes = "A2"
    sysws.auto_filter.ref = f"A1:E{sysws.max_row}"

    # ---------- Needs Checking (client confirms unclear readings here) ----------
    nc["A1"] = (
        "Readings that were unclear on the photo. Look at the photo, type the correct number in the yellow column, and mark it resolved. The whole workbook updates automatically."
    )
    nc["A1"].font = Font(italic=True)
    nc.merge_cells("A1:J1")
    nc_headers = [
        "Store",
        "SKU",
        "Description",
        "Entered Qty",
        "Photo",
        "Sheet Line",
        "What Was Seen",
        "Note",
        "Confirmed Qty",
        "Resolved?",
    ]
    nc.append(nc_headers)
    style_header(nc, 2, len(nc_headers))
    for r in checks:
        qty = int(r["counted_qty"]) if r["counted_qty"] else None
        nc.append(
            [
                r["store"],
                r["sku"],
                r["description"],
                qty,
                r["photo"],
                int(r["sheet_line"]),
                r["what_was_seen"],
                r["note"],
                None,
                "No",
            ]
        )
    nc_last = nc.max_row
    for row in range(3, nc_last + 1):
        nc.cell(row=row, column=9).fill = INPUT_FILL
        for c in range(1, 11):
            nc.cell(row=row, column=c).border = BOX
            nc.cell(row=row, column=c).alignment = Alignment(
                vertical="top", wrap_text=True
            )
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    nc.add_data_validation(dv)
    dv.add(f"J3:J{nc_last}")
    nc.conditional_formatting.add(
        f"A3:J{nc_last}",
        FormulaRule(
            formula=['$J3="Yes"'],
            fill=PatternFill("solid", fgColor="E2EFDA"),
            font=Font(color="375623"),
        ),
    )
    set_widths(nc, [8, 10, 26, 11, 13, 10, 36, 36, 13, 11])
    nc.freeze_panes = "A3"
    landscape(nc)

    # ---------- Counts (Entered) ----------
    cnt_headers = [
        "Store",
        "SKU",
        "Description",
        "Entered Qty",
        "Photo",
        "Sheet Line",
        "Reading",
        "What Was Seen",
        "Note",
        "Confirmed Qty",
        "Final Qty",
        "Has Qty",
    ]
    cnt.append(cnt_headers)
    style_header(cnt, 1, len(cnt_headers))
    NCR = f"'Needs Checking'!$A$3:$A${nc_last}"
    NCS = f"'Needs Checking'!$B$3:$B${nc_last}"
    NCQ = f"'Needs Checking'!$I$3:$I${nc_last}"
    for i, r in enumerate(counts, start=2):
        qty = int(r["counted_qty"]) if r["counted_qty"] else None
        cnt.append(
            [
                r["store"],
                r["sku"],
                r["description"],
                qty,
                r["photo"],
                int(r["sheet_line"]),
                r["reading"],
                r["what_was_seen"],
                r["note"],
            ]
        )
        cnt[f"J{i}"] = (
            f'=IF(COUNTIFS({NCR},A{i},{NCS},B{i},{NCQ},"<>")>0,'
            f'SUMIFS({NCQ},{NCR},A{i},{NCS},B{i}),"")'
        )
        cnt[f"K{i}"] = f'=IF(J{i}<>"",J{i},IF(D{i}="","",D{i}))'
        cnt[f"L{i}"] = f'=IF(K{i}="",0,1)'
    cnt_last = cnt.max_row
    set_widths(cnt, [8, 10, 26, 11, 13, 10, 9, 36, 36, 13, 10, 8])
    cnt.freeze_panes = "A2"
    cnt.auto_filter.ref = f"A1:L{cnt_last}"
    cnt.conditional_formatting.add(
        f"G2:G{cnt_last}",
        FormulaRule(
            formula=['$G2="Check"'],
            fill=PatternFill("solid", fgColor="FFEB9C"),
            font=Font(bold=True, color="9C5700"),
        ),
    )
    cnt.column_dimensions["L"].hidden = True

    # ---------- Reconciliation ----------
    keys = sorted(
        {(r["store"], r["sku"]) for r in counts}
        | {(s["store"], s["sku"]) for s in system}
    )
    desc = {(s["store"], s["sku"]): s["description"] for s in system}
    src = {}
    for r in counts:
        desc.setdefault((r["store"], r["sku"]), r["description"])
        src[(r["store"], r["sku"])] = f"{r['photo']} line {r['sheet_line']}"

    rc_headers = [
        "Store",
        "Store Name",
        "SKU",
        "Description",
        "Counted Qty",
        "System Qty",
        "Difference\n(Counted − System)",
        "Difference %",
        "Status",
        "Reading",
        "Source (photo, line)",
    ]
    recon.append(rc_headers)
    style_header(recon, 1, len(rc_headers))
    recon.row_dimensions[1].height = 48
    C = "'Counts (Entered)'!"
    CA, CB, CK, CL, CG = (
        f"{C}$A$2:$A${cnt_last}",
        f"{C}$B$2:$B${cnt_last}",
        f"{C}$K$2:$K${cnt_last}",
        f"{C}$L$2:$L${cnt_last}",
        f"{C}$G$2:$G${cnt_last}",
    )
    sys_last = sysws.max_row
    SA, SC, SE = (
        f"'System Export'!$A$2:$A${sys_last}",
        f"'System Export'!$C$2:$C${sys_last}",
        f"'System Export'!$E$2:$E${sys_last}",
    )
    for i, (store, sku) in enumerate(keys, start=2):
        recon.append([store, store_names.get(store, ""), sku, desc[(store, sku)]])
        recon[f"E{i}"] = (
            f'=IF(SUMIFS({CL},{CA},A{i},{CB},C{i})=0,"",SUMIFS({CK},{CA},A{i},{CB},C{i}))'
        )
        recon[f"F{i}"] = (
            f'=IF(COUNTIFS({SA},A{i},{SC},C{i})=0,"",SUMIFS({SE},{SA},A{i},{SC},C{i}))'
        )
        recon[f"G{i}"] = f'=IF(OR(E{i}="",F{i}=""),"",E{i}-F{i})'
        recon[f"H{i}"] = f'=IF(OR(G{i}="",F{i}=""),"",IF(F{i}=0,"",G{i}/F{i}))'
        recon[f"I{i}"] = (
            f'=IF(E{i}="","Not counted",IF(F{i}="","Not in system",'
            f'IF(ABS(G{i})<=Summary!$C$4,"Match",IF(G{i}>0,"Over","Under"))))'
        )
        recon[f"J{i}"] = (
            f'=IF(COUNTIFS({CA},A{i},{CB},C{i},{CG},"Check")=0,"Clear",'
            f"IF(COUNTIFS('Needs Checking'!$A$3:$A${nc_last},A{i},"
            f"'Needs Checking'!$B$3:$B${nc_last},C{i},"
            f'\'Needs Checking\'!$J$3:$J${nc_last},"Yes")>0,"Confirmed","Check"))'
        )
        recon[f"K{i}"] = src.get((store, sku), "—")
        recon[f"H{i}"].number_format = "0.0%"
        recon[f"G{i}"].number_format = "+0;-0;0"
    rc_last = recon.max_row
    rng_all = f"A2:K{rc_last}"
    recon.conditional_formatting.add(
        f"J2:J{rc_last}",
        FormulaRule(
            formula=['$J2="Check"'],
            fill=PatternFill("solid", fgColor="FFEB9C"),
            font=Font(bold=True, color="9C5700"),
        ),
    )
    for status, (bg, fg) in STATUS_COLOURS.items():
        recon.conditional_formatting.add(
            rng_all,
            FormulaRule(
                formula=[f'$I2="{status}"'],
                fill=PatternFill("solid", fgColor=bg),
                font=Font(color=fg),
            ),
        )
    for row in recon.iter_rows(min_row=2, max_row=rc_last):
        for cell in row:
            cell.border = BOX
    set_widths(recon, [8, 13, 9, 27, 11, 11, 15, 12, 14, 11, 20])
    recon.freeze_panes = "E2"
    recon.auto_filter.ref = f"A1:K{rc_last}"
    landscape(recon)

    # ---------- Summary ----------
    summary["A1"] = "Stock Count Reconciliation"
    summary["A1"].font = Font(bold=True, size=18, color="1F3864")
    summary["A2"] = "Sample project with fictional data · count date 02 Nov 2026"
    summary["A2"].font = Font(italic=True, color="595959")
    summary["A4"] = "Match tolerance:"
    summary["A4"].font = Font(bold=True)
    summary["C4"] = 0
    summary["C4"].fill, summary["C4"].border = INPUT_FILL, BOX
    summary["C4"].font = Font(bold=True, size=12)
    summary["C4"].alignment = Alignment(horizontal="center")
    summary["D4"] = "units. Differences this size or smaller count as a Match. Change it and every sheet updates."
    summary["D4"].font = Font(italic=True, color="595959")

    R = "Reconciliation!"
    RA, RE, RF, RG, RI, RJ = (
        f"{R}$A$2:$A${rc_last}",
        f"{R}$E$2:$E${rc_last}",
        f"{R}$F$2:$F${rc_last}",
        f"{R}$G$2:$G${rc_last}",
        f"{R}$I$2:$I${rc_last}",
        f"{R}$J$2:$J${rc_last}",
    )
    sm_headers = [
        "Store",
        "Store Name",
        "Items",
        "Counted Units",
        "System Units",
        "Net Difference\n(items in both)",
        "Matching",
        "Under",
        "Over",
        "Not Counted",
        "Not in System",
        "Needs Checking",
    ]
    hdr = 6
    for c, h in enumerate(sm_headers, start=1):
        summary.cell(row=hdr, column=c, value=h)
    style_header(summary, hdr, len(sm_headers))
    for i, store in enumerate(stores, start=hdr + 1):
        summary[f"A{i}"] = store
        summary[f"B{i}"] = store_names.get(store, "")
        summary[f"C{i}"] = f"=COUNTIFS({RA},A{i})"
        summary[f"D{i}"] = f"=SUMIFS({RE},{RA},A{i})"
        summary[f"E{i}"] = f"=SUMIFS({RF},{RA},A{i})"
        summary[f"F{i}"] = f"=SUMIFS({RG},{RA},A{i})"
        for col, status in zip(
            "GHIJK", ["Match", "Under", "Over", "Not counted", "Not in system"]
        ):
            summary[f"{col}{i}"] = f'=COUNTIFS({RA},A{i},{RI},"{status}")'
        summary[f"L{i}"] = f'=COUNTIFS({RA},A{i},{RJ},"Check")'
    first, last = hdr + 1, hdr + len(stores)
    tot = last + 1
    summary[f"A{tot}"] = "TOTAL"
    for col in "CDEFGHIJKL":
        summary[f"{col}{tot}"] = f"=SUM({col}{first}:{col}{last})"
    for c in range(1, len(sm_headers) + 1):
        cell = summary.cell(row=tot, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
    for row in summary.iter_rows(min_row=first, max_row=tot, max_col=len(sm_headers)):
        for cell in row:
            cell.border = BOX
            if cell.column > 2:
                cell.alignment = Alignment(horizontal="center")
    for col in "FGHIJKL":
        summary[f"{col}{first}"].number_format = "General"
    for i in range(first, tot + 1):
        summary[f"F{i}"].number_format = "+0;-0;0"
    colour_cols = {"H": "Under", "I": "Over", "J": "Not counted", "K": "Not in system"}
    for col, status in colour_cols.items():
        bg, fg = STATUS_COLOURS[status]
        summary.conditional_formatting.add(
            f"{col}{first}:{col}{tot}",
            FormulaRule(
                formula=[f"{col}{first}>0"],
                fill=PatternFill("solid", fgColor=bg),
                font=Font(bold=True, color=fg),
            ),
        )
    summary.conditional_formatting.add(
        f"L{first}:L{tot}",
        FormulaRule(
            formula=[f"L{first}>0"],
            fill=PatternFill("solid", fgColor="FFEB9C"),
            font=Font(bold=True, color="9C5700"),
        ),
    )

    leg = tot + 3
    summary[f"A{leg}"] = "Colour key"
    summary[f"A{leg}"].font = Font(bold=True, size=12)
    legend = [
        (
            "Under",
            "Counted LESS than the system says (possible shrinkage, theft or receiving error)",
        ),
        (
            "Over",
            "Counted MORE than the system says (possible unrecorded delivery or wrong item counted)",
        ),
        ("Not counted", "In the system but nothing was written on the count sheet"),
        (
            "Not in system",
            "Counted in store but not in the system export (new or wrong SKU?)",
        ),
        ("Check", "Handwriting was unclear - confirm on the 'Needs Checking' tab"),
    ]
    for j, (name, text) in enumerate(legend, start=leg + 1):
        bg, fg = STATUS_COLOURS.get(name, ("FFEB9C", "9C5700"))
        summary[f"A{j}"] = name
        summary[f"A{j}"].fill = PatternFill("solid", fgColor=bg)
        summary[f"A{j}"].font = Font(bold=True, color=fg)
        summary[f"A{j}"].border = BOX
        summary.merge_cells(f"A{j}:B{j}")
        summary[f"C{j}"] = text
    hw = leg + len(legend) + 2
    summary[f"A{hw}"] = "How to review"
    summary[f"A{hw}"].font = Font(bold=True, size=12)
    steps = [
        "1. Go to 'Needs Checking', look at each photo and type the correct number in the yellow column.",
        "2. Go to 'Reconciliation' and filter the Status column to Under / Over to review differences.",
        "3. Every count shows its source photo and line number, so any number can be traced back.",
    ]
    for j, s in enumerate(steps, start=hw + 1):
        summary[f"A{j}"] = s
    set_widths(summary, [12, 14, 9, 14, 13, 14, 10, 9, 9, 12, 13, 14])
    landscape(summary)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(
        f"Saved {OUT.relative_to(ROOT)}: {len(keys)} reconciliation rows, {len(checks)} to check"
    )


if __name__ == "__main__":
    build()
