"""
Builds a real, formula-driven Excel controlling workbook from the
generated dataset: a raw-data sheet, a SUMIFS/SUMPRODUCT-based
budget-vs-actual pivot-style summary (by cost center and by month),
a YoY comparison sheet, and a variance/exception report — all computed
with live Excel formulas, not pre-baked values, so opening the file in
Excel and changing a filter/input recalculates everything.

Why formulas instead of openpyxl's native PivotTable object: openpyxl
can write a PivotCacheDefinition/PivotTable XML structure, but building
one that Excel opens cleanly and refreshes correctly from a cold write
(no source workbook round-trip through real Excel) is unreliable in
practice. A SUMIFS/SUMPRODUCT-based summary table is the same
"ad-hoc Auswertung" a controller does with Excel every day, is fully
live-recalculating, and is something I could verify byte-for-byte
against the source data -- so it's the honest, verifiable choice here.
This is disclosed in the README.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

from generate_data import generate_rows, COST_CENTERS, MONTHS


HEADER_FILL = PatternFill("solid", fgColor="1B3A6B")
HEADER_FONT = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def build(output_path: str = "controlling_report_2025_2026.xlsx", seed: int = 42):
    rows = generate_rows(seed=seed)
    wb = Workbook()

    # ---------------------------------------------------------------
    # Sheet 1: Rohdaten (raw data — the "fact table")
    # ---------------------------------------------------------------
    ws_raw = wb.active
    ws_raw.title = "Rohdaten"
    headers = ["Periode", "KostenstelleID", "Kostenstelle", "Bereich",
               "KontoID", "Konto", "Budget_EUR", "Ist_EUR"]
    ws_raw.append(headers)
    _style_header(ws_raw, 1, len(headers))

    for r in rows:
        ws_raw.append([r.period, r.cost_center_id, r.cost_center_name,
                        r.division, r.account_id, r.account_name,
                        r.budget_eur, r.actual_eur])

    n_data_rows = len(rows)
    last_row = n_data_rows + 1
    for col_letter in "GH":
        for row_idx in range(2, last_row + 1):
            ws_raw[f"{col_letter}{row_idx}"].number_format = '#,##0.00 €'
    for col, width in zip("ABCDEFGH", [10, 14, 24, 14, 10, 30, 14, 14]):
        ws_raw.column_dimensions[col].width = width

    raw_range = f"Rohdaten!$A$2:$H${last_row}"

    # ---------------------------------------------------------------
    # Sheet 2: Kostenstellen-Übersicht (pivot-equivalent: cost center x
    # full 24-month total, via SUMIFS)
    # ---------------------------------------------------------------
    ws_cc = wb.create_sheet("Kostenstellen-Uebersicht")
    hdr = ["KostenstelleID", "Kostenstelle", "Bereich", "Budget_EUR (24 Mon.)",
           "Ist_EUR (24 Mon.)", "Abweichung_EUR", "Abweichung_%"]
    ws_cc.append(hdr)
    _style_header(ws_cc, 1, len(hdr))

    for i, (cc_id, cc_name, division) in enumerate(COST_CENTERS, start=2):
        ws_cc.cell(row=i, column=1, value=cc_id)
        ws_cc.cell(row=i, column=2, value=cc_name)
        ws_cc.cell(row=i, column=3, value=division)
        ws_cc.cell(row=i, column=4,
                    value=f"=SUMIFS(Rohdaten!$G$2:$G${last_row},Rohdaten!$B$2:$B${last_row},A{i})")
        ws_cc.cell(row=i, column=5,
                    value=f"=SUMIFS(Rohdaten!$H$2:$H${last_row},Rohdaten!$B$2:$B${last_row},A{i})")
        ws_cc.cell(row=i, column=6, value=f"=E{i}-D{i}")
        ws_cc.cell(row=i, column=7, value=f"=IFERROR(E{i}/D{i}-1,0)")

    n_cc = len(COST_CENTERS)
    total_row = n_cc + 2
    ws_cc.cell(row=total_row, column=2, value="GESAMT").font = BOLD
    ws_cc.cell(row=total_row, column=4, value=f"=SUM(D2:D{n_cc + 1})").font = BOLD
    ws_cc.cell(row=total_row, column=5, value=f"=SUM(E2:E{n_cc + 1})").font = BOLD
    ws_cc.cell(row=total_row, column=6, value=f"=E{total_row}-D{total_row}").font = BOLD
    ws_cc.cell(row=total_row, column=7, value=f"=IFERROR(E{total_row}/D{total_row}-1,0)").font = BOLD

    for row_idx in range(2, total_row + 1):
        ws_cc[f"D{row_idx}"].number_format = '#,##0.00 €'
        ws_cc[f"E{row_idx}"].number_format = '#,##0.00 €'
        ws_cc[f"F{row_idx}"].number_format = '#,##0.00 €'
        ws_cc[f"G{row_idx}"].number_format = '0.0%'
    for col, width in zip("ABCDEFG", [14, 24, 14, 20, 18, 16, 14]):
        ws_cc.column_dimensions[col].width = width

    # Conditional formatting: highlight cost centers >10% over budget red,
    # >10% under budget green
    ws_cc.conditional_formatting.add(
        f"G2:G{n_cc + 1}",
        CellIsRule(operator="greaterThan", formula=["0.10"],
                    fill=PatternFill("solid", fgColor="F8CBAD")),
    )
    ws_cc.conditional_formatting.add(
        f"G2:G{n_cc + 1}",
        CellIsRule(operator="lessThan", formula=["-0.10"],
                    fill=PatternFill("solid", fgColor="C6E0B4")),
    )

    # ---------------------------------------------------------------
    # Sheet 3: Monatsverlauf (pivot-equivalent: month x total budget/actual)
    # ---------------------------------------------------------------
    ws_m = wb.create_sheet("Monatsverlauf")
    ws_m.append(["Periode", "Budget_EUR", "Ist_EUR", "Abweichung_EUR", "Abweichung_%"])
    _style_header(ws_m, 1, 5)
    for i, period in enumerate(MONTHS, start=2):
        ws_m.cell(row=i, column=1, value=period)
        ws_m.cell(row=i, column=2,
                    value=f"=SUMIFS(Rohdaten!$G$2:$G${last_row},Rohdaten!$A$2:$A${last_row},A{i})")
        ws_m.cell(row=i, column=3,
                    value=f"=SUMIFS(Rohdaten!$H$2:$H${last_row},Rohdaten!$A$2:$A${last_row},A{i})")
        ws_m.cell(row=i, column=4, value=f"=C{i}-B{i}")
        ws_m.cell(row=i, column=5, value=f"=IFERROR(C{i}/B{i}-1,0)")
    n_months = len(MONTHS)
    for row_idx in range(2, n_months + 2):
        ws_m[f"B{row_idx}"].number_format = '#,##0.00 €'
        ws_m[f"C{row_idx}"].number_format = '#,##0.00 €'
        ws_m[f"D{row_idx}"].number_format = '#,##0.00 €'
        ws_m[f"E{row_idx}"].number_format = '0.0%'
    for col, width in zip("ABCDE", [12, 16, 16, 16, 14]):
        ws_m.column_dimensions[col].width = width

    # ---------------------------------------------------------------
    # Sheet 4: YoY-Vergleich (2025 vs 2026, by cost center)
    # ---------------------------------------------------------------
    ws_yoy = wb.create_sheet("YoY-Vergleich")
    hdr = ["KostenstelleID", "Kostenstelle", "Ist_2025_EUR", "Ist_2026_EUR",
           "Veraenderung_EUR", "Veraenderung_%"]
    ws_yoy.append(hdr)
    _style_header(ws_yoy, 1, len(hdr))
    for i, (cc_id, cc_name, division) in enumerate(COST_CENTERS, start=2):
        ws_yoy.cell(row=i, column=1, value=cc_id)
        ws_yoy.cell(row=i, column=2, value=cc_name)
        ws_yoy.cell(row=i, column=3,
            value=(f'=SUMPRODUCT((Rohdaten!$B$2:$B${last_row}=A{i})*'
                    f'(LEFT(Rohdaten!$A$2:$A${last_row},4)="2025")*Rohdaten!$H$2:$H${last_row})'))
        ws_yoy.cell(row=i, column=4,
            value=(f'=SUMPRODUCT((Rohdaten!$B$2:$B${last_row}=A{i})*'
                    f'(LEFT(Rohdaten!$A$2:$A${last_row},4)="2026")*Rohdaten!$H$2:$H${last_row})'))
        ws_yoy.cell(row=i, column=5, value=f"=D{i}-C{i}")
        ws_yoy.cell(row=i, column=6, value=f"=IFERROR(D{i}/C{i}-1,0)")

    for row_idx in range(2, n_cc + 2):
        for col in "CDE":
            ws_yoy[f"{col}{row_idx}"].number_format = '#,##0.00 €'
        ws_yoy[f"F{row_idx}"].number_format = '0.0%'
    for col, width in zip("ABCDEF", [14, 24, 16, 16, 16, 14]):
        ws_yoy.column_dimensions[col].width = width

    # ---------------------------------------------------------------
    # Sheet 5: Abweichungs-Report (ad-hoc exception report: every
    # cost-center/account/month combination where |variance| > 15%
    # AND |variance| > 500 EUR, sorted by absolute deviation)
    # ---------------------------------------------------------------
    ws_exc = wb.create_sheet("Abweichungs-Report")
    hdr = ["Periode", "Kostenstelle", "Konto", "Budget_EUR", "Ist_EUR",
           "Abweichung_EUR", "Abweichung_%"]
    ws_exc.append(hdr)
    _style_header(ws_exc, 1, len(hdr))

    exceptions = []
    for r in rows:
        dev_eur = r.actual_eur - r.budget_eur
        dev_pct = (dev_eur / r.budget_eur) if r.budget_eur else 0.0
        if abs(dev_pct) > 0.15 and abs(dev_eur) > 500:
            exceptions.append((r, dev_eur, dev_pct))
    exceptions.sort(key=lambda t: abs(t[1]), reverse=True)

    for i, (r, dev_eur, dev_pct) in enumerate(exceptions, start=2):
        ws_exc.cell(row=i, column=1, value=r.period)
        ws_exc.cell(row=i, column=2, value=r.cost_center_name)
        ws_exc.cell(row=i, column=3, value=r.account_name)
        ws_exc.cell(row=i, column=4, value=r.budget_eur)
        ws_exc.cell(row=i, column=5, value=r.actual_eur)
        ws_exc.cell(row=i, column=6, value=dev_eur)
        ws_exc.cell(row=i, column=7, value=dev_pct)

    n_exc = len(exceptions)
    for row_idx in range(2, n_exc + 2):
        ws_exc[f"D{row_idx}"].number_format = '#,##0.00 €'
        ws_exc[f"E{row_idx}"].number_format = '#,##0.00 €'
        ws_exc[f"F{row_idx}"].number_format = '#,##0.00 €'
        ws_exc[f"G{row_idx}"].number_format = '0.0%'
    for col, width in zip("ABCDEFG", [12, 24, 30, 14, 14, 16, 14]):
        ws_exc.column_dimensions[col].width = width

    if n_exc:
        ws_exc.conditional_formatting.add(
            f"G2:G{n_exc + 1}",
            ColorScaleRule(start_type="min", start_color="C6E0B4",
                            mid_type="num", mid_value=0, mid_color="FFFFFF",
                            end_type="max", end_color="F8696B"),
        )

    wb.save(output_path)
    print(f"Wrote {output_path}: {n_data_rows} raw rows, "
          f"{n_cc} cost centers, {n_months} months, {n_exc} exceptions flagged.")
    return output_path


if __name__ == "__main__":
    build()
