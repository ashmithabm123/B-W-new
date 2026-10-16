import logging
import re
from copy import copy
from datetime import datetime
from openpyxl import load_workbook

log = logging.getLogger("remittance")
MONEY_FMT = '"$"#,##0.00'
HEADERS = ["DEFENDANT'S NAME", "PLAINTIFF", "REF NO.", "TOTAL PAYMENT",
           "APPLIED TO PRINCIPAL", "COSTS RETURNED", "CLIENT SHARE",
           "BW, LLP SHARE", "BW, LLP INTEREST", "Notes"]

# ---------------------------------------------------------------------
# Business-defendant rules (client requirement)
# ---------------------------------------------------------------------
BUSINESS_SUFFIXES = {
    "llc", "llp", "lp", "pllc", "pc", "pa", "inc", "incorporated",
    "corp", "corporation", "co", "company", "ltd", "limited", "trust",
    "association", "assoc", "partnership", "holdings",
}   # periods are removed before comparing, so L.L.C. -> llc, P.C. -> pc, P.A. -> pa

DBA_PATTERN = re.compile(r"(?<![a-z0-9])d[./]?b[./]?a\.?(?![a-z0-9])")   # dba, d/b/a, d.b.a.
STARTS_PATTERN = re.compile(r"^(the|city of)(?![a-z0-9])")               # "The ...", "City of ..."


def is_business(name):
    """True if the defendant name looks like a business. Whole words only, not case-sensitive."""
    text = " ".join(str(name).lower().split())
    text = text.rstrip(" .,")                       # ignore trailing periods and commas
    if not text:
        return False

    if "&" in text or any(ch.isdigit() for ch in text):
        return True                                 # ampersand or a number anywhere
    if DBA_PATTERN.search(text):
        return True                                 # DBA / D/B/A anywhere
    if STARTS_PATTERN.match(text):
        return True                                 # begins with 'The' or 'City of'

    last_word = text.split(" ")[-1].replace(".", "").strip(",")
    return last_word in BUSINESS_SUFFIXES           # ends with a business suffix


def _clear_rows(ws, first_row, last_column):
    for row in ws.iter_rows(min_row=first_row, max_row=ws.max_row,
                            max_col=max(ws.max_column, last_column)):
        for cell in row:
            cell.value = None


def _copy_row_style(ws, source_row, target_row, last_column):
    if target_row <= ws.max_row:
        return
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for column in range(1, last_column + 1):
        ws.cell(target_row, column)._style = copy(ws.cell(source_row, column)._style)


def format_payment_sheet(path, df, cfg):
    thru = cfg["pay_thru_date"]
    if thru.upper() == "AUTO":
        thru = datetime.now().strftime("%m-%d-%Y")
    pct = cfg["client_share_pct"]

    include_cost_only = cfg.get("include_cost_only", False)
    fees_as_principal = cfg.get("fees_as_principal", True)
    applied_from_payment = cfg.get("applied_from_payment", False)
    skip_zero_payment = cfg.get("skip_zero_payment", False)
    skip_interest_only = cfg.get("skip_interest_only", False)
    total_from_components = cfg.get("total_from_components", False)
    applied_formula = cfg.get("applied_formula", False)
    separate_interest_only = cfg.get("separate_interest_only", False)
    master_share_calculation = cfg.get("master_share_calculation", False)
    skip_zero_total_payments = cfg.get("skip_zero_total_payments", False)
    log.info("DEBUG writer=%s | skip_interest_only=%s | applied_from_payment=%s | total_from_components=%s | applied_formula=%s",
             __file__, skip_interest_only, applied_from_payment, total_from_components, applied_formula)
    thru_word = cfg.get("thru_word", "Thru")
    pct_fmt = cfg.get("pct_fmt", "0%")

    rows = []
    interest_only_rows = []
    for _, r in df.iterrows():
        total_pay = round(float(r.iloc[cfg["pay_total_col"]]), 2)
        if skip_zero_payment and total_pay <= 0:
            continue

        principal = round(float(r.iloc[cfg["pay_principal_col"]]), 2)
        interest = float(r.iloc[cfg["pay_interest_col"]])
        fees = float(r.iloc[7])
        costs = round(float(r.iloc[cfg["pay_costs_col"]]), 2)
        other = round(float(r["Other"]), 2)

        report_principal = principal
        is_interest_only = report_principal == 0 and round(interest, 2) != 0
        if skip_interest_only and report_principal == 0 and fees == 0 and costs == 0:
            continue                    # interest-only row: not a principal payment

        if applied_from_payment:
            principal = round(total_pay - costs - round(interest, 2), 2)

        if applied_formula:
            # Applied = Principal + Fees - Direct Payments (report column 4)
            direct = float(r.iloc[4])
            principal = round(principal + fees - direct, 2)
        elif fees_as_principal and principal == 0 and interest == 0 and fees != 0:
            principal = round(fees, 2)      # Fees-only payment counts as applied amount

        if total_from_components:
            total_pay = round(principal + round(interest, 2) + costs + other, 2)
        if skip_zero_total_payments and total_pay == 0:
            continue

        if not str(r["Claim"]).strip():
            client = round(float(r["Client Proceeds"]), 2)
            firm_share = round(float(r["Fees Due"]), 2)
        elif master_share_calculation:
            client = round(principal * pct, 2)
            firm_share = round(principal * cfg["bw_share_pct"], 2)
        else:
            client = round(principal * pct, 2)
            firm_share = round(principal - client, 2)

        row = [
            r["Debtor Name"], r["Plaintiff"], r["Ref Acct #"],
            total_pay,
            principal,
            costs,
            client,
            firm_share,
            round(interest * cfg["bw_interest_pct"], 2)
            if master_share_calculation else round(interest, 2),
            "",
        ]

        if separate_interest_only and is_interest_only:
            interest_only_rows.append(row)
            continue

        if include_cost_only:
            if round(principal + costs, 2) == 0:
                continue
        elif round(principal + client, 2) == 0:      # current rule, unchanged
            continue

        business = is_business(r["Debtor Name"])
        # 'business' is True for business defendants. Add the client's required
        # handling here (e.g. keep the name unchanged, set a note, etc.).

        rows.append(row)

    wb = load_workbook(path)
    required_sheets = {"Payments", "IntOnly", "Summary1", "Summary2",
                       "Summary3", "Summary4"}
    missing_sheets = required_sheets.difference(wb.sheetnames)
    if missing_sheets:
        raise ValueError(f"Output template is missing required sheets: {sorted(missing_sheets)}")

    ws = wb["Payments"]
    data_start = 9
    style_source_row = data_start
    _clear_rows(ws, data_start, len(HEADERS))
    ws["A1"] = f"{cfg['pay_title'].upper()} PAYMENT SHEET"
    ws["A2"] = f"Payments {thru_word} {thru}"
    ws["G4"] = pct
    ws["H4"] = round(1 - pct, 4)
    for c in ("G4", "H4"):
        ws[c].number_format = pct_fmt
    for index, row in enumerate(rows, start=data_start):
        _copy_row_style(ws, style_source_row, index, len(HEADERS))
        for column, value in enumerate(row, start=1):
            ws.cell(row=index, column=column, value=value)

    int_ws = wb["IntOnly"]
    int_data_start = 8
    _clear_rows(int_ws, int_data_start, 3)
    int_ws["A1"] = cfg.get("intonly_title", cfg["pay_title"])
    int_ws["A2"] = f"Interest Only Payments {thru_word} {thru}"
    int_ws["B5"] = "Total"
    int_ws["C5"] = "=SUM(C8:C1000)"
    int_ws["C5"].number_format = MONEY_FMT
    for index, row in enumerate(interest_only_rows, start=int_data_start):
        _copy_row_style(int_ws, int_data_start, index, 3)
        for column, value in enumerate((row[0], row[1], row[8]), start=1):
            int_ws.cell(row=index, column=column, value=value)

    if thru.upper() == "AUTO":
        thru = datetime.now().strftime("%m-%d-%Y")
    for sheet_name in ("Summary1", "Summary2", "Summary3", "Summary4"):
        summary = wb[sheet_name]
        summary["A15"] = cfg["pay_title"].upper()
        summary["A20"] = f"Statement Thru {thru}"

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)
    wb.close()
    log.info("Template payment sheet formatted with %d rows and %d interest-only rows",
             len(rows), len(interest_only_rows))