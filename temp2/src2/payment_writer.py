import logging
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

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
    old = wb.worksheets[0]
    name = old.title
    wb.remove(old)
    ws = wb.create_sheet(name, 0)
    bold = Font(bold=True)

    ws["A1"] = cfg["pay_title"].upper()
    ws["A1"].font = bold
    ws["A2"] = f"Payments {thru_word} {thru}"
    ws["A2"].font = bold

    ws["G4"] = pct
    ws["H4"] = round(1 - pct, 4)
    for c in ("G4", "H4"):
        ws[c].number_format = pct_fmt
        ws[c].font = bold

    ws["C5"] = "Total"
    ws["C5"].font = bold
    for col in range(4, 10):                      # D..I
        total = round(sum(r[col - 1] for r in rows), 2)
        cell = ws.cell(row=5, column=col, value=total)
        cell.number_format = MONEY_FMT
        cell.font = bold

    for c, h in enumerate(HEADERS, start=1):
        ws.cell(row=7, column=c, value=h).font = bold

    for i, row in enumerate(rows, start=8):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            if 4 <= c <= 9:
                cell.number_format = MONEY_FMT

    for i, w in enumerate([26, 45, 12, 16, 22, 16, 15, 15, 18, 18], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    if separate_interest_only:
        if "IntOnly" in wb.sheetnames:
            wb.remove(wb["IntOnly"])
        int_ws = wb.create_sheet("IntOnly")
        if cfg.get("intonly_compact"):
            int_ws["A1"] = cfg.get("intonly_title", cfg["pay_title"])
            int_ws["A1"].font = bold
            int_ws["A2"] = f"Interest Only Payments {thru_word} {thru}"
            int_ws["A2"].font = bold
            int_ws["B5"] = "Total"
            int_ws["B5"].font = bold
            total_cell = int_ws.cell(
                row=5, column=3,
                value=round(sum(r[8] for r in interest_only_rows), 2))
            total_cell.number_format = MONEY_FMT
            total_cell.font = bold
            for col, heading in enumerate(
                    ("DEFENDANT", "PLAINTIFF", "BW LLP INTEREST"), start=1):
                int_ws.cell(row=7, column=col, value=heading).font = bold
            for i, row in enumerate(interest_only_rows, start=8):
                for col, value in enumerate((row[0], row[1], row[8]), start=1):
                    cell = int_ws.cell(row=i, column=col, value=value)
                    if col == 3:
                        cell.number_format = MONEY_FMT
            for col, width in enumerate((30, 55, 20), start=1):
                int_ws.column_dimensions[get_column_letter(col)].width = width
        else:
            int_ws["A1"] = f"{cfg['pay_title'].upper()} - INTEREST ONLY"
            int_ws["A1"].font = bold
            int_ws["A2"] = f"Payments {thru_word} {thru}"
            int_ws["A2"].font = bold
            int_ws["G4"] = pct
            int_ws["H4"] = round(1 - pct, 4)
            for c in ("G4", "H4"):
                int_ws[c].number_format = pct_fmt
                int_ws[c].font = bold
            int_ws["C5"] = "Total"
            int_ws["C5"].font = bold
            for col in range(4, 10):
                cell = int_ws.cell(row=5, column=col,
                                   value=round(sum(r[col - 1] for r in interest_only_rows), 2))
                cell.number_format = MONEY_FMT
                cell.font = bold
            for c, h in enumerate(HEADERS, start=1):
                int_ws.cell(row=7, column=c, value=h).font = bold
            for i, row in enumerate(interest_only_rows, start=8):
                for c, v in enumerate(row, start=1):
                    cell = int_ws.cell(row=i, column=c, value=v)
                    if 4 <= c <= 9:
                        cell.number_format = MONEY_FMT
            for i, w in enumerate([26, 45, 12, 16, 22, 16, 15, 15, 18, 18], start=1):
                int_ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)
    log.info("First sheet formatted with %d rows and %d interest-only rows",
             len(rows), len(interest_only_rows))