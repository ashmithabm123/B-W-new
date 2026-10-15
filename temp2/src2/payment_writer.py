import logging
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

log = logging.getLogger("remittance")
MONEY_FMT = '"$"#,##0.00'
HEADERS = ["DEFENDANT'S NAME", "PLAINTIFF", "REF NO.", "TOTAL PAYMENT",
           "APPLIED TO PRINCIPAL", "COSTS RETURNED", "CLIENT SHARE",
           "BW, LLP SHARE", "BW, LLP INTEREST", "Notes"]


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
    log.info("DEBUG skip_interest_only=%s", skip_interest_only)
    thru_word = cfg.get("thru_word", "Thru")
    pct_fmt = cfg.get("pct_fmt", "0%")

    rows = []
    for _, r in df.iterrows():
        total_pay = round(float(r.iloc[cfg["pay_total_col"]]), 2)
        if skip_zero_payment and total_pay <= 0:
            continue

        principal = round(float(r.iloc[cfg["pay_principal_col"]]), 2)
        interest = float(r.iloc[cfg["pay_interest_col"]])
        fees = float(r.iloc[7])
        costs = round(float(r.iloc[cfg["pay_costs_col"]]), 2)

        report_principal = principal
        if skip_interest_only and report_principal == 0 and fees == 0 and costs == 0:
            continue                    # interest-only row: not a principal payment

        if applied_from_payment:
            principal = round(total_pay - costs - round(interest, 2), 2)

        if fees_as_principal and principal == 0 and interest == 0 and fees != 0:
            principal = round(fees, 2)      # Fees-only payment counts as applied amount

        if not str(r["Claim"]).strip():
            client = round(float(r["Client Proceeds"]), 2)
            firm_share = round(float(r["Fees Due"]), 2)
        else:
            client = round(principal * pct, 2)
            firm_share = round(principal - client, 2)

        if include_cost_only:
            if round(principal + costs, 2) == 0:
                continue
        elif round(principal + client, 2) == 0:      # current rule, unchanged
            continue

        rows.append([
            r["Debtor Name"], r["Plaintiff"], r["Ref Acct #"],
            total_pay,
            principal,
            costs,
            client,
            firm_share,
            round(interest, 2),
            "",
        ])

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

    wb.save(path)
    log.info("First sheet formatted with %d rows", len(rows))