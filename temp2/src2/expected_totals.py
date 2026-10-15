import pandas as pd

from temp2.src2.reader import _is_totals_row, _to_number, _pick_sheet

LEADING_COLUMNS = 3    # Claim, Debtor Name, Ref Acct # come before the report columns


def find_expected_totals(path, sheet_name, columns, principal_idx, interest_idx,
                         cost_marker):
    result = {"principal": None, "interest": None, "cost": None}
    n = len(columns)

    with pd.ExcelFile(path) as xls:
        sheet = _pick_sheet(xls, sheet_name)
        raw = xls.parse(sheet, header=None)

    for _, row in raw.iterrows():
        vals = [v for v in row.tolist() if pd.notna(v) and str(v).strip() != ""]
        if not vals:
            continue

        if result["principal"] is None and _is_totals_row(vals, n):
            nums = [_to_number(v) for v in vals]
            result["principal"] = round(nums[principal_idx - LEADING_COLUMNS], 2)
            result["interest"] = round(nums[interest_idx - LEADING_COLUMNS], 2)

        if (result["cost"] is None and len(vals) > 1
                and str(vals[0]).strip().lower().startswith(cost_marker.lower())):
            result["cost"] = _to_number(vals[1])

    return result