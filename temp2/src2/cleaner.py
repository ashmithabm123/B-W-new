import logging
import pandas as pd

log = logging.getLogger("remittance")


def keep_valid_rows(df: pd.DataFrame, col_a: int, col_b: int,
                    rule_mode: str = "BOTH", fees_col: int = 7,
                    include_cost_only: bool = False,
                    fees_as_principal: bool = True,
                    skip_interest_only: bool = False,
                    include_other_only: bool = False) -> pd.DataFrame:
    a = df.iloc[:, col_a]
    b = df.iloc[:, col_b]
    mode = rule_mode.upper()

    if mode == "SUM":                 # Principal + Interest not 0
        mask = (a + b).round(2) != 0
        if fees_as_principal:         # Fees-only payment (current behaviour)
            fees = df.iloc[:, fees_col]
            mask = mask | ((a == 0) & (b == 0) & (fees != 0))
        if include_cost_only:         # only when enabled for this client
            mask = mask | (df.iloc[:, 8] != 0)
        if include_other_only:
            mask = mask | (df["Other"] != 0)
    elif mode == "ANY":               # either one is non-zero
        mask = (a != 0) | (b != 0)
    else:                             # BOTH: both columns non-zero
        mask = (a != 0) & (b != 0)

    if skip_interest_only:            # interest-only row: not a principal payment
        fees = df.iloc[:, fees_col]
        costs = df.iloc[:, 8]
        interest_only = (a == 0) & (fees == 0) & (costs == 0)
        mask = mask & ~interest_only

    valid = df[mask].reset_index(drop=True)
    log.info("Rule: %s | columns %d and %d | total: %d | valid: %d | removed: %d",
             mode, col_a, col_b, len(df), len(valid), len(df) - len(valid))
    return valid