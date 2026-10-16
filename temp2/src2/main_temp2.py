from pathlib import Path
from shutil import copyfile

import pandas as pd

from config.config_loader import load_config
from logs.logger_setup import setup_logger
from temp2.src2.reader import (
    _is_totals_row,
    _pick_sheet,
    _to_number,
    read_report,
)
from temp2.src2.validator import check_totals, check_cost_total
from temp2.src2.cleaner import keep_valid_rows
from temp2.src2.client_share import get_client_share
from temp2.src2.payment_writer import format_payment_sheet
from temp2.src2.cost_reader import read_costs
from temp2.src2.cost_writer import write_cost_sheet

# ---- Fixed logic values (not in config.xlsx) ----
PRINCIPAL_COL = 5            # 0-based column position of Principal
INTEREST_COL = 6             # 0-based column position of Interest
SHARE_COLUMN = "ClientShare"  # master table column used for every client
LEADING_COLUMNS = 3  # Claim, Debtor Name, Ref Acct # precede report columns

# Options applied to every file
COMMON_OPTIONS = {
    "include_cost_only": True,
    "applied_from_payment": True,
    "thru_word": "Through",
    "pct_fmt": "0.0%",
}

# Exceptions only. The key must match the input file name (without extension).
CLIENT_OVERRIDES = {
    "Executive Financial Consultants": {
        "applied_from_payment": False,
        "applied_formula": True,
        "skip_interest_only": False,
        "total_from_components": True,
        "separate_interest_only": True,
        "master_share_calculation": True,
    },
    "NKC Health": {"fees_as_principal": False},
    "Dentistry for Children": {"skip_zero_payment": True},
    "Consumer Collection Management, Inc.": {"skip_interest_only": True,
                                             "applied_from_payment": False,
                                             "total_from_components": True,
                                             "applied_formula": True},
    "Consumer Collection Management, Inc. 2": {
        "applied_from_payment": False,
        "applied_formula": True,
        "skip_interest_only": False,
        "total_from_components": True,
        "skip_zero_total_payments": True,
        "separate_interest_only": True,
        "intonly_compact": True,
        "intonly_title": "Consumer Collection Management",
        "master_share_calculation": True,
        "include_other_only": True,
    },
}

OPTION_KEYS = ("include_cost_only", "fees_as_principal", "thru_word",
               "pct_fmt", "applied_from_payment", "skip_zero_payment",
               "skip_interest_only", "total_from_components", "applied_formula",
               "separate_interest_only", "master_share_calculation",
               "skip_zero_total_payments", "intonly_compact", "intonly_title")


def _text(value, default):
    if isinstance(value, str) and value.strip() and value.strip().lower() != "nan":
        return value.strip()
    return default


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


def process_file(input_path: Path, config: dict, out_dir: Path, log):
    client_name = input_path.stem
    cfg = dict(config)                       # copy, so each file starts clean
    cfg["input_file"] = str(input_path)
    cfg["output_file"] = str(out_dir / f"{client_name}.xlsx")
    cfg["pay_title"] = client_name.upper()
    cfg["cost_title"] = client_name

    # ---- Options: common first, then client exceptions ----
    opts = {**COMMON_OPTIONS, **CLIENT_OVERRIDES.get(client_name, {})}
    for k in OPTION_KEYS:
        if k in opts:
            cfg[k] = opts[k]
    log.info("DEBUG client_name=%r | override found=%s", client_name, client_name in CLIENT_OVERRIDES)

    log.info("----- Processing: %s -----", input_path.name)

    # Client share comes only from the master table (error if client missing)
    cfg["client_share_pct"] = get_client_share(
        cfg.get("master_file", "input/MasterDatatable_1.xlsx"),
        cfg.get("master_sheet", "IntrestPercentage_Temp2"),
        client_name,
        SHARE_COLUMN,
    )
    if cfg.get("master_share_calculation"):
        master_file = cfg.get("master_file", "input/MasterDatatable_1.xlsx")
        master_sheet = cfg.get("master_sheet", "IntrestPercentage_Temp2")
        cfg["bw_share_pct"] = get_client_share(
            master_file, master_sheet, client_name, "BWShare")
        cfg["bw_interest_pct"] = get_client_share(
            master_file, master_sheet, client_name, "BW Interest %")

    expected = find_expected_totals(
        cfg["input_file"], cfg["sheet_name"], cfg["report_columns"],
        PRINCIPAL_COL, INTEREST_COL, cfg["cost_end_marker"])

    df = read_report(cfg["input_file"], cfg["sheet_name"], cfg["report_columns"])

    if expected["principal"] is not None:
        if not check_totals(df, PRINCIPAL_COL, INTEREST_COL,
                            expected["principal"], expected["interest"]):
            log.warning("Totals do not match the report for %s", input_path.name)
    else:
        log.warning("No totals row found in %s, skipping the PASS check", input_path.name)

    valid = keep_valid_rows(df, PRINCIPAL_COL, INTEREST_COL, cfg["rule_mode"],
                            include_cost_only=cfg.get("include_cost_only", False),
                            fees_as_principal=cfg.get("fees_as_principal", True),
                            skip_interest_only=cfg.get("skip_interest_only", False),
                            include_other_only=cfg.get("include_other_only", False))
    for _, r in valid.iterrows():
        if abs(r.iloc[3] - 50.69) < 0.005 or abs(r.iloc[5] - 50.69) < 0.005:
            log.info("DEBUG 50.69: %s | pay %.2f | prin %.2f | int %.2f | fees %.2f | costs %.2f",
                     r.iloc[1], r.iloc[3], r.iloc[5], r.iloc[6], r.iloc[7], r.iloc[8])

    # ---- Temporary debug: remove once the extra rows are found ----
    a = df.iloc[:, PRINCIPAL_COL]
    b = df.iloc[:, INTEREST_COL]
    costs = df.iloc[:, 8]
    added = df[((a + b).round(2) == 0) & (costs != 0)]
    log.info("DEBUG rows added by cost-only rule: %d", len(added))
    for _, r in added.iterrows():
        log.info("DEBUG added: %s | pay %.2f | prin %.2f | int %.2f | fees %.2f | costs %.2f",
                 r.iloc[1], r.iloc[3], r.iloc[5], r.iloc[6], r.iloc[7], r.iloc[8])
    for _, r in valid[valid.iloc[:, 3] <= 0].iterrows():
        log.info("DEBUG pay<=0: %s | pay %.2f | prin %.2f | int %.2f | costs %.2f",
                 r.iloc[1], r.iloc[3], r.iloc[5], r.iloc[6], r.iloc[8])

    out_dir.mkdir(parents=True, exist_ok=True)
    template = Path(str(cfg["output_template_temp2"]).strip().strip('"'))
    if not template.is_file():
        raise FileNotFoundError(f"Temp2 output template not found: {template}")
    copyfile(template, cfg["output_file"])
    log.info("Copied Temp2 output template from %s to %s",
             template, cfg["output_file"])

    format_payment_sheet(cfg["output_file"], valid, cfg)

    # ---- Cost sheet ----
    plaintiff_map = dict(zip(df["Claim"], df["Plaintiff"]))
    cost_df = read_costs(cfg["input_file"], cfg["cost_identifier"],
                         cfg["cost_end_marker"], plaintiff_map)
    if expected["cost"] is not None:
        check_cost_total(cost_df, expected["cost"])
    else:
        log.warning("No cost total found in %s, skipping the cost check", input_path.name)

    write_cost_sheet(cfg["output_file"], cfg["cost_sheet_name"],
                     cost_df, cfg["cost_title"], cfg["cost_thru_date"])
    log.info("Cost sheet added to %s", cfg["output_file"])

    print(f"{input_path.name}: valid rows {len(valid)}, cost rows {len(cost_df)} "
          f"-> {cfg['output_file']}")


def run_temp2():
    config = load_config()
    log = setup_logger(config["log_file"], config["log_level"])
    log.info("===== Run started =====")

    folder = Path(_text(config.get("input_folder"), "input"))
    out_dir = Path(config["output_file"]).parent
    master = Path(_text(config.get("master_file"), "input/MasterDatatable_1.xlsx"))

    files = sorted(
        p for p in folder.iterdir()
        if p.suffix.lower() in (".xls", ".xlsx")
        and not p.name.startswith("~$")
        and p.resolve() != master.resolve()
    )
    log.info("Input files found: %d -> %s", len(files), [p.name for p in files])

    done, failed = [], []
    for path in files:
        try:
            process_file(path, config, out_dir, log)
            done.append(path.name)
        except Exception:
            log.exception("FAILED: %s", path.name)
            failed.append(path.name)

    log.info("Finished. Success: %d | Failed: %d %s", len(done), len(failed), failed)
    print(f"\nSuccess: {len(done)} | Failed: {len(failed)} {failed}")
    log.info("===== Run finished =====")


if __name__ == "__main__":
    run_temp2()