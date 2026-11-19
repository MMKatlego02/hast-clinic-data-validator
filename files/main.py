"""
main.py
=======
Entry point of the HAST Data Quality Guardian.

Pipeline
--------
1. Read the messy clinic export (text only, nothing auto-converted).
2. Normalise spacing / null tokens / letter case (every change is logged).
3. Validate every record against the five rules (see validators.py).
4. Select the clean records (1 row per patient_id) and re-verify them.
5. Reconcile counts (input rows = accepted + rejected).
6. Write hast_data_clean.csv, errors_to_fix.csv, data_quality_report.xlsx
   and cleaning_log.csv.

Nothing is written until every check has passed, so a failed run never leaves
a mixture of new and old output files caused by a validation fault.

Exit codes: 0 = success, 1 = unexpected error, 2 = problem with the input file,
3 = internal integrity check failed (do NOT use any outputs).

Usage
-----
    python main.py
    python main.py --input C:\\Users\\Kat\\hast\\hast_data_extreme_5000.csv --output-dir output --reference-date 2026-11-14
"""

from __future__ import annotations
import argparse
import logging
import sys
from datetime import date, datetime, timezone
from pathlib import Path

try:
    import pandas as pd

    import utils
    import validators
except ImportError as exc:  # clear message instead of a raw traceback
    sys.exit(
        f"ERROR: a required package is missing ({exc}).\n"
        "Install the dependencies with:  pip install -r requirements.txt"
    )

logger = logging.getLogger("hast_guardian")

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = Path(r"C:\Users\Kat\hast\hast_data_extreme_5000.csv")
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

CLEAN_FILE = "hast_data_clean.csv"
ERRORS_FILE = "errors_to_fix.csv"
REPORT_FILE = "data_quality_report.xlsx"
LOG_FILE = "cleaning_log.csv"
RUN_LOG_FILE = "pipeline_run.log"


def iso_date_arg(text: str) -> date:
    """argparse type: parse a YYYY-MM-DD command-line date or raise a clear error."""
    try:
        parsed = utils.parse_iso_date(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"'{text}' {exc}") from exc
    if parsed is None:
        raise argparse.ArgumentTypeError("a date is required (YYYY-MM-DD)")
    return parsed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments (input file, output folder, reference date)."""
    parser = argparse.ArgumentParser(
        description="HAST Data Quality Guardian - clean a clinic export and report data-quality problems."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help=f"input CSV (default: {DEFAULT_INPUT})")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help=f"output folder (default: {DEFAULT_OUTPUT_DIR})")
    parser.add_argument(
        "--reference-date", type=iso_date_arg, default=None,
        help="date used for the 28-day LTFU rule and future-date check, YYYY-MM-DD (default: today). "
             "Set it explicitly to reproduce a run.",
    )
    return parser.parse_args(argv)


def configure_logging(output_dir: Path) -> None:
    """Send log messages to the console and (append mode) to output/pipeline_run.log."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(output_dir / RUN_LOG_FILE, mode="a", encoding="utf-8"),
        ],
        force=True,
    )


def issues_to_frames(issues: list[validators.Issue]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Convert Issue objects into (a) a detailed frame for summaries and (b) the
    errors_to_fix.csv frame with exactly the columns patient_id, error_type,
    error_description. The description starts with the source row number so a
    record with a blank patient_id can still be found.
    """
    detail = pd.DataFrame(
        [
            {
                "source_row": i.source_row,
                "patient_id": i.patient_id if i.patient_id else utils.BLANK_ID,
                "error_type": i.error_type,
                "description": i.description,
                "blocking": i.blocking,
            }
            for i in issues
        ],
        columns=["source_row", "patient_id", "error_type", "description", "blocking"],
    )
    detail["blocking"] = detail["blocking"].astype(bool)
    errors = pd.DataFrame(
        {
            "patient_id": detail["patient_id"],
            "error_type": detail["error_type"],
            "error_description": [
                f"[row {i.source_row}] {i.description} ({i.action})" for i in issues
            ],
        },
        columns=["patient_id", "error_type", "error_description"],
    )
    return detail, errors


def run_pipeline(input_path: Path, output_dir: Path, reference_date: date) -> dict[str, int]:
    """
    Execute the full pipeline and return the headline counts.
    Raises InputFileError (bad input) or IntegrityError (internal check failed).
    """
    audit = utils.AuditLog()
    run_started = datetime.now(timezone.utc)

    # ---- 1. read -----------------------------------------------------------
    logger.info("Reading input file: %s", input_path)
    raw_df = utils.read_input_csv(input_path, audit)
    audit.record(
        "RUN_START", utils.NA_ID,
        f"input={input_path.name}; sha256={utils.sha256_of_file(input_path)}; rows_read={len(raw_df)}; "
        f"reference_date={reference_date}; tool_version={utils.TOOL_VERSION}",
    )
    logger.info("Rows read: %d", len(raw_df))

    # ---- 2. normalise ------------------------------------------------------
    df = utils.normalise_dataframe(raw_df, audit)

    # ---- 3. validate -------------------------------------------------------
    result = validators.validate_dataframe(df, reference_date)
    for issue in result.issues:
        audit.record(issue.rule_label, issue.patient_id, f"row {issue.source_row}: {issue.description} -> {issue.action}")
    for decision in result.duplicate_decisions:
        rows = ", ".join(map(str, decision["rows"]))
        outcome = (
            f"row {decision['retained_row']} retained as first valid record"
            if decision["retained_row"] is not None
            else "no valid record - all excluded"
        )
        audit.record(validators.RULE_LABELS[validators.DUPLICATE_ID], decision["patient_id"],
                     f"duplicate group in rows {rows}: {outcome}")

    # ---- 4. select + verify clean records ---------------------------------
    accepted = set(result.accepted_rows)
    clean_with_row = df[df["source_row"].isin(accepted)].copy()
    validators.verify_clean_output(clean_with_row, reference_date)

    for rec in df[["source_row", "patient_id"]].itertuples(index=False):
        row = int(str(rec.source_row))
        pid_value = rec.patient_id
        pid = "" if pd.isna(pid_value) else str(pid_value)
        if row in accepted:
            audit.record("FINAL_DISPOSITION", pid, f"row {row}: ACCEPTED -> {CLEAN_FILE}")
        else:
            audit.record("FINAL_DISPOSITION", pid, f"row {row}: REJECTED -> {ERRORS_FILE}")

    # ---- 5. reconcile counts ----------------------------------------------
    rows_read = len(df)
    if len(result.accepted_rows) + len(result.rejected_rows) != rows_read or accepted & set(result.rejected_rows):
        raise utils.IntegrityError("accepted + rejected records do not reconcile with rows read")
    if len(clean_with_row) != len(result.accepted_rows):
        raise utils.IntegrityError("clean data row count does not match accepted record count")
    rejected_without_reason = set(result.rejected_rows) - {i.source_row for i in result.issues if i.blocking}
    if rejected_without_reason:
        raise utils.IntegrityError("a record was rejected without a recorded blocking reason")

    detail_df, errors_df = issues_to_frames(result.issues)
    warning_rows = {i.source_row for i in result.issues if not i.blocking} & accepted
    counts = {
        "rows_read": rows_read,
        "rows_accepted": len(result.accepted_rows),
        "rows_rejected": len(result.rejected_rows),
        "accepted_with_warnings": len(warning_rows),
        "total_issues": len(result.issues),
    }
    audit.record(
        "RUN_END", utils.NA_ID,
        f"rows_read={counts['rows_read']}; accepted={counts['rows_accepted']}; rejected={counts['rows_rejected']}; "
        f"accepted_with_warnings={counts['accepted_with_warnings']}; issues={counts['total_issues']}",
    )

    # ---- 6. write outputs --------------------------------------------------
    output_dir.mkdir(parents=True, exist_ok=True)
    utils.write_csv_atomic(clean_with_row[utils.EXPECTED_COLUMNS], output_dir / CLEAN_FILE)
    utils.write_csv_atomic(utils.sanitise_dataframe(errors_df), output_dir / ERRORS_FILE)
    meta = {
        "input_file": input_path.name,
        "input_sha256": utils.sha256_of_file(input_path),
        "reference_date": str(reference_date),
        "run_timestamp_utc": run_started.isoformat(timespec="seconds"),
        **counts,
    }
    utils.write_quality_report(
        output_dir / REPORT_FILE, meta, detail_df, errors_df, validators.ERROR_TYPES, validators.RULE_LABELS
    )
    utils.write_csv_atomic(utils.sanitise_dataframe(audit.to_dataframe()), output_dir / LOG_FILE)
    return counts


def main(argv: list[str] | None = None) -> int:
    """Run the pipeline, turning every failure into a clear message and exit code."""
    args = parse_args(argv)
    reference_date = args.reference_date or date.today()
    try:
        args.output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(f"ERROR: cannot create output folder '{args.output_dir}': {exc}", file=sys.stderr)
        return 1
    configure_logging(args.output_dir)

    try:
        counts = run_pipeline(args.input, args.output_dir, reference_date)
    except utils.InputFileError as exc:
        logger.error("INPUT PROBLEM: %s", exc)
        return 2
    except utils.IntegrityError as exc:
        logger.critical("INTEGRITY CHECK FAILED - no outputs were written and none may be used: %s", exc)
        return 3
    except PermissionError as exc:
        logger.error("Cannot write to '%s' (%s). Close any output file that is open in Excel and run again.", args.output_dir, exc)
        return 1
    except Exception:  # last-resort handler: log the full traceback to the run log
        logger.exception("Unexpected error - the run was aborted")
        return 1

    logger.info(
        "Done. Rows read: %d | accepted: %d | rejected: %d | accepted with warnings: %d | issues flagged: %d",
        counts["rows_read"], counts["rows_accepted"], counts["rows_rejected"],
        counts["accepted_with_warnings"], counts["total_issues"],
    )
    logger.info("Outputs written to: %s", args.output_dir)
    return 0


if __name__ == "__main__":
    sys.exit(main())
