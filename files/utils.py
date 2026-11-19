"""
utils.py
========
Shared constants, input/output helpers, audit logging and report writers for
the HAST Data Quality Guardian.

Design principles
-----------------
* All input is read as TEXT. Nothing is auto-converted, so facility codes keep
  their leading zeros and no value is silently turned into NaN, a number or a
  date.
* Every change made to a value is written to the audit log (cleaning_log.csv).
* Only the eight documented columns ever leave this program (POPIA data
  minimisation). Any extra column - especially one that looks like a personal
  identifier - is dropped at read time and is never written to any output.
* Output files are written to a temporary file first and then moved into
  place, so a crash can never leave a half-written output file behind.

This module has no dependency on validators.py (validators.py imports from
here, never the other way round).
"""

from __future__ import annotations

import csv
import hashlib
import io
import logging
import os
import re
from datetime import date, datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

TOOL_VERSION = "1.0.0"

# --------------------------------------------------------------------------
# Data specification (single source of truth for every module)
# --------------------------------------------------------------------------
EXPECTED_COLUMNS = [
    "patient_id",
    "facility_code",
    "dob",
    "gender",
    "art_start_date",
    "last_visit_date",
    "current_status",
    "regimen",
]
CRITICAL_FIELDS = ("patient_id", "dob", "gender", "art_start_date")
DATE_FIELDS = ("dob", "art_start_date", "last_visit_date")
ALLOWED_GENDERS = ("Male", "Female", "Other")
ALLOWED_STATUSES = ("Active", "LTFU", "Transferred Out", "Died")
LTFU_THRESHOLD_DAYS = 28  # Active + last visit MORE than this many days ago -> flag

BLANK_ID = "(blank)"  # shown in outputs when patient_id is missing
NA_ID = "N/A"  # shown in the audit log for run-level (non-patient) events

# Text values that are treated as "no value" (compared case-insensitively).
NULL_TOKENS = frozenset({"", "na", "n/a", "nan", "null", "none"})

# Strict patterns. [0-9] is used instead of \d so that non-ASCII digits are rejected.
FACILITY_CODE_RE = re.compile(r"[0-9]{6}")
ISO_DATE_RE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}")

# Case-insensitive look-ups used ONLY to repair letter case / spacing.
# Abbreviations such as "M" or "TO" are deliberately NOT mapped: they fail validation.
GENDER_CANONICAL = {g.casefold(): g for g in ALLOWED_GENDERS}
STATUS_CANONICAL = {s.casefold(): s for s in ALLOWED_STATUSES}

# Column-name fragments that suggest personal identifiers (used only to word the warning).
PII_HINTS = (
    "name", "surname", "phone", "cell", "mobile", "address", "email",
    "id_number", "idnumber", "passport", "contact", "kin",
)

_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


class InputFileError(Exception):
    """Raised when the input file is missing, unreadable or structurally invalid."""


class IntegrityError(Exception):
    """Raised when a post-condition on the outputs fails (outputs must not be used)."""


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def is_missing(value: object) -> bool:
    """Return True if ``value`` is None/NaN or a blank / null-token string."""
    if value is None:
        return True
    if isinstance(value, float) and np.isnan(value):
        return True
    return str(value).strip().casefold() in NULL_TOKENS


def parse_iso_date(value: object) -> date | None:
    """
    Strictly parse a YYYY-MM-DD date.

    Returns None for a blank value (the caller decides whether blank is allowed).
    Raises ValueError (with a human-readable reason) if the text is not a real
    YYYY-MM-DD date. No guessing: 03/04/2020 is never interpreted.
    """
    if is_missing(value):
        return None
    text = str(value).strip()
    if not ISO_DATE_RE.fullmatch(text):
        raise ValueError("is not in the required YYYY-MM-DD format")
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("is not a valid calendar date") from exc


def days_between(start: date, end: date) -> int:
    """Return the whole number of days from ``start`` to ``end`` (NumPy date arithmetic)."""
    delta = np.datetime64(end, "D") - np.datetime64(start, "D")
    return int(delta / np.timedelta64(1, "D"))


def shorten(value: object, limit: int = 40) -> str:
    """Return ``value`` as a single-line string of at most ``limit`` characters (for log text)."""
    text = " ".join(str(value).split())
    return text if len(text) <= limit else text[: limit - 3] + "..."


def quote(value: object, limit: int = 40) -> str:
    """
    Return ``value`` in single quotes EXACTLY as received (leading/trailing spaces stay
    visible), truncated to ``limit`` characters. Used when logging a value that is
    about to be changed, so the audit trail shows the original text.
    """
    text = str(value).replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
    if len(text) > limit:
        text = text[: limit - 3] + "..."
    return f"'{text}'"


def sanitise_cell(value: object) -> object:
    """
    Neutralise spreadsheet formula injection: a text cell that starts with
    = + - @ would be executed by Excel, so it is prefixed with an apostrophe.
    """
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value
    return value


def sanitise_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of ``df`` with every text cell passed through sanitise_cell."""
    return df.apply(lambda col: col.map(sanitise_cell))


def sha256_of_file(path: Path) -> str:
    """Return the SHA-256 hex digest of a file (recorded so the audit trail identifies the exact input)."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


# --------------------------------------------------------------------------
# Audit log
# --------------------------------------------------------------------------
class AuditLog:
    """
    In-memory audit trail. One entry is recorded for every validation failure,
    every value transformation and every record disposition; the trail is
    written to cleaning_log.csv at the end of a successful run.
    """

    COLUMNS = ["timestamp", "rule_applied", "patient_id", "action_taken"]

    def __init__(self) -> None:
        """Create an empty audit log."""
        self._entries: list[dict[str, str]] = []

    def record(self, rule_applied: str, patient_id: str, action_taken: str) -> None:
        """Append one entry stamped with the current UTC time (ISO 8601)."""
        self._entries.append(
            {
                "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "rule_applied": rule_applied,
                "patient_id": patient_id if patient_id else BLANK_ID,
                "action_taken": action_taken,
            }
        )

    def __len__(self) -> int:
        """Return the number of entries recorded so far."""
        return len(self._entries)

    def to_dataframe(self) -> pd.DataFrame:
        """Return the audit trail as a DataFrame with the four required columns."""
        return pd.DataFrame(self._entries, columns=self.COLUMNS)


# --------------------------------------------------------------------------
# Reading the input file
# --------------------------------------------------------------------------
def _detect_delimiter(header_line: str) -> str:
    """
    Guess the field delimiter from the header line. South African (en-ZA) Excel
    commonly exports with ';' instead of ','. Falls back to ','.
    """
    candidates = [",", ";", "\t", "|"]
    best = max(candidates, key=lambda d: header_line.count(d))
    return best if header_line.count(best) > 0 else ","


def read_input_csv(path: Path, audit: AuditLog) -> pd.DataFrame:
    """
    Read the messy clinic export and return a DataFrame of TEXT values with the
    columns ``source_row`` (line number in the file) + the eight expected columns.

    Fails loudly (InputFileError) when the file is missing, empty, unreadable,
    has missing/duplicated header names, or contains a row with MORE fields than
    the header (a shifted row could put values in the wrong column, so the run is
    stopped instead of guessing). Rows with FEWER fields are padded with blanks
    and logged. Completely blank lines are skipped and logged. Unexpected
    columns are dropped (POPIA data minimisation) and logged.
    """
    path = Path(path)
    if not path.exists():
        raise InputFileError(
            f"Input file not found: '{path}'. Check the file name and folder, "
            "or supply the file with --input <path>."
        )
    if not path.is_file():
        raise InputFileError(f"Input path is not a file: '{path}'.")
    try:
        raw = path.read_bytes()
    except PermissionError as exc:
        raise InputFileError(
            f"Permission denied reading '{path}'. Close the file if it is open in another program."
        ) from exc
    except OSError as exc:
        raise InputFileError(f"Could not read '{path}': {exc}") from exc
    if not raw.strip():
        raise InputFileError(f"Input file '{path}' is empty.")

    try:
        text = raw.decode("utf-8-sig")  # utf-8-sig also strips an Excel byte-order mark
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
        logger.warning("Input is not valid UTF-8; decoded as latin-1. Check special characters.")
        audit.record("INPUT_SCHEMA", NA_ID, "File is not valid UTF-8; decoded as latin-1 (fallback).")

    first_line = next((ln for ln in text.splitlines() if ln.strip()), "")
    delimiter = _detect_delimiter(first_line)
    if delimiter != ",":
        logger.warning("Detected delimiter %r instead of a comma.", delimiter)
        audit.record("INPUT_SCHEMA", NA_ID, f"Detected non-comma field delimiter {delimiter!r}.")

    raw_header: list[str] | None = None
    header: list[str] | None = None
    rows: list[list[str]] = []
    line_numbers: list[int] = []
    blank_lines: list[int] = []
    short_lines: list[int] = []

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    previous_line = 0
    try:
        for row in reader:
            start_line = previous_line + 1
            previous_line = reader.line_num
            if not row or all(not cell.strip() for cell in row):
                blank_lines.append(start_line)
                continue
            if header is None:
                raw_header = [cell.strip() for cell in row]
                header = [cell.lower() for cell in raw_header]
                continue
            if len(row) > len(header):
                raise InputFileError(
                    f"Line {start_line} has {len(row)} fields but the header has {len(header)}. "
                    "A shifted row could place values in the wrong column, so the run was stopped. "
                    "Fix that row and run again."
                )
            if len(row) < len(header):
                short_lines.append(start_line)
                row = row + [""] * (len(header) - len(row))
            rows.append(row)
            line_numbers.append(start_line)
    except csv.Error as exc:
        raise InputFileError(f"Malformed CSV near line {reader.line_num}: {exc}") from exc

    if header is None or raw_header is None:
        raise InputFileError(f"Input file '{path}' has no header row.")

    # ---- header checks -----------------------------------------------------
    duplicated = sorted({c for c in header if header.count(c) > 1})
    if duplicated:
        raise InputFileError(f"Duplicate column name(s) in header: {', '.join(duplicated)}.")
    missing_cols = [c for c in EXPECTED_COLUMNS if c not in header]
    if missing_cols:
        raise InputFileError(
            f"Required column(s) missing from header: {', '.join(missing_cols)}. "
            f"Expected columns: {', '.join(EXPECTED_COLUMNS)}."
        )
    if raw_header != header:
        audit.record("INPUT_SCHEMA", NA_ID, "Header names trimmed / converted to lower case.")

    extra_cols = [c for c in header if c not in EXPECTED_COLUMNS]
    if extra_cols:
        pii_like = [c for c in extra_cols if any(hint in c for hint in PII_HINTS)]
        message = f"Dropped {len(extra_cols)} unexpected column(s): {', '.join(shorten(c, 30) for c in extra_cols)}."
        if pii_like:
            message += " Possible personal identifiers detected - they are NOT carried into any output (POPIA)."
            logger.warning("Possible personal-identifier column(s) in input were excluded: %s", pii_like)
        audit.record("PRIVACY_POPIA", NA_ID, message)
    if blank_lines:
        audit.record(
            "INPUT_SCHEMA", NA_ID,
            f"Skipped {len(blank_lines)} fully blank line(s) (first lines: {blank_lines[:10]}).",
        )
    if short_lines:
        audit.record(
            "INPUT_SCHEMA", NA_ID,
            f"{len(short_lines)} line(s) had fewer fields than the header and were padded with blanks "
            f"(first lines: {short_lines[:10]}).",
        )
    if not rows:
        raise InputFileError("The input file has a header but no data rows.")

    df = pd.DataFrame(rows, columns=header)[EXPECTED_COLUMNS].copy()
    df.insert(0, "source_row", line_numbers)
    return df


# --------------------------------------------------------------------------
# Normalisation (the ONLY place values are changed - every change is logged)
# --------------------------------------------------------------------------
def normalise_dataframe(df: pd.DataFrame, audit: AuditLog) -> pd.DataFrame:
    """
    Apply the only value changes this tool is allowed to make, logging each one:

    1. Trim leading/trailing spaces and collapse repeated inner spaces.
    2. Convert null tokens (NA, N/A, null, none, nan) to a blank value.
    3. Repair the letter case of gender and current_status
       (e.g. 'female' -> 'Female', 'transferred out' -> 'Transferred Out').

    Nothing else is altered. Abbreviations, wrong date formats and unknown
    categories are NOT guessed - they are left as-is and fail validation.
    """
    records = df.to_dict("records")
    for rec in records:
        row = int(rec["source_row"])
        pid = BLANK_ID
        for column in EXPECTED_COLUMNS:  # patient_id is first, so pid is known for later columns
            original = str(rec[column])
            value = " ".join(original.split())
            if value != "" and value.casefold() in NULL_TOKENS:
                audit.record(
                    "NORMALISE_NULL_TOKEN", pid if column != "patient_id" else BLANK_ID,
                    f"row {row}: {column} {quote(original)} treated as blank",
                )
                value = ""
            elif value != original:
                label = value if column == "patient_id" else pid
                audit.record(
                    "NORMALISE_WHITESPACE", label or BLANK_ID,
                    f"row {row}: {column} {quote(original)} -> {quote(value)} (spaces trimmed)",
                )
            lookup = GENDER_CANONICAL if column == "gender" else STATUS_CANONICAL if column == "current_status" else None
            if lookup is not None and value.casefold() in lookup and value != lookup[value.casefold()]:
                canonical = lookup[value.casefold()]
                audit.record(
                    "NORMALISE_CASE", pid if column != "patient_id" else BLANK_ID,
                    f"row {row}: {column} '{shorten(value)}' -> '{canonical}' (letter case corrected)",
                )
                value = canonical
            rec[column] = value
            if column == "patient_id":
                pid = value if value else BLANK_ID
    return pd.DataFrame(records, columns=["source_row"] + EXPECTED_COLUMNS)


# --------------------------------------------------------------------------
# Writing outputs
# --------------------------------------------------------------------------
def write_csv_atomic(df: pd.DataFrame, path: Path) -> None:
    """Write ``df`` to ``path`` as UTF-8 CSV via a temporary file, then move it into place."""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    try:
        df.to_csv(tmp, index=False, encoding="utf-8")
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


def summarise_issues(issues_df: pd.DataFrame, error_types: tuple[str, ...], rule_labels: dict[str, str]) -> pd.DataFrame:
    """
    Build the per-error-type summary table used in the Excel report.

    ``issues_df`` needs the columns: source_row, error_type, blocking.
    Every error type is listed, including those with zero issues, so a reviewer
    can see that each check was executed.
    """
    lines = []
    for error_type in error_types:
        subset = issues_df[issues_df["error_type"] == error_type]
        lines.append(
            {
                "Error type": error_type,
                "Rule": rule_labels[error_type],
                "Blocking (record excluded)": int(subset["blocking"].sum()),
                "Warnings (record retained)": int((~subset["blocking"]).sum()),
                "Total issues": int(len(subset)),
                "Records affected": int(subset["source_row"].nunique()),
            }
        )
    return pd.DataFrame(lines)


def _render_pie_png(counts: dict[str, int], palette: list[str]) -> io.BytesIO:
    """Render a pie chart of issue counts with Matplotlib and return it as an in-memory PNG."""
    import matplotlib

    matplotlib.use("Agg")  # headless backend: no display needed
    import matplotlib.pyplot as plt

    labels = [k for k, v in counts.items() if v > 0]
    sizes = [counts[k] for k in labels]
    total = sum(sizes)
    colours = [palette[i % len(palette)] for i, k in enumerate(counts) if counts[k] > 0]

    fig, ax = plt.subplots(figsize=(6.6, 4.8), dpi=110)
    wedges, _texts, _autotexts = ax.pie(
        sizes,
        colors=colours,
        startangle=90,
        counterclock=False,
        autopct=lambda pct: f"{int(round(pct * total / 100))}",
        pctdistance=0.72,
        textprops={"color": "white", "fontsize": 11, "fontweight": "bold"},
        wedgeprops={"edgecolor": "white", "linewidth": 1.5},
    )
    ax.legend(
        wedges, [f"{label} ({size})" for label, size in zip(labels, sizes)],
        title="Error type (issues)", loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False,
    )
    ax.set_title("Data quality issues by type", fontsize=13, fontweight="bold")
    ax.axis("equal")
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def write_quality_report(
    path: Path,
    meta: dict[str, object],
    issues_df: pd.DataFrame,
    errors_df: pd.DataFrame,
    error_types: tuple[str, ...],
    rule_labels: dict[str, str],
) -> None:
    """
    Write data_quality_report.xlsx.

    Sheet 'Summary'      - run details, reconciliation counts, issue counts per error
                           type and a pie chart of issues by type.
    Sheet 'Error_Detail' - the same rows as errors_to_fix.csv.

    Values are written as static numbers because this is a point-in-time audit
    report of one run (it is not a live model). All text cells are sanitised
    against spreadsheet formula injection.
    """
    font = "Arial"
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    head_font = Font(name=font, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E78")

    ws["A1"] = "HAST Data Quality Report"
    ws["A1"].font = Font(name=font, bold=True, size=14)
    ws["A2"] = f"Generated by HAST Data Quality Guardian v{TOOL_VERSION}"
    ws["A2"].font = Font(name=font, italic=True, size=9)

    meta_rows = [
        ("Input file", meta["input_file"]),
        ("Input SHA-256", meta["input_sha256"]),
        ("Reference date (for LTFU rule)", meta["reference_date"]),
        ("Run timestamp (UTC)", meta["run_timestamp_utc"]),
        ("Rows read", meta["rows_read"]),
        ("Rows accepted -> hast_data_clean.csv", meta["rows_accepted"]),
        ("Rows rejected -> errors_to_fix.csv", meta["rows_rejected"]),
        ("Accepted rows carrying a warning", meta["accepted_with_warnings"]),
        ("Total issues flagged", meta["total_issues"]),
    ]
    row = 4
    for label, value in meta_rows:
        ws.cell(row=row, column=1, value=label).font = Font(name=font, bold=True)
        ws.cell(row=row, column=2, value=sanitise_cell(value)).font = Font(name=font)
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="left")
        row += 1

    row += 1
    summary_df = summarise_issues(issues_df, error_types, rule_labels)
    header_row = row
    for col_idx, title in enumerate(summary_df.columns, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.font, cell.fill = head_font, head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r_idx, record in enumerate(summary_df.itertuples(index=False), start=header_row + 1):
        for c_idx, value in enumerate(record, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value).font = Font(name=font)
    total_row = header_row + len(summary_df) + 1
    ws.cell(row=total_row, column=1, value="Total").font = Font(name=font, bold=True)
    for c_idx in (3, 4, 5):
        ws.cell(row=total_row, column=c_idx, value=int(summary_df.iloc[:, c_idx - 1].sum())).font = Font(name=font, bold=True)
    note_row = total_row + 2
    notes = [
        "Notes:",
        "- A record can carry several issues, so 'Total issues' can exceed 'Records affected'.",
        "- Blocking issues exclude the record from hast_data_clean.csv; warnings (e.g. Possible LTFU) do not.",
        "- Source row numbers quoted in errors_to_fix.csv refer to line numbers in the input file (header = line 1).",
    ]
    for offset, line in enumerate(notes):
        ws.cell(row=note_row + offset, column=1, value=line).font = Font(name=font, italic=(offset > 0), bold=(offset == 0), size=9)

    for letter, width in zip("ABCDEF", (40, 44, 26, 26, 14, 16)):
        ws.column_dimensions[letter].width = width
    ws.row_dimensions[header_row].height = 32
    # Print setup: landscape, one page wide, so the report prints cleanly for review.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1  # Summary is short: keep it on a single page
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- pie chart (Matplotlib image embedded in the sheet) ---------------
    counts = {error_type: int(n) for error_type, n in zip(summary_df["Error type"], summary_df["Total issues"])}
    chart_anchor = f"A{note_row + len(notes) + 1}"
    if sum(counts.values()) == 0:
        ws[chart_anchor] = "No data quality issues were flagged in this run."
        ws[chart_anchor].font = Font(name=font, bold=True)
    else:
        try:
            palette = ["#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7", "#56B4E9"]
            image = XLImage(_render_pie_png(counts, palette))
            image.width, image.height = int(image.width * 0.85), int(image.height * 0.85)
            ws.add_image(image, chart_anchor)
        except Exception as exc:  # chart is a convenience: never lose the numbers if it fails
            logger.warning("Pie chart could not be generated: %s", exc)
            ws[chart_anchor] = f"Pie chart could not be generated ({type(exc).__name__}). Counts above are complete."

    # ---- detail sheet ------------------------------------------------------
    detail = wb.create_sheet("Error_Detail")
    safe_errors = sanitise_dataframe(errors_df)
    for col_idx, title in enumerate(safe_errors.columns, start=1):
        cell = detail.cell(row=1, column=col_idx, value=title)
        cell.font, cell.fill = head_font, head_fill
    for r_idx, record in enumerate(safe_errors.itertuples(index=False), start=2):
        for c_idx, value in enumerate(record, start=1):
            cell = detail.cell(row=r_idx, column=c_idx, value=value)
            cell.font = Font(name=font)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for c_idx, width in enumerate((16, 22, 110), start=1):
        detail.column_dimensions[get_column_letter(c_idx)].width = width
    detail.freeze_panes = "A2"
    detail.page_setup.orientation = "landscape"
    detail.page_setup.fitToWidth = 1
    detail.page_setup.fitToHeight = 0
    detail.sheet_properties.pageSetUpPr.fitToPage = True
    detail.print_title_rows = "1:1"

    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
