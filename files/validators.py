"""
validators.py
=============
Validation rules for the HAST Data Quality Guardian.

The five rules requested for DHIS readiness are implemented here, plus one
supporting format check that the rules depend on (a date can only be compared
once it is known to be a real YYYY-MM-DD date, and facility_code must be six
digits).

Rule map
--------
RULE_1_DUPLICATE_CHECK    patient_id must be unique (first VALID record is kept)
RULE_2_MISSING_CRITICAL   patient_id, dob, gender, art_start_date must not be blank
RULE_3_DATE_LOGIC         dob < art_start_date < last_visit_date (and no future dates)
RULE_4_STATUS_LOGIC       Active + last visit > 28 days ago -> "Possible LTFU" (warning)
RULE_5_VALID_VALUES       gender in Male/Female/Other; current_status in the 4 allowed
SUPPORT_FORMAT_CHECK      facility_code = 6 digits; dates = real YYYY-MM-DD dates

Severity
--------
* BLOCKING issue : the record is excluded from hast_data_clean.csv.
* WARNING        : the record stays in hast_data_clean.csv unchanged but is listed
                   in errors_to_fix.csv for clinical review. The tool NEVER changes a
                   clinical value (for example it never rewrites Active to LTFU).
                   Warnings are reported only for records that are retained; a record
                   that is excluded is described by its blocking issues alone.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from utils import (
    ALLOWED_GENDERS,
    ALLOWED_STATUSES,
    CRITICAL_FIELDS,
    DATE_FIELDS,
    EXPECTED_COLUMNS,
    FACILITY_CODE_RE,
    LTFU_THRESHOLD_DAYS,
    IntegrityError,
    days_between,
    is_missing,
    parse_iso_date,
    shorten,
)

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------
# Error types and rule labels
# --------------------------------------------------------------------------
DUPLICATE_ID = "DUPLICATE_ID"
MISSING_CRITICAL = "MISSING_CRITICAL"
DATE_LOGIC = "DATE_LOGIC"
POSSIBLE_LTFU = "POSSIBLE_LTFU"
INVALID_VALUE = "INVALID_VALUE"
FORMAT_ERROR = "FORMAT_ERROR"

ERROR_TYPES = (DUPLICATE_ID, MISSING_CRITICAL, DATE_LOGIC, POSSIBLE_LTFU, INVALID_VALUE, FORMAT_ERROR)

RULE_LABELS = {
    DUPLICATE_ID: "RULE_1_DUPLICATE_CHECK",
    MISSING_CRITICAL: "RULE_2_MISSING_CRITICAL",
    DATE_LOGIC: "RULE_3_DATE_LOGIC",
    POSSIBLE_LTFU: "RULE_4_STATUS_LOGIC",
    INVALID_VALUE: "RULE_5_VALID_VALUES",
    FORMAT_ERROR: "SUPPORT_FORMAT_CHECK",
}

ACTION_EXCLUDED = "record EXCLUDED from clean file"
ACTION_RETAINED = "warning only - record RETAINED in clean file"


@dataclass(frozen=True)
class Issue:
    """One validation finding for one input row."""

    source_row: int
    patient_id: str
    error_type: str
    description: str
    blocking: bool = True

    @property
    def rule_label(self) -> str:
        """Return the audit-log rule label for this issue's error type."""
        return RULE_LABELS[self.error_type]

    @property
    def action(self) -> str:
        """Return the action taken as a result of this issue (exclude or retain)."""
        return ACTION_EXCLUDED if self.blocking else ACTION_RETAINED


@dataclass
class ValidationResult:
    """Outcome of validating a whole file."""

    issues: list[Issue] = field(default_factory=list)
    accepted_rows: list[int] = field(default_factory=list)  # source_row of every record kept
    rejected_rows: list[int] = field(default_factory=list)  # source_row of every record excluded
    duplicate_decisions: list[dict] = field(default_factory=list)  # which duplicate was retained


def _issue(rec: dict, error_type: str, description: str, blocking: bool = True) -> Issue:
    """Build an Issue for the record ``rec`` (keeps row number and patient_id in one place)."""
    return Issue(
        source_row=int(rec["source_row"]),
        patient_id=rec["patient_id"],
        error_type=error_type,
        description=description,
        blocking=blocking,
    )


def _out_of_order(earlier: date, later: date) -> bool:
    """
    Return True when ``earlier`` is NOT strictly before ``later``.
    The specification uses strict '<', so two equal dates are a failure.
    (To allow same-day dates, change this to ``earlier > later``.)
    """
    return earlier >= later


# --------------------------------------------------------------------------
# Rule 2 - missing critical fields
# --------------------------------------------------------------------------
def check_missing_critical(rec: dict) -> list[Issue]:
    """Rule 2: patient_id, dob, gender and art_start_date must not be blank or null."""
    return [
        _issue(rec, MISSING_CRITICAL, f"Critical field '{name}' is blank or null.")
        for name in CRITICAL_FIELDS
        if is_missing(rec[name])
    ]


# --------------------------------------------------------------------------
# Rule 5 - valid values
# --------------------------------------------------------------------------
def check_valid_values(rec: dict) -> list[Issue]:
    """
    Rule 5: gender must be Male/Female/Other; current_status must be one of
    Active/LTFU/Transferred Out/Died. A blank gender is reported by Rule 2 only
    (no double reporting); a blank current_status is reported here.
    """
    issues: list[Issue] = []
    gender = rec["gender"]
    if not is_missing(gender) and gender not in ALLOWED_GENDERS:
        issues.append(
            _issue(rec, INVALID_VALUE, f"gender '{shorten(gender)}' is not one of: {', '.join(ALLOWED_GENDERS)}.")
        )
    status = rec["current_status"]
    if is_missing(status):
        issues.append(
            _issue(rec, INVALID_VALUE, f"current_status is blank; it must be one of: {', '.join(ALLOWED_STATUSES)}.")
        )
    elif status not in ALLOWED_STATUSES:
        issues.append(
            _issue(
                rec, INVALID_VALUE,
                f"current_status '{shorten(status)}' is not one of: {', '.join(ALLOWED_STATUSES)}.",
            )
        )
    return issues


# --------------------------------------------------------------------------
# Supporting format check
# --------------------------------------------------------------------------
def check_formats(rec: dict) -> tuple[list[Issue], dict[str, date | None]]:
    """
    Supporting check: facility_code must be exactly six digits and each date
    must be a real YYYY-MM-DD date. Returns the issues plus the parsed dates
    (None for blank or unparseable) for use by the date and status rules.
    Blank dates are not reported here: Rule 2 covers the critical ones and
    Rule 3 notes a blank last_visit_date.
    """
    issues: list[Issue] = []
    facility = rec["facility_code"]
    if is_missing(facility):
        issues.append(_issue(rec, FORMAT_ERROR, "facility_code is blank; it must be exactly 6 digits."))
    elif not FACILITY_CODE_RE.fullmatch(facility):
        issues.append(
            _issue(rec, FORMAT_ERROR, f"facility_code '{shorten(facility)}' must be exactly 6 digits.")
        )

    parsed: dict[str, date | None] = {}
    for name in DATE_FIELDS:
        try:
            parsed[name] = parse_iso_date(rec[name])
        except ValueError as exc:
            parsed[name] = None
            issues.append(
                _issue(rec, FORMAT_ERROR, f"{name} '{shorten(rec[name])}' {exc}.")
            )
    return issues, parsed


# --------------------------------------------------------------------------
# Rule 3 - date logic
# --------------------------------------------------------------------------
def check_date_logic(rec: dict, parsed: dict[str, date | None], reference_date: date) -> list[Issue]:
    """
    Rule 3: dob < art_start_date < last_visit_date, and no date after the
    reference date. Each comparison is only made when both dates are present and
    valid, so one bad date does not hide or invent other findings. A blank
    last_visit_date is not a critical field, so it only raises a warning that
    part of this rule (and Rule 4) could not be evaluated.
    """
    issues: list[Issue] = []
    dob, art, last = parsed["dob"], parsed["art_start_date"], parsed["last_visit_date"]

    for name, value in parsed.items():
        if value is not None and value > reference_date:
            issues.append(
                _issue(rec, DATE_LOGIC, f"{name} ({value}) is after the reference date ({reference_date}).")
            )
    if dob and art and _out_of_order(dob, art):
        issues.append(
            _issue(rec, DATE_LOGIC, f"dob ({dob}) is not before art_start_date ({art}); required order is dob < art_start_date < last_visit_date.")
        )
    if art and last and _out_of_order(art, last):
        issues.append(
            _issue(rec, DATE_LOGIC, f"art_start_date ({art}) is not before last_visit_date ({last}); required order is dob < art_start_date < last_visit_date.")
        )
    if is_missing(rec["last_visit_date"]):
        issues.append(
            _issue(
                rec, DATE_LOGIC,
                "last_visit_date is blank; the art_start_date < last_visit_date check and the LTFU check could not be evaluated.",
                blocking=False,
            )
        )
    return issues


# --------------------------------------------------------------------------
# Rule 4 - status logic
# --------------------------------------------------------------------------
def check_status_logic(rec: dict, parsed: dict[str, date | None], reference_date: date) -> list[Issue]:
    """
    Rule 4: current_status 'Active' but last_visit_date MORE than 28 days before
    the reference date -> 'Possible LTFU'. This is a WARNING: the status is never
    changed automatically because only the clinic can confirm the patient's outcome.
    """
    last = parsed["last_visit_date"]
    if rec["current_status"] != "Active" or last is None or last > reference_date:
        return []
    days = days_between(last, reference_date)
    if days > LTFU_THRESHOLD_DAYS:
        return [
            _issue(
                rec, POSSIBLE_LTFU,
                f"Possible LTFU: current_status is 'Active' but last_visit_date ({last}) is {days} days before "
                f"the reference date ({reference_date}), more than {LTFU_THRESHOLD_DAYS} days. "
                "Confirm the patient's status with the clinic.",
                blocking=False,
            )
        ]
    return []


# --------------------------------------------------------------------------
# Record-level and file-level validation
# --------------------------------------------------------------------------
def validate_record(rec: dict, reference_date: date) -> list[Issue]:
    """
    Run every record-level rule (all rules except the duplicate check, which
    needs the whole file) on one record and return all findings.
    """
    issues = check_missing_critical(rec)
    issues += check_valid_values(rec)
    format_issues, parsed = check_formats(rec)
    issues += format_issues
    issues += check_date_logic(rec, parsed, reference_date)
    issues += check_status_logic(rec, parsed, reference_date)
    return issues


def _is_blocked(issues: list[Issue]) -> bool:
    """Return True if any issue in the list is blocking."""
    return any(issue.blocking for issue in issues)


def validate_dataframe(df: pd.DataFrame, reference_date: date) -> ValidationResult:
    """
    Validate every record and apply Rule 1 (duplicates).

    Duplicate handling: for each patient_id that appears more than once, the
    FIRST record that passes every blocking check is retained. All other records
    with that patient_id receive a DUPLICATE_ID issue and are excluded. If no
    record of the group is valid, all are excluded for manual resolution.
    Records with a blank patient_id are never grouped (Rule 2 already rejects them).
    """
    records = df.to_dict("records")
    by_row: dict[int, list[Issue]] = {int(r["source_row"]): validate_record(r, reference_date) for r in records}

    groups: dict[str, list[dict]] = {}
    for rec in records:
        if not is_missing(rec["patient_id"]):
            groups.setdefault(rec["patient_id"], []).append(rec)

    decisions: list[dict] = []
    for pid, members in groups.items():
        if len(members) < 2:
            continue
        rows_text = ", ".join(str(int(m["source_row"])) for m in members)
        keeper = next((m for m in members if not _is_blocked(by_row[int(m["source_row"])])), None)
        decisions.append(
            {
                "patient_id": pid,
                "rows": [int(m["source_row"]) for m in members],
                "retained_row": int(keeper["source_row"]) if keeper else None,
            }
        )
        for member in members:
            if member is keeper:
                continue
            if keeper is None:
                text = (
                    f"Duplicate patient_id (appears {len(members)} times, rows {rows_text}) and none of the "
                    "records is valid; all were excluded - resolve manually."
                )
            else:
                differing = [c for c in EXPECTED_COLUMNS[1:] if member[c] != keeper[c]]
                detail = (
                    "is identical to the retained record"
                    if not differing
                    else "differs from the retained record in: " + ", ".join(differing)
                )
                text = (
                    f"Duplicate patient_id (appears {len(members)} times, rows {rows_text}); row "
                    f"{int(keeper['source_row'])} was retained as the first valid record; this row {detail}."
                )
            by_row[int(member["source_row"])].append(_issue(member, DUPLICATE_ID, text))

    # Warnings are only reported for records that are retained. An excluded record is
    # fully described by its blocking issues, and a warning computed from data that
    # already failed validation (e.g. a LTFU day-count from an impossible date) would
    # be misleading. Warnings reappear automatically once the record is corrected.
    for row, row_issues in by_row.items():
        if _is_blocked(row_issues):
            by_row[row] = [i for i in row_issues if i.blocking]

    result = ValidationResult(duplicate_decisions=decisions)
    for rec in records:
        row = int(rec["source_row"])
        (result.rejected_rows if _is_blocked(by_row[row]) else result.accepted_rows).append(row)
    all_issues = [issue for issues in by_row.values() for issue in issues]
    result.issues = sorted(all_issues, key=lambda i: (i.source_row, ERROR_TYPES.index(i.error_type)))
    return result


def verify_clean_output(clean_df: pd.DataFrame, reference_date: date) -> None:
    """
    Independent self-check run BEFORE any output file is written. The records
    selected for hast_data_clean.csv are validated again from scratch and must
    show: no blank patient_id, no duplicate patient_id and zero blocking issues.
    Raises IntegrityError otherwise, so a faulty file is never produced silently.
    """
    if clean_df.empty:
        return
    problems: list[str] = []
    if clean_df["patient_id"].map(is_missing).any():
        problems.append("a blank patient_id reached the clean data")
    if clean_df["patient_id"].duplicated().any():
        problems.append("a duplicate patient_id reached the clean data")
    recheck = validate_dataframe(clean_df, reference_date)
    blocking = [i for i in recheck.issues if i.blocking]
    if blocking:
        problems.append(f"{len(blocking)} blocking issue(s) found when re-validating the clean data")
    if problems:
        raise IntegrityError("; ".join(problems))
