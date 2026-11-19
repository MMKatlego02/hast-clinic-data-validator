# HAST Data Quality Guardian

A defensive Python data-quality pipeline that takes a messy HIV/TB clinic export (CSV) and produces a DHIS-ready clean file, an error list for the clinic to fix, a summary report and a full audit trail.

## Problem Statement

Clinic exports routinely contain duplicate patients, blank critical fields, impossible dates, inconsistent spelling and patients recorded as *Active* who have not visited for months. When such files are loaded into DHIS they are rejected, or worse, accepted and used for reporting and care decisions.

This tool finds those problems **before** upload, separates the records that are safe to load from those that need correction, and records every decision so that the process can be audited. It prevents DHIS rejections and improves the quality of patient care data.

**Core safety principle:** the tool never invents or silently changes clinical data. The only changes it makes are trimming spaces, treating `NA`/`null` as blank, and repairing letter case in `gender` and `current_status` (e.g. `female` → `Female`). Every such change is written to the audit log. Anything else that is wrong is flagged, not guessed.

## Quick Start (3 steps)

1. **Install Python 3.10 or newer** and open a terminal in this folder.
2. **Create an environment and install the packages**

   Windows:
   ```
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
   macOS / Linux:
   ```
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. **Run it** on the sample file (the fixed reference date reproduces the numbers in this README):
   ```
   python main.py --reference-date 2026-11-14
   ```
   Results appear in the `output/` folder. To use your own export: `python main.py --input path\to\your_file.csv`.

## Folder Structure

```
hast_data_quality_guardian/
├── main.py                  # entry point: orchestrates the pipeline
├── validators.py            # the validation rules
├── utils.py                 # constants, file reading, audit log, report writers
├── requirements.txt
├── README.md
├── data/
│   └── hast_data_messy.csv  # synthetic sample, 20 rows, intentional errors
└── output/                  # created automatically
    ├── hast_data_clean.csv
    ├── errors_to_fix.csv
    ├── data_quality_report.xlsx
    ├── cleaning_log.csv
    └── pipeline_run.log
```

## Input File

`hast_data_messy.csv` must contain these columns (names are matched ignoring case and surrounding spaces; extra columns are ignored):

| Column | Meaning |
|---|---|
| `patient_id` | Unique identifier. Must not be blank or duplicated |
| `facility_code` | 6-digit code |
| `dob` | Date of birth, `YYYY-MM-DD` |
| `gender` | `Male`, `Female` or `Other` |
| `art_start_date` | Date ART started, `YYYY-MM-DD` |
| `last_visit_date` | Date of last clinic visit, `YYYY-MM-DD` |
| `current_status` | `Active`, `LTFU`, `Transferred Out` or `Died` |
| `regimen` | e.g. `TDF/3TC/DTG` (passed through, not validated) |

Comma, semicolon, tab and pipe delimiters are detected automatically (South African Excel often exports with `;`).

## Validation Rules

| # | Rule | Error type | Effect on the record |
|---|---|---|---|
| 1 | **Duplicate check** – `patient_id` must be unique. The first *valid* record is kept; the others are flagged | `DUPLICATE_ID` | Duplicates excluded |
| 2 | **Missing critical** – `patient_id`, `dob`, `gender`, `art_start_date` must not be blank/null | `MISSING_CRITICAL` | Excluded |
| 3 | **Date logic** – `dob < art_start_date < last_visit_date` (strict), and no date after the reference date | `DATE_LOGIC` | Excluded |
| 4 | **Status logic** – `current_status` is `Active` and `last_visit_date` is more than 28 days before the reference date → "Possible LTFU" | `POSSIBLE_LTFU` | **Warning only: record kept, status not changed** |
| 5 | **Valid values** – `gender` ∈ {Male, Female, Other}; `current_status` ∈ {Active, LTFU, Transferred Out, Died} | `INVALID_VALUE` | Excluded |
| – | **Supporting format check** – `facility_code` is exactly 6 digits; dates are real `YYYY-MM-DD` dates (needed before rule 3 can compare them) | `FORMAT_ERROR` | Excluded |

### Decisions the tool makes (please review)

* **Duplicates.** Records are matched on `patient_id` after trimming spaces, case-sensitive. The first record that passes every blocking check is kept; every other record with that ID is excluded and described (identical, or which fields differ). If no record of the group is valid, all are excluded for manual resolution.
* **Possible LTFU is a warning, not an error.** Only the clinic can confirm a patient's outcome, so the status is never rewritten. Warnings are reported only for records that are kept.
* **Strict date order.** Equal dates fail (e.g. ART started on the day of the last visit). To allow same-day dates, change `_out_of_order` in `validators.py`.
* **`last_visit_date` is not a critical field** in the specification, so a blank one does not reject the record; it raises a warning that the date-order and LTFU checks could not be run.
* **No guessing.** `M`, `TO`, `Defaulted`, `14/03/1987` and `1985-3-4` are not interpreted; they fail validation.
* **Reference date.** The 28-day rule and future-date check use *today* unless `--reference-date` is given. The sample file's "Active" patients will all become Possible LTFU as time passes, so use `--reference-date 2026-11-14` to reproduce the results below.

## Output Files

| File | Content |
|---|---|
| `hast_data_clean.csv` | One row per `patient_id`, the eight original columns only, ready for DHIS. Records with a warning (Possible LTFU) are included unchanged |
| `errors_to_fix.csv` | Columns `patient_id`, `error_type`, `error_description`. One row per problem. The description starts with `[row N]` (line number in the input file, header = line 1) and states whether the record was excluded or kept |
| `data_quality_report.xlsx` | **Summary** sheet: run details, input file SHA-256, row counts, error counts per type, and a pie chart. **Error_Detail** sheet: same rows as `errors_to_fix.csv` |
| `cleaning_log.csv` | Audit trail with columns `timestamp` (UTC), `rule_applied`, `patient_id`, `action_taken`. Contains every normalisation, every failed rule, every duplicate decision and a final `ACCEPTED`/`REJECTED` line for each input row |
| `pipeline_run.log` | Technical run log (appended on every run) |

Expected result for the sample file with `--reference-date 2026-11-14`: 20 rows read → 11 accepted, 9 rejected, 12 issues flagged (1 warning).

> **Tip:** open CSV files in Excel via *Data → From Text/CSV* and set `facility_code` to Text; double-clicking a CSV makes Excel drop the leading zero of codes such as `012345`.

## Safeguards Built In

* **POPIA:** no patient names are used. Any column outside the eight expected ones (for example a name or phone column) is dropped when the file is read, is never written to any output, and the drop is logged.
* **Nothing is written until all checks pass.** The records chosen for the clean file are re-validated from scratch, the counts must reconcile (rows read = accepted + rejected) and every rejection must have a recorded reason. If any check fails the run stops with exit code 3 and no outputs are written.
* **Stops instead of guessing.** A missing/empty file, missing or duplicated column, or a row with more fields than the header ends the run with a clear message (exit code 2). A shifted row could put values in the wrong column, so it is never auto-repaired.
* **Text-only reading.** Values are never auto-converted, so `012345` stays `012345`.
* **Traceability.** The audit log records the SHA-256 of the input file, the reference date and the tool version.
* **Spreadsheet formula injection:** error and log outputs prefix cells starting with `=`, `+`, `-` or `@` with an apostrophe.
* Output files are written to a temporary file and then moved into place.

### Exit codes

`0` success · `1` unexpected error (see `pipeline_run.log`) · `2` problem with the input file · `3` internal integrity check failed – do not use any outputs.

## Not Validated (out of scope)

`regimen` content, whether `facility_code` exists in the national master facility list, `patient_id` format, plausibility of age, and clinical consistency (e.g. regimen versus age). Add rules to `validators.py` as DOH requirements are confirmed.

## Disclaimer

This uses synthetic data. For training purposes only.
