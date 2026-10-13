# OfficeAutomationPro — Pay Register Splitter Plan

Status: DRAFT — items marked **[TBC]** depend on open questions.

## Goal
Split a master Pay Register workbook into separate output files by a grouping key
**[TBC: department / location / cost centre / employee]**, preserving values and formatting,
with a verifiable audit trail. Runs offline on Windows.

## Architecture
```
app/
  main.py              entry point (CLI first; GUI [TBC])
  excel_inspector.py   existing read-only COM inspector (diagnostic tool)
modules/
  config.py            settings: header row, group column, output naming, paths
  reader.py            load register (read-only); detect header, data rows, totals rows
  validator.py         schema checks: required columns, blanks, duplicates, numeric types
  splitter.py          group rows -> per-group workbooks (format-preserving)
  exporter.py          optional PDF export [TBC]
  reconciler.py        control totals: source vs sum of outputs
  audit_log.py         run log + manifest (file, rows, totals, SHA-256)
tests/                 pytest; synthetic fixtures only
```
- **Engine:** Excel COM (pywin32) when formatting/PDF must match the source exactly;
  openpyxl/pandas for reading and validation (faster, no Excel dependency). **[TBC]**
- Source workbook is always opened read-only and never saved.
- Outputs go to a new timestamped folder; never overwrite without confirmation.

## Milestones
1. **M0 Setup** — .gitignore, README, requirements.txt, config file, synthetic sample register.
2. **M1 Read & validate** — header detection, column mapping, validation report.
3. **M2 Split** — per-group workbooks with formatting, column widths, subtotals.
4. **M3 Reconcile & audit** — control totals, row counts, manifest, log file.
5. **M4 Optional outputs** — PDF export, password protection, email drafts **[TBC]**.
6. **M5 UI** — simple Tkinter GUI (pick file, pick key, run, open output) **[TBC]**.
7. **M6 Installer** — packaged Windows build and installer.

## Testing
- pytest unit tests on synthetic (fake) registers only — no real employee data in the repo.
- Edge cases: blank/merged header rows, totals rows, blank group values, special characters
  in group names (invalid filename chars), duplicate employee IDs, very large files,
  workbook already open/locked, Excel not installed.
- Golden-file tests: split then recombine == source (row count and every amount column).
- Manual UAT on a copy of a real register before release.

## Data integrity & privacy
- Reconciliation must pass before a run is marked successful: row count and each amount
  column total (e.g. Gross, Deductions, Net) per group and overall equal the source.
- Rows with no group value go to an `UNASSIGNED` file and are flagged — never dropped.
- Decimal-safe comparison of money (round to 2 dp; no float drift).
- Manifest records source hash, output hashes, user, timestamp, app version.
- Logs record counts and totals only — no names, IDs, bank details or salaries.
- `.gitignore` blocks .venv, data files (xlsx/xls/csv/pdf), logs, outputs, sample files.

## Windows installer
- Build: PyInstaller (`--onedir`, faster start, fewer AV false positives) → Inno Setup installer.
- Per-user install (no admin rights) to `%LOCALAPPDATA%\OfficeAutomationPro`; config and logs
  under `%APPDATA%`/`%LOCALAPPDATA%`, not in Program Files.
- Requirement check: Excel installed (if COM engine used); clear error otherwise.
- Code signing **[TBC]**; versioned releases with changelog.
