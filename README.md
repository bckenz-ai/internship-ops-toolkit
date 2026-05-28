# Internship Operations Toolkit

Google Apps Script automations built during a finance internship at **Parallel Dimensions Inc.** (Makati, Metro Manila) to support event operations, host management, and financial workflows across large-scale nightlife events.

---

## Scripts

### `AUTOCOMPUTE.gs` — Batch Commission Calculator

Automates payroll computation for a roster of event hosts. Rather than calculating commissions one by one, the script iterates through every name in a designated column, feeds each into a formula-driven calculator cell, forces a sheet flush to resolve all dependent formulas, reads the output, and writes all results back in a single batch operation.

A checkbox in cell `E3` serves as a run button — ticking it triggers the calculation and resets itself automatically.

**Key behaviors:**
- Reads host names from a configurable column (default: Column M, starting Row 3)
- Uses `SpreadsheetApp.flush()` to ensure formula evaluation before reading output
- Writes all results back in one `setValues()` call to minimize API calls
- Skips blank rows cleanly

---

### `INVITE COUNT.gs` — Baseline-Delta Reporting Tool

Generates structured 3PM and 9PM invite count reports for event hosts, comparing current live counts against a previously saved baseline to show per-host progress.

**Workflow:**

1. **Save Baseline** — Captures a snapshot of the current invite counts (Column M of the AUTOCOUNTER sheet) along with a timestamp, stored in a hidden BASELINE tab
2. **Generate Report** — Compares live counts against the baseline, computes deltas per host, and splits hosts into two lists:
   - **Eligible** (5+ invites) — shown with delta and a fire indicator for hosts with 5+ new invites since the last baseline
   - **Ineligible** (under 5 invites)

Output is displayed in a modal dialog with a copyable text area and a toggle to show or hide the total invite count. Designed to be pasted directly into group chats or internal communications.

**Key behaviors:**
- Excludes internal house accounts (e.g., `PARADIMES`, `MANILA MI VIDA`) from all counts and totals
- Normalizes usernames to lowercase and trims whitespace for consistent key matching
- Report includes a Top Host Bonus section at the bottom

---

### `GUESTLIST.gs` — Guest Registration Pipeline

Transforms raw Google Form submissions into a clean, tier-organized guest list, and keeps payment verification data in sync between the formatted report and the source database.

**Workflow:**

1. Reads all responses from the raw `GUESTLIST` sheet
2. Filters and groups entries by ticket tier (Early Bird Phase 1, Early Bird Phase 2, Regular)
3. Writes a formatted report to a separate sheet with per-tier headers showing guest counts, sequential numbering, and bold section dividers
4. Inserts checkboxes only on actual guest rows (not headers or spacer rows) for payment verification
5. Any edit to the verification column (Col M) or payment column (Col N) in the formatted sheet is automatically synced back to the raw database via `onEdit`

**Key behaviors:**
- Clears stale checkboxes before regenerating to prevent duplicate data validations
- Report auto-regenerates on any new form submission via `onEdit` on the source sheet
- Two-way sync ensures the raw database stays as the single source of truth

---

## Stack

`Google Apps Script` `Google Sheets API` `JavaScript`

---

## Setup

These scripts are intended to run as bound scripts within a Google Spreadsheet (via **Extensions > Apps Script**). Each `.gs` file corresponds to one script file in the Apps Script editor.

Sheet names and column indices are configurable at the top of each function. Refer to the inline comments for the specific cells and ranges each script expects.

---

## Context

Built as internal tooling during a finance internship at Parallel Dimensions Inc. All scripts were written to reduce manual data entry, eliminate calculation errors, and allow non-technical staff to trigger operations through simple UI controls (checkboxes, menu items, modal dialogs).
