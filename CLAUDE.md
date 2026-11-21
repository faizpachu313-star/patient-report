# Patient report repo — rules

## Global rule: design system
Every Excel workbook/sheet created or edited in this repo (and every PDF exported from one)
MUST follow `DESIGN_SYSTEM.md`, whose source of truth is the `MUHAMMED DATA - OCT 2026` workbook
(the `Oct-2026` sheet of the master workbook). Use the helpers in `scripts/design_system.py`
instead of hand-picking fonts, colours, widths or page setup. Do not reintroduce the old
style (zebra banding, text dates, portrait pages).

## Workflow
- After changing the master `.xlsx`, recalculate formulas and re-export the matching `.pdf`
  (`soffice --headless --convert-to pdf <file>.xlsx`) so the two stay in sync.
- Keep the Summary sheet's per-month formulas pointing at each month's `TOTAL PATIENTS (VISITS)` cell.

## File numbers
- Nov 2025 - Jun 2026 file numbers carry the `TMC` prefix (`TMC124285`, no space). Emirates IDs
  (`784-...`) and other non-TMC IDs are left as they are.

- Software: the old TMC system was used until Jun 2026; the new system (TVIP numbers) from Jul 2026.
  The new system does NOT show any pre-July data, so Part A amounts can only come from the old TMC
  system (or other records of it). Invoice screens from the new system are the valid source for Part B.

- **Source of truth:** the clinic system's invoice screens are the real data; staff Excel sheets can
  contain mistakes. Reconcile each patient's FULL invoice list both ways: invoice days missing in the
  main file are added (amount = sum of that day's invoices, breakdown in a note); main-file rows with no
  invoice are NEVER deleted: they are real visits (free follow-ups, or services billed on another
  day's invoice) and get amount 0 with a note.

- Work order: finish Part B (Jul 2026 onwards) month by month first, then Part A. The current
  (in-progress) month is kept in the Part B TEMP file even when nothing is pending.

## Change permissions (user rule, 21 Nov 2026)
- ALLOWED without asking, only when the data is confirmed valid in the user's images (invoice
  screens): add new visit rows, edit amounts, and correct visit details (date, file no., name,
  department/doctor). Keep the old value in a cell note when correcting.
- NOT allowed without the user's explicit confirmation: deleting anything from the main workbook
  (rows, visits, amounts, sheets). Ask first and list exactly what would be removed.

## Naming rule
- If a patient's full name is not clear or not complete in the source (cut off, partly hidden,
  unreadable), write only the **first and second name** (e.g. `Mr MUHAMMAD RIAZ MUHAMM` → `Muhammad Riaz`).
  Do not guess the missing part. Names are Title Case without Mr/Ms/Mrs.

## Merging new uploads
- Before adding anything, match every source row against the master by date + file number, then by
  name (radiology referral lists often use another file system, e.g. Town Centre EMR #, and the
  referral date differs from the scan date). Only add rows that are genuinely missing; never duplicate.
- Match tolerantly as well: exported files can carry an extra trailing digit on file numbers
  (`TVIP010174555` = `TVIP01017455`), run-together names (`Jeganathanperumalperumal`) or a different
  section (a 'visit' that is really a radiology referral). Same patient + same day = already there.
- **Duplicate rule:** same patient + same date + same doctor/visit = a double entry → keep only one
  (merge any extra detail into the kept row). The same patient on different dates is a separate visit
  and is always kept, however many times they come.
- Missing bill amounts stay blank until the data is cleared; keep whatever amounts exist. When one visit
  lists an amount per doctor (`454.2/75`), enter the total and keep the breakdown in a cell note.
- Rebuild order: base design workbook → `scripts/remove_duplicates.py` → `scripts/merge_uploads_*.py`
  → `scripts/add_month_*.py` (new months) → `scripts/add_radiology_sections.py` → `scripts/prefix_tmc.py` → `scripts/add_bill_amounts_*.py` → `scripts/add_invoice_visits_*.py` → recalc → PDF.

## Temporary amount files (Part A / Part B)
- `TEMP_Part_A_Nov2025-Jun2026.xlsx` and `TEMP_Part_B_Jul2026-<last month>.xlsx` are TEMPORARY working
  copies for filling pending bill amounts (`scripts/export_temp_parts.py`). Yellow = amount pending.
- They are generated from the main workbook and never edited back into it by hand: when the user
  returns one, check every amount against the main file row by row (same date, file, name), add the
  amounts via an `add_bill_amounts_*.py` step, rebuild, then regenerate the TEMP files. A month with no
  pending amounts drops out of its TEMP file automatically; delete a TEMP file once it has no months.
- The main workbook must never be changed by the export (verify its checksum before/after).
