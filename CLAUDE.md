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
  → `scripts/add_radiology_sections.py` → `scripts/prefix_tmc.py` → `scripts/add_bill_amounts_*.py` → recalc → PDF.
