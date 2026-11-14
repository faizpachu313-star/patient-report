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

## Naming rule
- If a patient's full name is not clear or not complete in the source (cut off, partly hidden,
  unreadable), write only the **first and second name** (e.g. `Mr MUHAMMAD RIAZ MUHAMM` → `Muhammad Riaz`).
  Do not guess the missing part. Names are Title Case without Mr/Ms/Mrs.

## Merging new uploads
- Before adding anything, match every source row against the master by date + file number, then by
  name (radiology referral lists often use another file system, e.g. Town Centre EMR #, and the
  referral date differs from the scan date). Only add rows that are genuinely missing; never duplicate.
- Rebuild order: base design workbook → `scripts/merge_uploads_*.py` → `scripts/add_radiology_sections.py`
  → recalc → PDF.
