# Patient Report Design System

**Source of truth:** the `MUHAMMED DATA - OCT 2026` workbook (now the `Oct-2026` sheet of
`Faiz_Patient_Data_Master_Nov2025-Oct2026.xlsx`). Code version: `scripts/design_system.py`.

This design applies to **every sheet of every patient workbook** (and the PDFs exported from them).

## Tokens

| Token | Value | Used for |
|---|---|---|
| Font | Times New Roman | Everything |
| Navy | `#1F3864` | Name line text, title bar, month total bars |
| Header blue | `#2F4B7C` | Column header row |
| Teal | `#1F7A8C` | Thin spacer strip under the data, section title bars (e.g. Radiology) |
| Yellow | `#FFE699` | Combined-total bars (black bold text) |
| Grid | `#9AA5B1` | Thin borders on every cell |
| Date format | `dd-mmm-yyyy` | Real Excel dates (e.g. `04-Oct-2026`), never text |
| Amount format | `#,##0.00` | Bill amounts (AED), right-aligned |

## Month sheet layout (tab name `Mon-YYYY`)

| Row | Content | Style |
|---|---|---|
| 1 | `NAME : MUHAMMED FAIZ` (merged across) | 12 pt bold navy text, left, height 21.75 |
| 2 | `PATIENT DATA OF <MONTH YYYY>` (merged across) | 15 pt bold white on navy, centred, height 27.75 |
| 3 | `Sl No` · `Date` · `File Number` · `Patient Name` · `Consultation / Service / Department` · `Bill Amount (AED)` | 11 pt bold white on header blue, centred, wrap, height 30 |
| 4… | Data rows | 11 pt, **no zebra banding**, centred + wrap (Patient Name left, Bill Amount right), vertical centre, height 18 (taller when wrapped) |
| spacer | Empty merged row | Teal fill, height 7.5 |
| total | `TOTAL PATIENTS (VISITS) - <MONTH YYYY>` → `=COUNTA(B…)` in E | 12 pt bold white on navy, label right-aligned, height 24 |
| total | `TOTAL BILL AMOUNT (AED) - <MONTH YYYY>` → `=SUM(F…)` in F | same (only when bill amounts exist) |
| combined | `COMBINED TOTAL PATIENTS / REFERRALS`, `COMBINED TOTAL BILL AMOUNT (AED)` | 12 pt bold black on yellow (only when a second section exists) |
| section | e.g. `RADIOLOGY DATA - OUT PATIENTS` + its own header row, data and totals | 14 pt bold white on teal, height 31.5 |

Column widths: A 8 · B 16 · C 20 · D 52 · E 50 · F 18.
Multiple services in one visit go on separate lines inside the Consultation cell (`GP - Dr. X⏎ENT - Dr. Y; Labs`).

## Page setup (all sheets)

A4 **landscape**, fit to 1 page wide, margins L/R 0.4" T/B 0.5", footer `Page &P of &N`,
horizontally centred, rows 1–3 repeat on every printed page, gridlines hidden, panes frozen at `A4`.

## Rules

- Totals are always formulas (`COUNTA`, `SUM`), never typed numbers.
- The Summary sheet uses the same tokens (name line, navy title, header row, no banding, navy grand total).
- Months without bill data keep columns A–E only; add column F once amounts are available.
