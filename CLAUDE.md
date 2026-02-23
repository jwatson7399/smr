# CLAUDE.md — SMR (School-based Medical Reimbursement)

This repo contains two parts of the same workflow: a mobile-first PWA for entering weekly SMR stats, and a Python CLI that fills `.xlsm` billing workbooks from plain-text input.

---

## SMR Web App (root directory)

A mobile-first PWA for entering SMR weekly stats directly on a phone. Deployed as static files via GitHub Pages; installable as a home screen app.

### Files

- `smr-app.html` — the entire application (single self-contained file; uses SheetJS/xlsx via CDN)
- `smr-sw.js` — service worker for offline caching (cache name `smr-v2`)
- `smr-manifest.json` — PWA manifest
- `index.html` — redirect shim to `smr-app.html`

### Architecture

All logic lives in `smr-app.html`. The app:
1. Reads an `.xlsm` workbook client-side via SheetJS
2. **Imports existing data** — session codes from the grid (row 5+, col 6+), comments from column W, and notes from cells A131-A133
3. Allows the user to enter/edit session codes, comments, and notes through a mobile UI
4. Exports results as a downloadable `.txt` file (compatible with smr-filler)

No server is involved. When updating app code, bump the `CACHE_NAME` in `smr-sw.js` so clients pick up the new version.

### Key implementation details

- **xlsm import**: `parseXlsm()` returns `{students, dates, sessions, comments, notes}`. Many cells contain formulas referencing other sheets — SheetJS reads cached formula values.
- **Notes format in xlsm**: All in cell A131 (overflow to A132-A133), `\r\n`-separated lines. Format is `date: Activity1 (time) Activity2 (time)` per line, with a preamble line to skip and date ranges like `2/23-2/24: School closed`.
- **Notes are drag-and-drop reorderable** within the same day (touch + desktop) for chronological ordering.
- **Treatment code system**: Direct (T, G1, G2, I, M, Y), Status (A, A2, O, H — 0 minutes), Indirect (C, S, R, D). Modifiers: `/` = half (15m), `//` = less than half (10m). Base session = 30 minutes.
- **Chart tab**: Shows full student×date grid with totals rows for Codes, Notes, and combined Total per date.

### Live URL

https://jwatson7399.github.io/smr/smr-app.html

---

## SMR Filler (`filler/` directory)

A Python CLI that fills an `.xlsm` billing workbook from a plain-text input file.

### Commands

```bash
cd filler
pip install openpyxl
python3 smr_filler.py <workbook.xlsm> <input.txt>   # produces <workbook>_filled.xlsm
```

### How It Works

1. Parses a 3-section input file (`## Sessions`, `## Comments`, `## Notes`)
2. Builds lookup maps from the existing workbook — student names (rows 6-49) and dates (row 5, columns G-V)
3. Fills the workbook: treatment codes into the grid, comments into column W, notes into A131+
4. Calculates daily totals (row 101) — billable session minutes from grid + durations from notes
5. Saves as `_filled.xlsm`, preserving VBA macros

### Workbook Layout (STANDARD Stats sheet)

| Region | Purpose |
|---|---|
| A6-A49 | Student names (`Last, First`) |
| Row 5, G-V | Date headers (datetime objects) |
| G6-V49 | Treatment code grid |
| W6-W49 (col 23) | Per-student comments |
| Row 101 | Daily totals (billable minutes as `h:mm`) |
| A131-A133 | Notes (free-text grouped by date) |

### Input File Format

```
## Sessions
10/1, Moret, T
10/2, Estrada, T/

## Comments
Estrada, 10/2: attempted to transition student...

## Notes
10/1, Consult re: caseload review (1hr) Treatment planning/prep (30 min)
10/2, Check and respond to emails (15 min) Documentation (1hr)
```

### Key Constants (top of `smr_filler.py`)

- `STUDENT_ROW_START/END` — 6-49
- `DATE_COL_START/END` — G-V
- `COMMENT_COL` — W (23)
- `TOTAL_ROW` — 101
- `NOTES_ROW` — 131

### Test Files

- `input_week_test.txt` — sample input file
- `FATSTINKY.xlsm` — test workbook with pre-filled data (used to verify import parsing)
