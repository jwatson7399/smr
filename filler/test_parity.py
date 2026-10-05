#!/usr/bin/env python3
"""Both calculators agree on filler/input_week_test.txt.

Hand totals. Full session is 30 even when column C says 45 (entry 021).
M2 and M3 dedup once per day. A(M) is 0. // bills as 15.

3/2 codes: T 30 + G1 once 30 + T/ 15 + TT/ (30+15) 45 = 120. Note (30 min) = 30. Total 150.
3/3 codes: TD (30+30) 60 + G2 once 30 + G1/ 15 + A 0 + A1 0 = 105. Note (1hr) = 60. Total 165.
3/4 codes: A2 0 + O 0 + H 0 + S 30 + S/ 15 + C 30 = 75. Note (1hr 15 min) = 75. Total 150.
3/5 codes: CC/ (30+15) 45 + R/C (15+30) 45 + D// billed as 15 = 105. Note (45 mins) = 45. Total 150.
3/6 codes: plain T for the student whose column C is 45 min = 30. No note. Total 30.
3/9 codes: M2 once 30 + M3 once 30 + A(M) 0 = 60. Total 60.
3/10 codes: TT/C (30+15+30) 75 + D/CC/ (15+30+15) 60 + G1C (30+30) 60 = 195. Total 195.
3/11 codes: MM/ (30+15) 45 + C/DD (15+30+30) 75 + MDD (30+30+30) 90 = 210. Total 210.
3/12 codes: two G1 plus one G1D = G1 once 30 + D 30 = 60. Total 60.
3/13 codes: one name on rows 12 (Ind) and 13 (Gp). T on row 12 30 + G1 on row 13 30 = 60.
  Each row keeps its own code and each counts once (entry 048). Total 60.
3/16 codes: G1/ listed first, then G1, one group. It counts its longest entry once = 30.
  First-listed would give 15 (entry 048). Total 30.
The 6/8-6/12 note is a range and adds nothing to any of these days.
"""

import io
import json
import subprocess
import sys
import tempfile
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

import smr_filler

FIXTURE = Path(__file__).with_name("input_week_test.txt")
APP = Path(__file__).resolve().parents[1] / "smr-app.html"

# date -> total minutes, codes plus notes
EXPECTED = {
    "3/2": 150,
    "3/3": 165,
    "3/4": 150,
    "3/5": 150,
    "3/6": 30,
    "3/9": 60,
    "3/10": 195,
    "3/11": 210,
    "3/12": 60,
    "3/13": 60,
    "3/16": 30,
}
# date -> code minutes only, before notes
EXPECTED_CODES = {
    "3/2": 120,
    "3/3": 105,
    "3/4": 75,
    "3/5": 105,
    "3/6": 30,
    "3/9": 60,
    "3/10": 195,
    "3/11": 210,
    "3/12": 60,
    "3/13": 60,
    "3/16": 30,
}

# (name, column B). Sheet rows start at 6 in this order. Evans Kai is one
# invented student on two rows, as in the therapist's workbooks.
STUDENTS = [
    ("Alvarez Mia", "1xWk Ind"),
    ("Alvarez Leo", "1xWk Gp"),
    ("Brooks Jonah (A)", "1xWk Ind"),
    ("Brooks Jonah (B)", "1xWk Gp"),
    ("Chen Priya", "1xWk Ind"),
    ("Diaz Omar", "1xWk Ind"),
    ("Evans Kai", "1xWk Ind"),
    ("Evans Kai", "1xWk Gp"),
]
ROW_OF = {name: smr_filler.STUDENT_ROW_START + i for i, (name, _) in enumerate(STUDENTS)}
EVANS_IND, EVANS_GP = 12, 13
DATES = [(3, 2), (3, 3), (3, 4), (3, 5), (3, 6), (3, 9), (3, 10), (3, 11), (3, 12), (3, 13), (3, 16)]


def build_sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = smr_filler.SHEET_NAME
    date_map = {}
    for index, (month, day) in enumerate(DATES):
        col = smr_filler.DATE_COL_START + index
        ws.cell(row=5, column=col, value=datetime(2026, month, day))
        date_map[(month, day)] = col
    for index, (name, freq) in enumerate(STUDENTS):
        row = smr_filler.STUDENT_ROW_START + index
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=freq)
        # Diaz Omar's column C says 45 min. A plain T must still bill 30.
        duration = "45 min" if name == "Diaz Omar" else "30 min"
        ws.cell(row=row, column=3, value=duration)
    return ws, date_map


def fill(sessions, comments):
    """Write through the filler's own row lookup. Returns (ws, date_map, log)."""
    ws, date_map = build_sheet()
    student_map, ambiguous = smr_filler.build_student_map(ws)
    row_names = smr_filler.build_row_names(ws)
    buf = io.StringIO()
    with redirect_stdout(buf):
        smr_filler.fill_treatment_codes(ws, sessions, student_map, ambiguous, date_map, row_names)
        smr_filler.fill_comments(ws, comments, student_map, ambiguous, row_names)
    return ws, date_map, buf.getvalue()


def notes_text(notes):
    lines = []
    for date_str, text in notes:
        lines.append(date_str)
        lines.append(text)
    return "\n".join(lines)


def filler_totals():
    sessions, comments, notes = smr_filler.parse_input(str(FIXTURE))
    ws, date_map, log = fill(sessions, comments)
    buf = io.StringIO()
    with redirect_stdout(buf):
        totals = smr_filler.calculate_daily_totals(ws, ws, date_map, notes_text(notes))
    return totals, log + buf.getvalue(), sessions, ws, date_map


def app_section():
    html = APP.read_text()
    parts = [
        ("function fmtDate(d)", "let lastReportedSaveError"),
        ("function parseNoteText", "function dateTotalMins"),
        ("function generateTxt", "function downloadTxt"),
        ("// Billing numbers and code classes.", "// ── Counts ──"),
    ]
    return "\n".join(html[html.index(a):html.index(b)] for a, b in parts)


def app_results(sessions):
    script = r"""
function escHtml(s) { return s; }
let appData = {students: [], dates: [], sessions: {}, comments: [], notes: {}};
""" + app_section() + r"""
const input = __INPUT__;
function jsDate(label) {
  const parts = label.split('/').map(Number);
  return new Date(2026, parts[0] - 1, parts[1]);
}
appData.students = input.students;
appData.dates = input.dates.map(jsDate);
const byRow = row => appData.students.find(s => s.row === row);
for (const s of input.sessions) {
  appData.sessions[sessionKey(byRow(s.row), jsDate(s.date))] = s.code;
}
const codes = {};
for (const label of input.dates) {
  const tot = calcDayCodeTotals(jsDate(label));
  codes[label] = tot.direct + tot.indirect;
}
appData.comments = [
  {student: 'Evans Kai', row: 12, date: '3/13', text: 'seen alone'},
  {student: 'Evans Kai', row: 13, date: '3/13', text: 'seen in group'},
];
const txt = generateTxt();

// A schema 2 record keyed by name. TC/ belongs on the Ind row, G1C on the Gp row.
const legacy = migrateToRows({
  students: [
    {name: 'Evans Kai', freq: '1xWk Ind'},
    {name: 'Evans Kai', freq: '1xWk Gp'},
    {name: 'Chen Priya', freq: '1xWk Ind'},
  ],
  dates: [jsDate('3/12').toISOString(), jsDate('3/13').toISOString()],
  sessions: {'Evans Kai|2026-03-13': 'TC/', 'Evans Kai|2026-03-12': 'G1C', 'Chen Priya|2026-03-13': 'T'},
  comments: [{student: 'Evans Kai', date: '3/12', text: 'x'}, {student: 'Chen Priya', date: '3/13', text: 'y'}],
});

const rc = tokenizeCode('R/C');
const dbl = tokenizeCode('D//');
const ttc = tokenizeCode('TT/C');
const am = tokenizeCode('A(M)');
const notes = {
  a: parseNoteMins('30 min'),
  b: parseNoteMins('1hr'),
  c: parseNoteMins('1hr 15 min'),
  d: parseNoteMins('45 mins'),
};
const imported = parseNoteText('3/6\nTask (30 min)\n6/8-6/12 Consult week (60 min)\n3/10: Emails (15 min)\n');
console.log(JSON.stringify({codes, txt, legacy, rc, dbl, ttc, am, notes, imported}));
"""
    payload = {
        "students": [{"name": name, "freq": freq, "row": smr_filler.STUDENT_ROW_START + i}
                     for i, (name, freq) in enumerate(STUDENTS)],
        "dates": [f"{m}/{d}" for m, d in DATES],
        "sessions": [],
    }
    for date_str, name, code, row in sessions:
        payload["sessions"].append({"date": date_str, "row": row or ROW_OF[name], "code": code})
    script = script.replace("__INPUT__", json.dumps(payload))
    result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        sys.stderr.write(result.stderr)
        raise SystemExit(result.returncode)
    return json.loads(result.stdout)


def check_rows():
    """One name on two rows: each row keeps its own code and comment."""
    totals, fixture_log, sessions, ws, date_map = filler_totals()
    log = fixture_log
    col = date_map[(3, 13)]
    assert ws.cell(row=EVANS_IND, column=col).value == "T"
    assert ws.cell(row=EVANS_GP, column=col).value == "G1"
    assert ws.cell(row=EVANS_IND, column=smr_filler.COMMENT_COL).value == "3/13: seen alone"
    assert ws.cell(row=EVANS_GP, column=smr_filler.COMMENT_COL).value == "3/13: seen in group"
    assert "does not hold" not in log and "more than one row" not in log

    # An old export with no row line cannot pick a row for that name. It warns
    # and writes nothing, rather than stacking both lines on the lower row.
    ws, date_map, log = fill([("3/13", "Evans Kai", "T", None), ("3/13", "Evans Kai", "G1", None)], [])
    assert ws.cell(row=EVANS_IND, column=col).value is None
    assert ws.cell(row=EVANS_GP, column=col).value is None
    assert "on more than one row" in log

    # An old export still works for every name on one row.
    ws, date_map, log = fill([("3/13", "Chen Priya", "T", None)], [("Chen Priya", "3/13: ok", None)])
    assert ws.cell(row=ROW_OF["Chen Priya"], column=col).value == "T"
    assert ws.cell(row=ROW_OF["Chen Priya"], column=smr_filler.COMMENT_COL).value == "3/13: ok"

    # A row line that names the wrong row falls back to the name, with a warning.
    ws, date_map, log = fill([("3/13", "Chen Priya", "T", EVANS_GP)], [])
    assert ws.cell(row=ROW_OF["Chen Priya"], column=col).value == "T"
    assert ws.cell(row=EVANS_GP, column=col).value is None
    assert "does not hold" in log
    return totals, fixture_log, sessions


def check_round_trip(txt):
    """The app's .txt export goes to the right rows in the filler."""
    with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False) as handle:
        handle.write(txt)
        path = handle.name
    sessions, comments, _notes = smr_filler.parse_input(path)
    Path(path).unlink()
    tagged = {(date_str, row): code for date_str, _name, code, row in sessions}
    assert tagged[("3/13", EVANS_IND)] == "T"
    assert tagged[("3/13", EVANS_GP)] == "G1"
    assert all(row is not None for *_rest, row in sessions)
    ws, date_map, log = fill(sessions, comments)
    col = date_map[(3, 13)]
    assert ws.cell(row=EVANS_IND, column=col).value == "T"
    assert ws.cell(row=EVANS_GP, column=col).value == "G1"
    assert ws.cell(row=EVANS_IND, column=smr_filler.COMMENT_COL).value == "3/13: seen alone"
    assert ws.cell(row=EVANS_GP, column=smr_filler.COMMENT_COL).value == "3/13: seen in group"
    assert "WARN" not in log.replace("WARN: // in row", ""), log


def check_migration(legacy):
    rows = [s["row"] for s in legacy["students"]]
    assert rows == [6, 7, 8]
    assert legacy["sessions"] == {
        "r6|2026-03-13": "TC/",
        "r7|2026-03-12": "G1C",
        "r8|2026-03-13": "T",
    }
    moved = {(m["row"], m["dateKey"], m.get("code"), bool(m.get("comment"))) for m in legacy["rowMoves"]}
    assert moved == {
        (6, "2026-03-13", "TC/", False),
        (7, "2026-03-12", "G1C", False),
        (7, "2026-03-12", None, True),
    }
    assert [c["row"] for c in legacy["comments"]] == [7, 8]


def main():
    totals, warnings, sessions = check_rows()
    assert smr_filler.parse_code_cell("R/C") == [("R", 0.5), ("C", 1.0)]
    assert smr_filler.parse_code_cell("C/I") == [("C", 0.5), ("I", 1.0)]
    assert smr_filler.parse_code_cell("D//") == [("D", 0.25)]
    assert smr_filler.parse_code_cell("A1") == [("A1", 1.0)]
    assert smr_filler.parse_code_cell("TT/") == [("T", 1.0), ("T", 0.5)]
    assert smr_filler.parse_code_cell("S") == [("S", 1.0)]
    assert smr_filler.parse_code_cell("S/") == [("S", 0.5)]
    assert smr_filler.parse_code_cell("A(M)") == [("A(M)", 1.0)]
    assert smr_filler.parse_code_cell("TT/C") == [("T", 1.0), ("T", 0.5), ("C", 1.0)]
    assert smr_filler.parse_code_cell("D/CC/") == [("D", 0.5), ("C", 1.0), ("C", 0.5)]
    assert smr_filler.parse_code_cell("MM/") == [("M", 1.0), ("M", 0.5)]
    assert smr_filler.parse_code_cell("C/DD") == [("C", 0.5), ("D", 1.0), ("D", 1.0)]
    assert smr_filler.parse_code_cell("MDD") == [("M", 1.0), ("D", 1.0), ("D", 1.0)]
    assert smr_filler.parse_code_cell("G1C") == [("G1", 1.0), ("C", 1.0)]
    assert smr_filler.parse_code_cell("G1D") == [("G1", 1.0), ("D", 1.0)]
    assert smr_filler.parse_code_cell("M2") == [("M2", 1.0)]
    assert "row " in warnings and "column " in warnings
    span = "3/6\nTask (30 min)\n6/8-6/12 Consult week (60 min)\n3/10: Emails (15 min)\n"
    assert smr_filler.parse_notes_durations(span, 3, 6) == 30
    assert smr_filler.parse_notes_durations(span, 3, 10) == 15
    assert smr_filler.parse_notes_durations(span, 6, 8) == 0
    assert smr_filler.parse_notes_durations(span, 6, 12) == 0
    for key, expected in EXPECTED.items():
        assert totals[key] == expected, (key, totals[key], expected)
    app = app_results(sessions)
    for key, expected in EXPECTED_CODES.items():
        assert app["codes"][key] == expected, (key, app["codes"][key], expected)
    check_round_trip(app["txt"])
    check_migration(app["legacy"])
    assert app["rc"][0]["modifier"] == "/" and app["rc"][1]["modifier"] is None
    assert app["dbl"][0]["invalid"] == "//"
    assert app["dbl"][0]["modifier"] == "/"
    assert [tok["code"] for tok in app["ttc"]] == ["T", "T", "C"]
    assert [tok["modifier"] for tok in app["ttc"]] == [None, "/", None]
    assert app["am"] == [{"code": "A(M)", "modifier": None}]
    assert app["notes"] == {"a": 30, "b": 60, "c": 75, "d": 45}
    assert app["imported"]["3/6"][0]["activity"] == "Task"
    assert "6/8" not in app["imported"]
    assert app["imported"]["3/10"][0]["time"] == "15 min"
    assert len(sessions) == 39
    print("parity ok", totals)


if __name__ == "__main__":
    main()
