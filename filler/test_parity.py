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
The 6/8-6/12 note is a range and adds nothing to any of these days.
"""

import io
import json
import subprocess
import sys
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
}

STUDENTS = [
    "Alvarez Mia",
    "Alvarez Leo",
    "Brooks Jonah (A)",
    "Brooks Jonah (B)",
    "Chen Priya",
    "Diaz Omar",
]
DATES = [(3, 2), (3, 3), (3, 4), (3, 5), (3, 6), (3, 9), (3, 10), (3, 11), (3, 12)]


def build_sheet(sessions):
    wb = Workbook()
    ws = wb.active
    ws.title = smr_filler.SHEET_NAME
    date_map = {}
    for index, (month, day) in enumerate(DATES):
        col = smr_filler.DATE_COL_START + index
        ws.cell(row=5, column=col, value=datetime(2026, month, day))
        date_map[(month, day)] = col
    rows = {}
    for index, name in enumerate(STUDENTS):
        row = smr_filler.STUDENT_ROW_START + index
        rows[name] = row
        ws.cell(row=row, column=1, value=name)
        # Diaz Omar's column C says 45 min. A plain T must still bill 30.
        duration = "45 min" if name == "Diaz Omar" else "30 min"
        ws.cell(row=row, column=3, value=duration)
    for date_str, name, code in sessions:
        month, day = (int(part) for part in date_str.split("/"))
        ws.cell(row=rows[name], column=date_map[(month, day)], value=code)
    return ws, date_map


def notes_text(notes):
    lines = []
    for date_str, text in notes:
        lines.append(date_str)
        lines.append(text)
    return "\n".join(lines)


def filler_totals():
    sessions, _comments, notes = smr_filler.parse_input(str(FIXTURE))
    ws, date_map = build_sheet(sessions)
    buf = io.StringIO()
    with redirect_stdout(buf):
        totals = smr_filler.calculate_daily_totals(ws, ws, date_map, notes_text(notes))
    return totals, buf.getvalue(), sessions


def app_code_totals(sessions):
    html = APP.read_text()
    note_start = html.index("function parseNoteText")
    note_end = html.index("function dateTotalMins")
    start = html.index("// Billing numbers and code classes.")
    end = html.index("// ── Counts ──")
    section = html[note_start:note_end] + "\n" + html[start:end]
    script = r"""
function fmtDateKey(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
function sessionKey(studentName, date) {
  return studentName + '|' + fmtDateKey(date);
}
let appData = {students: [], sessions: {}};
""" + section + r"""
const rows = __ROWS__;
const labels = [...new Set(rows.map(r => r.date))];
const names = [...new Set(rows.map(r => r.name))];
appData.students = names.map(name => ({name}));
function jsDate(label) {
  const parts = label.split('/').map(Number);
  return new Date(2026, parts[0] - 1, parts[1]);
}
const out = {};
for (const label of labels) {
  const date = jsDate(label);
  appData.sessions = {};
  for (const row of rows) {
    if (row.date === label) appData.sessions[sessionKey(row.name, date)] = row.code;
  }
  const tot = calcDayCodeTotals(date);
  out[label] = tot.direct + tot.indirect;
}
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
console.log(JSON.stringify({codes: out, rc, dbl, ttc, am, notes, imported}));
"""
    payload = json.dumps([
        {"date": date_str, "name": name, "code": code}
        for date_str, name, code in sessions
    ])
    script = script.replace("__ROWS__", payload)
    result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        sys.stderr.write(result.stderr)
        raise SystemExit(result.returncode)
    return json.loads(result.stdout)


def main():
    totals, warnings, sessions = filler_totals()
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
    app = app_code_totals(sessions)
    for key, expected in EXPECTED_CODES.items():
        assert app["codes"][key] == expected, (key, app["codes"][key], expected)
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
    assert len(sessions) == 35
    print("parity ok", totals)


if __name__ == "__main__":
    main()
