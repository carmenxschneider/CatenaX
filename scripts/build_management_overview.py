"""Build the "Use Case Pulse – Management Overview" slide.

One white card per use case (read top to bottom), one KPI per row, on the VW
dark petrol background. Cards keep the colours and shape vocabulary of the
PULSE single slides.

Row heights are computed from the actual text: every string is wrapped with
Liberation Sans (Arial) metrics plus a 7 % margin for Segoe UI, so a layout
that fits here also fits in PowerPoint.

    python scripts/build_management_overview.py [template.pptx] [output.pptx]
"""

import os
import sys

from PIL import ImageFont

from build_overview import (
    BORDER, COMPLETED, DIM_NAMES, DOT, ESCALATION, FAINT, INK, INK_STRONG, MUTED, NOT_STARTED, ON_TRACK, PROBLEM,
    STATUS, TEMPLATE, Slide, r, write_pptx,
)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Management_Overview_CW41_26.pptx"

# ---------------------------------------------------------------- data
# CW 41 / 2026 PULSE slides (Quality: CW 40 deck, content confirmed as current; N-Tier: CW 41 screenshot).
# Content in short business English. Abbreviations lower case.
# ms = next critical milestone: the most critical amber/red milestone (red before amber, then earliest);
#      None = the timeline has no amber/red milestone.
# dims = status per dimension, in the order of USE_CASE_DIMS.
CW = "CW 41 / 2026"
USE_CASE_DIMS = DIM_NAMES + ["Supplier Activation"]

COLUMNS = [  # sorted: Escalation needed -> Problem solving -> On track
    dict(name="Quality", abbr="quality-cx", dl="Kutritz (K-GQY)", status="esc", phase="Piloting",
         go_live=("MVP Jan '27", "Scaling until 2030"),
         sup=(0, 2), sup_note="rescoped from 15",
         ms=("esc", "SQA CX MVP scale-up", "31 Jan 2027"),
         dims=["ok", "ps", "ps", "ps", "ps"],
         decisions=[(1, "Approve rescoping from 15 to 2 suppliers until the Early Warning Production MVP is live?")]),
    dict(name="Battery Passport", abbr="batt-cx", dl="Alp (ZG-R)", status="ps", phase="Piloting",
         go_live=("Feb 2027", None), sup=(0, 1), ms=None,
         dims=["ps", "ps", "ok", "ok", "ok"]),
    dict(name="Product Passes", abbr="pass-cx", dl="Drobir (K-GEP-2)", status="ps", phase="Definition & Clarification",
         go_live=None, sup=(0, 1),
         ms=("ps", "Product passes prioritized", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok", "ok"]),
    dict(name="PURIS", abbr="puris-cx", dl="Timpe (KL-GP), Behrens (BZ-PX)", status="ps", phase="Piloting",
         go_live=None, sup=(1, 6),
         ms=("ps", "Plan for existing challenges ready", "9 Oct 2026"),
         dims=["ok", "ps", "ps", "ps", "ps"],
         decisions=[(2, "Who decides how and when on the WINGS connectivity rollout roadmap?")]),
    dict(name="Business Partner Data Mgmt", abbr="bpdm-cx", dl="Fehlner (I/BZ)", status="ps", phase="Implementation",
         go_live=("Feb 2027", None), sup=(0, 1), ms=None,
         dims=["ps", "ps", "ok", "ps", "ns"]),
    dict(name="Certificate Management", abbr="cert-cx", dl="Poetsch (K-DDX/5)", status="ps", phase="Scaling",
         go_live=("May 2028", None), sup=(82, 100),
         ms=("ps", "Official CX CCM release", "Sep 2026"),
         dims=["ps", "ok", "ok", "ok", "ok"]),
    dict(name="Product Carbon Footprint", abbr="pcf-cx", dl="Dettmer (K-GEN), Voeste (K-GSS)", status="ok",
         phase="Piloting", go_live=("01 Jan 2028", None), sup=(0, 1), ms=None,
         dims=["ok", "ok", "ok", "ok", "ok"]),
    dict(name="N-Tier", abbr="ntier-cx", dl="Fink (BZI2)", status="ok", phase="Definition & Clarification",
         go_live=None, sup=None, ms=None,
         dims=["ok", "ns", "ns", "ns", "ns"]),
]
GO_LIVE_MISSING = "not defined yet"


# ---------------------------------------------------------------- typography
F = 9           # one size for every card value, secondary line and row label
F_NAME = 11     # use case name
F_PILL = 8      # status pill text
LH = 11.5 / 72  # line height (inch) of 9 pt text with 11.5 pt line spacing
LH_NAME = 13 / 72

# VW corporate frame around the cards
BG = "002733"          # Volkswagen Group dark petrol (theme dk2 / accent1)
ON_DARK = "FFFFFF"     # headline
ACCENT_LIGHT = "99D1CD"  # theme accent3 (mint): eyebrow and row labels
ON_DARK_MUTED = "CCD3D6"  # theme accent5: legend text
TEAL = "008C82"        # theme accent2: info panel divider
PANEL = "0B3A45"       # slightly lighter petrol for the info panel
HEAD_FONT = "The Group HEAD Light"

_FONT_DIR = "/usr/share/fonts/truetype/liberation"
SEGOE_FACTOR = 1.07  # Segoe UI runs up to ~7 % wider than Arial
_FONTS = {}


def _font(bold):
    if bold not in _FONTS:
        name = "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf"
        _FONTS[bold] = ImageFont.truetype(os.path.join(_FONT_DIR, name), 1000)
    return _FONTS[bold]


def text_w(text, pt, bold=False):
    return _font(bold).getlength(text) / 1000 * pt / 72 * SEGOE_FACTOR


def n_lines(text, pt, width, bold=False, indent=0.0):
    """Greedy word wrap; the first line may be shortened by a hanging indent."""
    lines, cur, avail = 1, "", width - indent
    for word in text.split():
        trial = (cur + " " + word).strip()
        if text_w(trial, pt, bold) <= avail or not cur:
            cur = trial
        else:
            lines, cur, avail = lines + 1, word, width
    return lines


def hanging(shape_xml, first_line_indent):
    """Indent only the first line of a text box (text flows under the decision flag)."""
    return shape_xml.replace('marL="0" indent="0"', f'marL="0" indent="{round(first_line_indent * 914400)}"')


# ---------------------------------------------------------------- cells
DOT_W = 0.15  # dot + gap before a dotted label
DIM_LINES = {"External · CX Association": ["External ·", "CX Association"]}  # break after the dot


def para_box(s, name, x, top, w, k, paras):
    s.text(name, x, top, w, k * LH + 0.02, paras, anchor="t", line_pts=1150)


def cell(key, uc, w, board):
    """Return (content height, draw) for one cell; draw(s, x, top) adds the shapes."""
    n = uc["name"]

    if key == "status":
        lbl, fill, fg = STATUS[uc["status"]]

        pw = min(w, text_w(lbl, F_PILL, bold=True) + 0.22)

        def draw(s, x, top):
            s.shape(f"{n} Status Pill", x, top - 0.025, pw, 0.21, "roundRect", 50000, fill)
            s.text(f"{n} Status Text", x, top - 0.025, pw, 0.21, [[r(lbl, F_PILL, True, fg)]], algn="ctr", wrap=False)
        return 0.18, draw

    if key == "phase":
        k = n_lines(uc["phase"], F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} Phase", x, top, w, k, [[r(uc["phase"], F, False, INK)]])

    if key == "goal":
        k = n_lines(uc["goal"], F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} Goal", x, top, w, k, [[r(uc["goal"], F, False, INK)]])

    if key == "kpi":
        if uc.get("kpi"):  # (title, actual, target); actual None = not reported yet
            title, actual, target = uc["kpi"]
            value = [r(actual, F, True, COMPLETED) if actual else r("tbd", F, False, FAINT, None, True),
                     r(f" / {target}", F, False, INK)]
            paras = [[r(title, F, False, MUTED)], value]
            k = n_lines(title, F, w) + n_lines((actual or "tbd") + f" / {target}", F, w)
        else:
            paras = [[r("–  ", F, False, FAINT), r(GO_LIVE_MISSING, F, False, FAINT, None, True)]]
            k = n_lines("–  " + GO_LIVE_MISSING, F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} KPI", x, top, w, k, paras)

    if key == "golive":
        if uc["go_live"]:
            main, second = uc["go_live"]
            paras = [[r(main, F, False, INK)]] + ([[r(second, F, False, MUTED)]] if second else [])
            k = sum(n_lines(t, F, w) for t in (main, second) if t)
        else:
            paras = [[r("–  ", F, False, FAINT), r(GO_LIVE_MISSING, F, False, FAINT, None, True)]]
            k = n_lines("–  " + GO_LIVE_MISSING, F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} Go-live", x, top, w, k, paras)

    if key == "sup":
        if uc["sup"] is None:
            paras, k = [[r(uc.get("sup_text", "N/A"), F, False, FAINT)]], 1
        else:
            done, total = uc["sup"]
            pct = uc.get("sup_pct", round(100 * done / total))  # as reported, if the source states it
            paras = [[r(f"{done} / {total}", F, True, COMPLETED), r(f"  {pct}%", F, False, MUTED)]]
            k = 1
            if uc.get("sup_note"):
                paras.append([r(uc["sup_note"], F, False, MUTED)])
                k += n_lines(uc["sup_note"], F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} Suppliers", x, top, w, k, paras)

    if key == "ms":
        if not uc["ms"]:  # verified: no amber/red milestone in the source deck
            note = "No critical milestone"
            k = n_lines(note, F, w)
            return k * LH, lambda s, x, top: para_box(s, f"{n} Milestone", x, top, w, k,
                                                      [[r(note, F, False, FAINT, None, True)]])
        st, ms_name, ms_date = uc["ms"]
        k = n_lines(ms_name, F, w - DOT_W) + n_lines(ms_date, F, w - DOT_W)

        def draw(s, x, top):
            s.shape(f"{n} Milestone Dot", x, top + 0.04, 0.08, 0.08, "ellipse", fill=DOT[st])
            para_box(s, f"{n} Milestone", x + DOT_W, top, w - DOT_W, k,
                     [[r(ms_name, F, False, INK)], [r(ms_date, F, False, MUTED)]])
        return k * LH, draw

    if key == "dims":
        items = [(dim, DIM_LINES.get(dim, [dim]), st) for dim, st in zip(board["dims"], uc["dims"])]
        items = [(dim, lines, st, sum(n_lines(t, F, w - DOT_W) for t in lines)) for dim, lines, st in items]
        gap = board.get("dim_gap", 0.05)

        def draw(s, x, top):
            y = top
            for dim, lines, st, k in items:
                s.shape(f"{n} Dim Dot {dim}", x, y + 0.04, 0.08, 0.08, "ellipse", fill=DOT[st])
                para_box(s, f"{n} Dim {dim}", x + DOT_W, y, w - DOT_W, k, [[r(t, F, False, INK)] for t in lines])
                y += k * LH + gap
        return sum(k for *_, k in items) * LH + gap * (len(items) - 1), draw

    if key == "dec":
        if not uc.get("decisions"):
            return LH, lambda s, x, top: para_box(s, f"{n} Decision None", x, top, w, 1, [[r("–", F, False, FAINT)]])
        flag, indent, gap = 0.16, 0.22, 0.08
        items = [(num, q, n_lines(q, F, w, indent=indent)) for num, q in uc["decisions"]]

        def draw(s, x, top):
            y = top
            for num, question, k in items:
                s.shape(f"{n} Decision {num} Flag", x, y, flag, flag, "ellipse", fill=COMPLETED)
                s.text(f"{n} Decision {num} No", x, y, flag, flag, [[r(str(num), F_PILL, True, "FFFFFF")]], algn="ctr")
                para_box(s, f"{n} Decision {num}", x, y, w, k, [[r(question, F, False, INK)]])
                s.shapes[-1] = hanging(s.shapes[-1], indent)
                y += k * LH + gap
        return sum(k for *_, k in items) * LH + gap * (len(items) - 1), draw

    raise KeyError(key)


HEAD_LABELS = ["Name", "Domain Lead"]

USE_CASE_BOARD = dict(
    eyebrow="USE CASE PULSE · CATENA-X",
    title="Management Overview of Use Cases",
    columns=COLUMNS,
    dims=USE_CASE_DIMS,
    rows=[
        ("status", "Overall status"),
        ("phase", "Phase"),
        ("golive", "E2E go-live target date"),
        ("sup", "Suppliers enabled 2026"),
        ("ms", "Next critical milestone"),
        ("dims", "Dimensions"),
        ("dec", "Decision required"),
    ],
    notes=f"Use Case Pulse Management Overview {CW} – one card per use case (read top to bottom), "
          "sorted by overall status. Next critical milestone = most critical amber/red milestone.",
)


# ---------------------------------------------------------------- layout
def add_frame(s, board):
    """VW-style header on the dark background: eyebrow, white Group HEAD headline, info panel (CW + legend)."""
    s.text("Eyebrow", 0.30, 0.26, 6.0, 0.14, [[r(board["eyebrow"], 7.5, True, ACCENT_LIGHT, 150)]])
    s.text("Headline", 0.30, 0.44, 9.0, 0.46, [[r(board["title"], 26, False, ON_DARK)]],
           font=HEAD_FONT)

    # Calendar week + status legend: one compact, quiet panel at the top right (no border, slightly lighter petrol)
    f_leg = 8
    legend = [("On Track", ON_TRACK), ("Problem Solving", PROBLEM),        # column 1
              ("Escalation needed", ESCALATION), ("Not Started", NOT_STARTED)]  # column 2
    col_w = [0.14 + max(text_w(lbl, f_leg) for lbl, _ in legend[i:i + 2]) + 0.04 for i in (0, 2)]
    cw_w = max(text_w(CW, 10, bold=True), text_w("Status as of", 7.5)) + 0.04
    pad, gap = 0.14, 0.16
    pw = pad + cw_w + gap + 0.008 + gap + col_w[0] + 0.12 + col_w[1] + pad
    ph, py = 0.50, 0.30
    px = 13.03 - pw
    s.shape("Info Panel", px, py, pw, ph, "roundRect", 18000, PANEL)
    s.text("Panel Label", px + pad, py + 0.07, cw_w, 0.14, [[r("Status as of", 7.5, False, ACCENT_LIGHT)]], anchor="t")
    s.text("Panel CW", px + pad, py + 0.23, cw_w, 0.20, [[r(CW, 10, True, ON_DARK)]], anchor="t")
    dx = px + pad + cw_w + gap
    s.shape("Panel Divider", dx, py + 0.10, 0.008, ph - 0.20, fill=TEAL)
    for i, (lbl, col) in enumerate(legend):
        lx = dx + 0.008 + gap + (0 if i < 2 else col_w[0] + 0.12)
        ly = py + 0.08 + (i % 2) * 0.18
        s.shape(f"Legend Dot {lbl}", lx, ly + 0.035, 0.075, 0.075, "ellipse", fill=col)
        s.text(f"Legend {lbl}", lx + 0.14, ly, col_w[i // 2] - 0.14, 0.15, [[r(lbl, f_leg, False, ON_DARK_MUTED)]],
               anchor="t", wrap=False)


def build(board=USE_CASE_BOARD):
    columns, rows = board["columns"], board["rows"]
    s = Slide()
    add_frame(s, board)

    top_y, bottom_y = 1.07, 7.30
    label_x, label_w = 0.30, 0.95
    col_x0, col_x1, gap = 1.30, 13.03, 0.045
    cw = (col_x1 - col_x0 - gap * (len(columns) - 1)) / len(columns)
    pad = 0.075
    w = cw - 2 * pad
    pad_top, pad_bottom = 0.07, 0.06

    # Header block, top-aligned: name with the abbreviation directly below it, then Domain Lead
    # (aligned across cards and with its label) and the status bar
    name_lines = max(n_lines(uc["name"], F_NAME, w * SEGOE_FACTOR, bold=True) for uc in columns)  # Segoe UI Bold ≈ Arial Bold
    dl_lines = max(n_lines(uc["dl"], F, w) for uc in columns)
    y_name = top_y + 0.11
    y_dl = y_name + name_lines * LH_NAME + 0.02 + LH + 0.06
    y_bar = y_dl + dl_lines * LH + 0.09
    rows_top = y_bar + 0.10

    # Each row is as tall as its tallest cell (or its label)
    cells = {key: [cell(key, uc, w, board) for uc in columns] for key, _ in rows}
    heights = {key: pad_top + max(max(h for h, _ in cells[key]), n_lines(label, F, label_w, bold=True) * LH)
               + pad_bottom for key, label in rows}
    spare = (bottom_y - 0.06) - (rows_top + sum(heights.values()))
    if spare < 0:
        print(f"WARNING: content is {-spare:.2f} in taller than the slide", file=sys.stderr)
    for key in heights:  # share leftover space evenly
        heights[key] += max(spare, 0) / len(rows)

    for c, uc in enumerate(columns):
        cx = col_x0 + c * (cw + gap)
        x = cx + pad
        n = uc["name"]
        s.shape(f"Col {n} Card", cx, top_y, cw, bottom_y - top_y, "roundRect", 4000, "FFFFFF")
        s.text(f"Col {n} Name", x, y_name, w, name_lines * LH_NAME + 0.02, [[r(n, F_NAME, True, INK_STRONG)]],
               anchor="t", line_pts=1300)
        y_abbr = y_name + n_lines(n, F_NAME, w * SEGOE_FACTOR, bold=True) * LH_NAME + 0.02
        para_box(s, f"Col {n} Abbr", x, y_abbr, w, 1, [[r(uc["abbr"], F, False, MUTED)]])
        para_box(s, f"Col {n} Domain Lead", x, y_dl, w, dl_lines,
                 [[r(uc["dl"], F, False, FAINT if uc["dl"] == "tbd" else INK)]])
        s.shape(f"Col {n} Status Bar", x, y_bar, w, 0.04, "roundRect", 50000, STATUS[uc["status"]][1])

    # Left column on the dark background: labels aligned with name / Domain Lead and with each row
    for label, y in zip(HEAD_LABELS, (y_name + 0.03, y_dl)):
        para_box(s, f"Head Label {label}", label_x, y, label_w, 1, [[r(label, F, True, ACCENT_LIGHT)]])

    y = rows_top
    for key, label in rows:
        h = heights[key]
        top = y + pad_top
        s.text(f"Row Label {label}", label_x, top, label_w, h - pad_top, [[r(label, F, True, ACCENT_LIGHT)]],
               anchor="t", line_pts=1150)
        for c, uc in enumerate(columns):
            x = col_x0 + c * (cw + gap) + pad
            if key != rows[0][0]:
                s.shape(f"{uc['name']} Row Divider {label}", x, y, w, 0.007, fill=BORDER)
            cells[key][c][1](s, x, top)
        y += h
    return s


def main():
    write_pptx(build(USE_CASE_BOARD).shapes, OUTPUT, USE_CASE_BOARD["notes"], bg=BG)


if __name__ == "__main__":
    main()
