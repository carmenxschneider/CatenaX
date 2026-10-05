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
    BORDER, COMPLETED, DECISIONS, DIM_NAMES, DOT, ESCALATION, FAINT, INK, INK_STRONG, MUTED, NOT_STARTED,
    ON_TRACK, PROBLEM, STATUS, TEMPLATE, USE_CASES, Slide, r, write_pptx,
)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Management_Overview_CW40_26.pptx"

# ---------------------------------------------------------------- data
# Per use case: lower-case abbreviation, Domain Lead (surname + department), E2E go-live target date.
# Domain Leads from the Catena-X use case overview (PURIS = "Capacity & Bottleneck Mgmt.",
# Product Passes = "Product Passports"; Battery Passport not listed yet).
# TODO: next critical milestones of the first seven use cases still to be taken from their PULSE decks;
#       until then the previous next milestone is shown.
MANAGEMENT = {
    "Quality": dict(abbr="quality-cx", dl="Kutritz (K-GQY)", go_live=("MVP Jan '27", "Scaling until 2030")),
    "Battery Passport": dict(abbr="battpass-cx", dl="tbd", go_live=("Feb 2027", None)),
    "Product Passes": dict(abbr="pass-cx", dl="Drobir (K-GEP)", go_live=None, go_live_note="not defined yet"),
    "PURIS": dict(abbr="puris-cx", dl="Timpe (KL-GP), Behrens (BZ-PX)", go_live=None, go_live_note="not specified"),
    "Business Partner Data Mgmt": dict(abbr="bpdm-cx", dl="Fehlner (I/BZ-P)", go_live=("Feb 2027", None)),
    "Certificate Management": dict(abbr="cert-mgmt", dl="Poetsch (K-DDX/5)", go_live=("May 2028", None)),
    "Product Carbon Footprint": dict(abbr="pcf-cx", dl="Dettmer (K-GEN), Voeste (K-GSS)", go_live=("01 Jan 2028", None)),
    "N-Tier": dict(abbr="ntier-cx", dl="Fink (BZI2)", go_live=None, go_live_note="not defined yet"),
}

# N-Tier, from Template_UseCase_PULSE_Meeting_Update_1.pptx, slide CW 40 / 2026
N_TIER = dict(
    name="N-Tier", status="ok", phase="Definition & Clarification",
    ms=None,                                   # no amber/red milestone (completed, on track, not started)
    dims=["ok", "ns", "ns", "ns"], sa="ns",    # External on track, the other four not started
    sup=None,                                  # supplier enablement N/A
)

# Same sort as before: Escalation needed -> Problem solving -> On track
COLUMNS = USE_CASES + [N_TIER]

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
TEAL = "008C82"        # theme accent2: info panel border and divider
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


def cell(key, uc, w):
    """Return (content height, draw) for one cell; draw(s, x, top) adds the shapes."""
    n = uc["name"]
    m = MANAGEMENT[n]

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

    if key == "golive":
        if m["go_live"]:
            main, second = m["go_live"]
            paras = [[r(main, F, False, INK)]] + ([[r(second, F, False, MUTED)]] if second else [])
            k = sum(n_lines(t, F, w) for t in (main, second) if t)
        else:
            paras = [[r("–  ", F, False, FAINT), r(m["go_live_note"], F, False, FAINT, None, True)]]
            k = n_lines("–  " + m["go_live_note"], F, w)
        return k * LH, lambda s, x, top: para_box(s, f"{n} Go-live", x, top, w, k, paras)

    if key == "sup":
        if uc["sup"] is None:
            paras, k = [[r("N/A", F, False, FAINT)]], 1
        else:
            done, total = uc["sup"]
            pct = round(100 * done / total)
            paras = [[r(f"{done} / {total}", F, True, COMPLETED), r(f"  {pct}%", F, False, MUTED)]]
            k = 1
            if uc.get("note"):
                paras.append([r("rescoped from 15", F, False, MUTED)])
                k += n_lines("rescoped from 15", F, w)
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
        items = [(DIM_LINES.get(dim, [dim]), st)
                 for dim, st in list(zip(DIM_NAMES, uc["dims"])) + [("Supplier Activation", uc["sa"])]]
        items = [(dim, lines, st, sum(n_lines(t, F, w - DOT_W) for t in lines)) for dim, (lines, st) in
                 zip([d for d in DIM_NAMES] + ["Supplier Activation"], items)]
        gap = 0.05

        def draw(s, x, top):
            y = top
            for dim, lines, st, k in items:
                s.shape(f"{n} Dim Dot {dim}", x, y + 0.04, 0.08, 0.08, "ellipse", fill=DOT[st])
                para_box(s, f"{n} Dim {dim}", x + DOT_W, y, w - DOT_W, k, [[r(t, F, False, INK)] for t in lines])
                y += k * LH + gap
        return sum(k for *_, k in items) * LH + gap * (len(items) - 1), draw

    if key == "dec":
        if not uc.get("decision"):
            return LH, lambda s, x, top: para_box(s, f"{n} Decision None", x, top, w, 1, [[r("–", F, False, FAINT)]])
        num = uc["decision"]
        question = DECISIONS[num - 1][1]
        flag, indent = 0.16, 0.22
        k = n_lines(question, F, w, indent=indent)

        def draw(s, x, top):
            s.shape(f"{n} Decision Flag", x, top, flag, flag, "ellipse", fill=COMPLETED)
            s.text(f"{n} Decision No", x, top, flag, flag, [[r(str(num), F_PILL, True, "FFFFFF")]], algn="ctr")
            para_box(s, f"{n} Decision", x, top, w, k, [[r(question, F, False, INK)]])
            s.shapes[-1] = hanging(s.shapes[-1], indent)
        return k * LH, draw

    raise KeyError(key)


ROWS = [
    ("status", "Overall status"),
    ("phase", "Phase"),
    ("golive", "E2E go-live target date"),
    ("sup", "Suppliers enabled 2026"),
    ("ms", "Next critical milestone"),
    ("dims", "Dimensions"),
    ("dec", "Decision required"),
]
HEAD_LABELS = ["Name", "Domain Lead"]


# ---------------------------------------------------------------- layout
def add_frame(s):
    """VW-style header on the dark background: eyebrow, white Group HEAD headline, info panel (CW + legend)."""
    s.text("Eyebrow", 0.30, 0.26, 6.0, 0.14, [[r("USE CASE PULSE · CATENA-X", 7.5, True, ACCENT_LIGHT, 150)]])
    s.text("Headline", 0.30, 0.44, 9.0, 0.46, [[r("Management Overview of Catena-X Use Cases", 26, False, ON_DARK)]],
           font=HEAD_FONT)

    # Calendar week + status legend grouped in one panel (style of the Catena-X use case overview tiles)
    px, py, pw, ph = 8.45, 0.24, 4.58, 0.66
    s.shape("Info Panel", px, py, pw, ph, "roundRect", 14000, PANEL, TEAL, 9525)
    s.text("Panel Label", px + 0.18, py + 0.11, 1.25, 0.16, [[r("Status as of", F, False, ACCENT_LIGHT)]], anchor="t")
    s.text("Panel CW", px + 0.18, py + 0.30, 1.25, 0.24, [[r("CW 40 / 2026", 12, True, ON_DARK)]], anchor="t")
    s.shape("Panel Divider", px + 1.42, py + 0.14, 0.008, ph - 0.28, fill=TEAL)

    legend = [("On Track", ON_TRACK), ("Problem Solving", PROBLEM),        # column 1
              ("Escalation needed", ESCALATION), ("Not Started", NOT_STARTED)]  # column 2
    for i, (lbl, col) in enumerate(legend):
        lx = px + 1.62 + (i // 2) * 1.40
        ly = py + 0.13 + (i % 2) * 0.24
        s.shape(f"Legend Dot {lbl}", lx, ly + 0.035, 0.09, 0.09, "ellipse", fill=col)
        s.text(f"Legend {lbl}", lx + 0.16, ly, 1.22, 0.16, [[r(lbl, F, False, ON_DARK_MUTED)]], anchor="t",
               wrap=False)


def build():
    s = Slide()
    add_frame(s)

    top_y, bottom_y = 1.07, 7.30
    label_x, label_w = 0.30, 0.95
    col_x0, col_x1, gap = 1.30, 13.03, 0.045
    cw = (col_x1 - col_x0 - gap * (len(COLUMNS) - 1)) / len(COLUMNS)
    pad = 0.075
    w = cw - 2 * pad
    pad_top, pad_bottom = 0.07, 0.06

    # Header block, top-aligned: name with the abbreviation directly below it, then Domain Lead
    # (aligned across cards and with its label) and the status bar
    name_lines = max(n_lines(uc["name"], F_NAME, w * SEGOE_FACTOR, bold=True) for uc in COLUMNS)  # Segoe UI Bold ≈ Arial Bold
    dl_lines = max(n_lines(MANAGEMENT[uc["name"]]["dl"], F, w) for uc in COLUMNS)
    y_name = top_y + 0.11
    y_dl = y_name + name_lines * LH_NAME + 0.02 + LH + 0.06
    y_bar = y_dl + dl_lines * LH + 0.09
    rows_top = y_bar + 0.10

    # Each row is as tall as its tallest cell (or its label)
    cells = {key: [cell(key, uc, w) for uc in COLUMNS] for key, _ in ROWS}
    heights = {key: pad_top + max(max(h for h, _ in cells[key]), n_lines(label, F, label_w, bold=True) * LH)
               + pad_bottom for key, label in ROWS}
    spare = (bottom_y - 0.06) - (rows_top + sum(heights.values()))
    if spare < 0:
        print(f"WARNING: content is {-spare:.2f} in taller than the slide", file=sys.stderr)
    for key in heights:  # share leftover space evenly
        heights[key] += max(spare, 0) / len(ROWS)

    for c, uc in enumerate(COLUMNS):
        cx = col_x0 + c * (cw + gap)
        x = cx + pad
        n = uc["name"]
        m = MANAGEMENT[n]
        s.shape(f"Col {n} Card", cx, top_y, cw, bottom_y - top_y, "roundRect", 4000, "FFFFFF")
        s.text(f"Col {n} Name", x, y_name, w, name_lines * LH_NAME + 0.02, [[r(n, F_NAME, True, INK_STRONG)]],
               anchor="t", line_pts=1300)
        y_abbr = y_name + n_lines(n, F_NAME, w * SEGOE_FACTOR, bold=True) * LH_NAME + 0.02
        para_box(s, f"Col {n} Abbr", x, y_abbr, w, 1, [[r(m["abbr"], F, False, MUTED)]])
        para_box(s, f"Col {n} Domain Lead", x, y_dl, w, dl_lines,
                 [[r(m["dl"], F, False, FAINT if m["dl"] == "tbd" else INK)]])
        s.shape(f"Col {n} Status Bar", x, y_bar, w, 0.04, "roundRect", 50000, STATUS[uc["status"]][1])

    # Left column on the dark background: labels aligned with name / Domain Lead and with each row
    for label, y in zip(HEAD_LABELS, (y_name + 0.03, y_dl)):
        para_box(s, f"Head Label {label}", label_x, y, label_w, 1, [[r(label, F, True, ACCENT_LIGHT)]])

    y = rows_top
    for key, label in ROWS:
        h = heights[key]
        top = y + pad_top
        s.text(f"Row Label {label}", label_x, top, label_w, h - pad_top, [[r(label, F, True, ACCENT_LIGHT)]],
               anchor="t", line_pts=1150)
        for c, uc in enumerate(COLUMNS):
            x = col_x0 + c * (cw + gap) + pad
            if key != ROWS[0][0]:
                s.shape(f"{uc['name']} Row Divider {label}", x, y, w, 0.007, fill=BORDER)
            cells[key][c][1](s, x, top)
        y += h
    return s


def main():
    write_pptx(build().shapes, OUTPUT,
               "Use Case Pulse Management Overview CW 40 / 2026 – one card per use case (read top to bottom), "
               "sorted by overall status. Next critical milestone = most critical amber/red milestone.", bg=BG)


if __name__ == "__main__":
    main()
