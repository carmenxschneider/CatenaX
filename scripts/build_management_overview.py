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
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Management_Overview_CW42_26.pptx"

# ---------------------------------------------------------------- data
# CW 42 / 2026 PULSE slides (screenshots). PCF colours confirmed by the reporting team (lime green = amber).
# Content in short business English. Abbreviations lower case.
# Date formats: go-live "Mon YYYY"; milestones "DD Mon YYYY" (or "Mon YYYY" if the source has no day).
# ms = next critical milestone: the most critical amber/red milestone (red before amber, then earliest);
#      None = the timeline has no amber/red milestone.
# dims = status per dimension, in the order of USE_CASE_DIMS.
CW = "CW 42 / 2026"
USE_CASE_DIMS = DIM_NAMES + ["Supplier Activation"]

COLUMNS = [  # sorted: Escalation needed -> Problem solving -> On track (no escalation in CW 42)
    dict(name="Quality", abbr="quality-cx", dl="Kutritz (K-GQY)", status="ps", phase="Piloting",
         go_live=("MVP Jan 2027", "Scaling until 2030"),
         sup=(0, 2), sup_note="rescoped from 15",
         ms=("ps", "Domain expertise & ownership aligned", "30 Sep 2026"),
         dims=["ok", "ps", "ps", "ps", "ps"]),  # rescope 15 -> 2 decided in CW 41
    dict(name="Battery Passport", abbr="batt-cx", dl="Alp (ZG-R)", status="ps", phase="Piloting",
         go_live=("Feb 2027", None), sup=(0, 1), ms=("ok", "Tech requirements done", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok", "ok"]),
    dict(name="Product Passes", abbr="pass-cx", dl="Drobir (K-GEP-2)", status="ps", phase="Definition & Clarification",
         go_live=None, sup=(0, 1),
         ms=("ps", "Product passes prioritized", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok", "ok"]),
    dict(name="PURIS", abbr="puris-cx", dl="Timpe (KL-GP), Behrens (BZ-PX)", status="ps", phase="Piloting",
         go_live=None, sup=(1, 6),
         ms=("ps", "Plan for existing challenges ready", "09 Oct 2026"),
         dims=["ok", "ps", "ps", "ps", "ps"]),
    dict(name="Business Partner Data Mgmt", abbr="bpdm-cx", dl="Fehlner (I/BZ)", status="ps", phase="Implementation",
         go_live=("Feb 2027", None), sup=(0, 1), ms=("ok", "BPDM target picture defined", "31 Oct 2026"),
         dims=["ok", "ps", "ok", "ps", "ns"]),
    dict(name="Certificate Management", abbr="cert-cx", dl="Poetsch (K-DDX/5)", status="ps", phase="Scaling",
         go_live=("May 2028", None), sup=(89, 100),
         ms=("ps", "Official CX CCM release", "Sep 2026"),
         dims=["ok", "ok", "ok", "ok", "ok"]),
    dict(name="Product Carbon Footprint", abbr="pcf-cx", dl="Dettmer (K-GEN), Voeste (K-GSS)", status="ps",
         phase="Piloting", go_live=("Jan 2028", None), sup=(0, 1),
         ms=("ps", "Supplier activation pilot stage 1", "15 Oct 2026"),
         dims=["ok", "ok", "ps", "ok", "ps"]),
    dict(name="N-Tier", abbr="ntier-cx", dl="Fink (BZI2)", status="ok", phase="Definition & Clarification",
         go_live=None, sup=None, ms=("ok", "Working group framework created", "Q4 2026"),
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

# Dimensions block: "Dimensions" header line, then one thin swimlane per dimension (dot only in the cards)
DIM_HEAD = LH + 0.04   # header line inside the dimensions row
DIM_SUB_MIN = 0.22     # minimum height of one swimlane
DIM_DOT = 0.12         # dot diameter in the swimlanes
DIM_LABEL_PT = 8.5     # dimension names in the label column
LANE = "F0F3F6"        # tint of every other swimlane in the cards
LANE_DARK = "0B3A45"   # matching tint behind every other dimension name (label column)
RING = "AEB9C1"        # outline of the "not started" ring on white cards
DIM_SHORT = {"External · CX Association": "CX Association"}  # label column is narrow


def dim_rows(top, row_h, n):
    """Top y of each swimlane and the swimlane height, for a dimensions row starting at `top` (after pad_top)."""
    start = top + DIM_HEAD
    sub = (row_h - GRID["pad_top"] - GRID["pad_bottom"] - DIM_HEAD) / n
    return [start + i * sub for i in range(n)], sub


def status_dot(s, name, x, cy, d, st, ring=RING, ring_fill="FFFFFF"):
    """Filled status dot; "not started" is drawn as a hollow ring."""
    if st == "ns":
        s.shape(f"{name} Ring", x, cy - d / 2, d, d, "ellipse", fill=ring_fill, line=ring, line_w=15875)
    else:
        s.shape(f"{name} Dot", x, cy - d / 2, d, d, "ellipse", fill=DOT[st])


def para_box(s, name, x, top, w, k, paras):
    s.text(name, x, top, w, k * LH + 0.02, paras, anchor="t", line_pts=1150)


def cell(key, uc, w, board, row_h=None):
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
        if not uc["ms"]:  # no upcoming milestone in the source deck
            note = "No milestone planned"
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
        n_sub = len(board["dims"])
        min_h = DIM_HEAD + n_sub * DIM_SUB_MIN

        def draw(s, x, top):
            ys, sub = dim_rows(top, row_h or (GRID["pad_top"] + min_h + GRID["pad_bottom"]), n_sub)
            for i, (dim, st, y) in enumerate(zip(board["dims"], uc["dims"], ys)):
                if i % 2 == 0:  # zebra swimlanes: every other lane tinted, edge to edge of the card
                    s.shape(f"{n} Dim Lane {dim}", x - GRID["pad"], y, w + 2 * GRID["pad"], sub, fill=LANE)
                status_dot(s, f"{n} Dim {dim}", x, y + sub / 2, DIM_DOT, st)
        return min_h, draw

    if key == "dec":
        if not uc.get("decisions"):
            return LH, lambda s, x, top: para_box(s, f"{n} Decision None", x, top, w, 1, [[r("–", F, False, FAINT)]])
        flag, indent, gap = 0.16, 0.22, 0.06
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
    eyebrow="CATENA-X",
    title="Management Overview of Use Cases",
    columns=COLUMNS,
    dims=USE_CASE_DIMS,
    rows=[
        ("status", "Overall status"),
        ("phase", "Phase"),
        ("golive", "E2E go-live target date"),
        ("sup", "Suppliers enabled 2026"),
        ("ms", "Next (critical) milestone"),
        ("dims", "Dimensions"),
    ],
    notes=f"Use Case Pulse Management Overview {CW} – one card per use case (read top to bottom), "
          "sorted by overall status. Next (critical) milestone = most critical amber/red milestone, else the next one.",
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
        status_dot(s, f"Legend {lbl}", lx, ly + 0.0725, 0.075, "ns" if col == NOT_STARTED else "ok",
                   ring=NOT_STARTED, ring_fill=PANEL) if col == NOT_STARTED else \
            s.shape(f"Legend Dot {lbl}", lx, ly + 0.035, 0.075, 0.075, "ellipse", fill=col)
        s.text(f"Legend {lbl}", lx + 0.14, ly, col_w[i // 2] - 0.14, 0.15, [[r(lbl, f_leg, False, ON_DARK_MUTED)]],
               anchor="t", wrap=False)


GRID = dict(top_y=1.07, bottom_y=7.30, label_x=0.30, label_w=1.08, col_x0=1.42, col_x1=13.03, gap=0.045,
            pad=0.075, pad_top=0.07, pad_bottom=0.06)
# Leftover height: top row / middle band / dimensions.
SPARE_SHARE = dict(top=0.5, middle=2.5, dims=5)
FOOTER_H = 0.42   # reserved at the bottom of every card for an optional tag (e.g. agenda link)


def card_width(board):
    n = len(board["columns"])
    cw = (GRID["col_x1"] - GRID["col_x0"] - GRID["gap"] * (n - 1)) / n
    return cw, cw - 2 * GRID["pad"]


def label_runs(label):
    """A row label is "Title" or ("Title", "sub-line")."""
    title, sub = (label, None) if isinstance(label, str) else label
    paras = [[r(title, F, True, ACCENT_LIGHT)]] + ([[r(sub, F, False, ACCENT_LIGHT)]] if sub else [])
    k = n_lines(title, F, GRID["label_w"], bold=True) + (n_lines(sub, F, GRID["label_w"]) if sub else 0)
    return paras, k


def compute_layout(boards):
    """One vertical grid for all boards.

    The first row (overall status) and the last row (dimensions) sit at the same height on every slide;
    the rows in between share the band between them. Below the dimensions, FOOTER_H is kept free for
    card tags.
    """
    name_lines = dl_lines = 1
    content = {}
    for board in boards:
        _, w = card_width(board)
        cols = board["columns"]
        name_lines = max(name_lines, *(n_lines(uc["name"], F_NAME, w * SEGOE_FACTOR, bold=True) for uc in cols))
        dl_lines = max(dl_lines, *(dl_height(uc["dl"], w) for uc in cols))
        content[board["title"]] = [
            GRID["pad_top"] + max(max(cell(key, uc, w, board)[0] for uc in cols), label_runs(label)[1] * LH)
            + GRID["pad_bottom"] for key, label in board["rows"]]

    y_name = GRID["top_y"] + 0.11
    y_dl = y_name + name_lines * LH_NAME + 0.02 + LH + 0.06  # name slot, abbreviation, gap
    y_bar = y_dl + dl_lines * LH + 0.09
    rows_top = y_bar + 0.10

    top = max(h[0] for h in content.values())
    middle = max(sum(h[1:-1]) for h in content.values())
    dims = max(h[-1] for h in content.values())
    rows_bottom = GRID["bottom_y"] - FOOTER_H
    spare = rows_bottom - (rows_top + top + middle + dims)
    if spare < 0:
        print(f"WARNING: content is {-spare:.2f} in taller than the slide", file=sys.stderr)
    share = {k: max(spare, 0) * v / sum(SPARE_SHARE.values()) for k, v in SPARE_SHARE.items()}
    top, middle, dims = top + share["top"], middle + share["middle"], dims + share["dims"]

    heights = {}
    for title, h in content.items():
        mids = h[1:-1]
        weights = [1] + [2] * (len(mids) - 1)  # first middle row (phase / goal) gets less extra air
        extra = middle - sum(mids)
        mids = [m + extra * wt / sum(weights) for m, wt in zip(mids, weights)]
        heights[title] = [top] + mids + [dims]
    return dict(name_lines=name_lines, dl_lines=dl_lines, y_name=y_name, y_dl=y_dl, y_bar=y_bar,
                rows_top=rows_top, heights=heights, rows_bottom=rows_top + top + middle + dims)


def dl_height(dl, w):
    """Lines needed by a Domain / Workstream Lead entry; a list means one paragraph per entry."""
    parts = dl if isinstance(dl, list) else [dl]
    return sum(n_lines(p, F, w) for p in parts)


def build(board=USE_CASE_BOARD, layout=None):
    """Build one slide; pass the layout from compute_layout() to align it with other slides."""
    columns, rows = board["columns"], board["rows"]
    layout = layout or compute_layout([board])
    g = GRID
    cw, w = card_width(board)
    s = Slide()
    add_frame(s, board)

    name_lines, dl_lines = layout["name_lines"], layout["dl_lines"]
    y_name, y_dl, y_bar = layout["y_name"], layout["y_dl"], layout["y_bar"]
    for c, uc in enumerate(columns):
        cx = g["col_x0"] + c * (cw + g["gap"])
        x = cx + g["pad"]
        n = uc["name"]
        s.shape(f"Col {n} Card", cx, g["top_y"], cw, g["bottom_y"] - g["top_y"], "roundRect", 4000, "FFFFFF")
        s.text(f"Col {n} Name", x, y_name, w, name_lines * LH_NAME + 0.02, [[r(n, F_NAME, True, INK_STRONG)]],
               anchor="t", line_pts=1300)
        y_abbr = y_name + n_lines(n, F_NAME, w * SEGOE_FACTOR, bold=True) * LH_NAME + 0.02  # directly below the name
        para_box(s, f"Col {n} Abbr", x, y_abbr, w, 1, [[r(uc["abbr"], F, False, MUTED)]])
        dls = uc["dl"] if isinstance(uc["dl"], list) else [uc["dl"]]
        para_box(s, f"Col {n} Domain Lead", x, y_dl, w, dl_lines,
                 [[r(d, F, False, FAINT if d == "tbd" else INK)] for d in dls])
        if uc.get("tag"):  # e.g. link to an agenda item, in the reserved footer
            tw = min(text_w(uc["tag"], F, bold=True) + 0.30, w)
            ty = layout["rows_bottom"] + (FOOTER_H - 0.26) / 2
            s.shape(f"Col {n} Tag", x, ty, tw, 0.26, "roundRect", 50000, COMPLETED)
            s.text(f"Col {n} Tag Text", x, ty, tw, 0.26, [[r(uc["tag"], F, True, "FFFFFF")]], algn="ctr", wrap=False)
        s.shape(f"Col {n} Status Bar", x, y_bar, w, 0.04, "roundRect", 50000, STATUS[uc["status"]][1])

    # Left column on the dark background: labels aligned with name / Domain Lead and with each row
    for label, y in zip(board.get("head_labels", HEAD_LABELS), (y_name + 0.03, y_dl)):
        para_box(s, f"Head Label {label}", g["label_x"], y, g["label_w"], 1, [[r(label, F, True, ACCENT_LIGHT)]])

    y = layout["rows_top"]
    for (key, label), h in zip(rows, layout["heights"][board["title"]]):
        top = y + g["pad_top"]
        paras, _ = label_runs(label)
        title = label if isinstance(label, str) else label[0]
        s.text(f"Row Label {title}", g["label_x"], top, g["label_w"], h - g["pad_top"], paras,
               anchor="t", line_pts=1150)
        if key == "dims":  # dimension names once, in the label column, centred on their swimlane
            ys, sub = dim_rows(top, h, len(board["dims"]))
            for i, (dim, yy) in enumerate(zip(board["dims"], ys)):
                if i % 2 == 0:  # same zebra rhythm as in the cards, so each name lines up with its dots
                    s.shape(f"Dim Lane {dim}", g["label_x"] - 0.08, yy, g["col_x0"] - g["label_x"] + 0.08, sub,
                            fill=LANE_DARK)
                s.text(f"Dim Label {dim}", g["label_x"] + 0.06, yy, g["label_w"] - 0.06, sub,
                       [[r(DIM_SHORT.get(dim, dim), DIM_LABEL_PT, False, ACCENT_LIGHT)]], anchor="ctr", wrap=False)
        for c, uc in enumerate(columns):
            x = g["col_x0"] + c * (cw + g["gap"]) + g["pad"]
            if key != rows[0][0]:
                s.shape(f"{uc['name']} Row Divider {title}", x, y, w, 0.007, fill=BORDER)
            cell(key, uc, w, board, row_h=h if key == "dims" else None)[1](s, x, top)
        y += h
    return s


def main():
    from build_enabling_overview import ENABLING_BOARD  # same grid as the enabling slide
    layout = compute_layout([USE_CASE_BOARD, ENABLING_BOARD])
    write_pptx(build(USE_CASE_BOARD, layout).shapes, OUTPUT, USE_CASE_BOARD["notes"], bg=BG)


if __name__ == "__main__":
    main()
