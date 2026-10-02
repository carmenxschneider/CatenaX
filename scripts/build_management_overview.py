"""Build the "Use Case Pulse – Management Overview" slide (variant 2).

Transposed version of build_overview.py: one column per use case, one row per
KPI. Each use case reads top to bottom as its own column. Uses the same
template, colours and shape vocabulary as variant 1.

    python scripts/build_management_overview.py [template.pptx] [output.pptx]
"""

import sys

from build_overview import (
    BORDER, COMPLETED, DECISIONS, DIM_NAMES, DOT, FAINT, INK, INK_STRONG, MUTED, NOT_STARTED, STATUS, TEMPLATE,
    TILE, USE_CASES, Slide, add_header, r, write_pptx,
)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Management_Overview_CW40_26.pptx"

# Fixed font sizes for every column (pt) — no per-cell shrinking
F_LABEL, F_NAME, F_ABBR, F_CELL, F_SUB = 9, 11, 8, 9, 8

# Lower-case abbreviations for this slide (cert-mgmt is the original source abbreviation)
ABBR = {
    "Quality": "quality-cx",
    "Battery Passport": "battpass-cx",
    "Product Passes": "pass-cx",
    "PURIS": "puris-cx",
    "Business Partner Data Mgmt": "bpdm-cx",
    "Certificate Management": "cert-mgmt",
    "Product Carbon Footprint": "pcf-cx",
}

# E2E go-live target (KPIS box of the original PULSE slides): (main value, second line),
# or None with a note when the source slide does not define it.
GO_LIVE = {
    "Quality": (("MVP Jan '27", "Scaling until 2030"), None),
    "Battery Passport": (("Feb 2027", None), None),
    "Product Passes": (None, "not defined yet"),
    "PURIS": (None, "not specified"),
    "Business Partner Data Mgmt": (("Feb 2027", None), None),
    "Certificate Management": (("May 2028", None), None),
    "Product Carbon Footprint": (("01 Jan 2028", None), None),
}

# Dimension labels split where they would otherwise wrap mid-phrase in a narrow column
DIM_LINES = {"External · CX Association": ["External ·", "CX Association"]}
LINE = 0.165  # height of one 9pt text line

# Row key, label, base height (inches); spare card height is shared out evenly below
ROWS = [
    ("status", "Overall status", 0.38),
    ("phase", "Phase", 0.34),
    ("ms", "Next milestone", 0.82),
    ("dims", "Dimensions", 1.22),
    ("sup", "Suppliers enabled 2026", 0.52),
    ("golive", "Goal: E2E go-live target", 0.52),
    ("dec", "Decision required", 1.14),
]


def pill_width(label):
    # dot + left padding, ~0.068" per 8pt bold character, right padding
    return 0.19 + 0.068 * len(label) + 0.08


def build():
    s = Slide()
    add_header(s, "Management Overview", "Management report on the status of all Catena-X use cases")

    card_x, card_y, card_w, card_bottom = 0.35, 1.07, 12.64, 7.24
    s.shape("Matrix Card", card_x, card_y, card_w, card_bottom - card_y, "roundRect", 2000, "FFFFFF", BORDER)

    panel_x, panel_w = 0.45, 1.45
    label_x, label_w = 0.55, 1.28
    col_x0, col_x1 = 2.00, 12.85
    cw = (col_x1 - col_x0) / len(USE_CASES)
    pad = 0.06
    head_y, head_h = card_y + 0.10, 0.92

    rows_top = head_y + head_h
    rows_bottom = card_bottom - 0.10
    spare = rows_bottom - rows_top - sum(h for _, _, h in ROWS)
    heights = {k: h + spare / len(ROWS) for k, _, h in ROWS}

    # Row label column on a tinted panel, clearly apart from the matrix
    s.shape("Row Label Panel", panel_x, rows_top, panel_w, rows_bottom - rows_top, "roundRect", 3000, TILE)

    # Column headers: name (up to three lines, bottom-aligned), abbreviation, status accent bar
    for c, uc in enumerate(USE_CASES):
        x = col_x0 + c * cw + pad
        w = cw - 2 * pad
        n = uc["name"]
        s.text(f"Col {n} Name", x, head_y, w, 0.58, [[r(n, F_NAME, True, INK)]], anchor="b", line_pts=1300)
        s.text(f"Col {n} Abbr", x, head_y + 0.61, w, 0.16, [[r(ABBR[n], F_ABBR, False, MUTED)]])
        s.shape(f"Col {n} Status Bar", x, head_y + head_h - 0.08, w, 0.045, "roundRect", 50000,
                STATUS[uc["status"]][1])

    # Thin vertical dividers (0.5pt) between the use case columns, header to bottom
    for c in range(1, len(USE_CASES)):
        s.shape(f"Col Divider {c}", col_x0 + c * cw - 0.0035, head_y, 0.007, rows_bottom - head_y, fill=NOT_STARTED)

    y = rows_top
    for key, label, _ in ROWS:
        h = heights[key]
        s.shape(f"Row Divider {label}", col_x0, y, col_x1 - col_x0, 0.007, fill=BORDER)
        s.text(f"Row Label {label}", label_x, y + 0.11, label_w, h - 0.16,
               [[r(label.upper(), F_LABEL, True, INK_STRONG)]], anchor="t", line_pts=1150)
        top = y + 0.11
        for c, uc in enumerate(USE_CASES):
            x = col_x0 + c * cw + pad
            w = cw - 2 * pad
            n = uc["name"]

            if key == "status":
                lbl, fill, fg = STATUS[uc["status"]]
                pw, ph = pill_width(lbl), 0.22
                s.shape(f"{n} Status Pill", x, top - 0.02, pw, ph, "roundRect", 50000, fill)
                s.shape(f"{n} Status Dot", x + 0.09, top - 0.02 + ph / 2 - 0.022, 0.045, 0.045, "ellipse", fill=fg)
                s.text(f"{n} Status Text", x + 0.19, top - 0.02, pw - 0.23, ph,
                       [[r(lbl, F_SUB, True, fg)]], wrap=False)

            elif key == "phase":
                s.text(f"{n} Phase", x, top, w, LINE, [[r(uc["phase"], F_CELL, True, INK_STRONG)]], anchor="t")

            elif key == "ms":
                st, ms_name, ms_date = uc["ms"]
                s.shape(f"{n} Milestone Dot", x, top + 0.02, 0.12, 0.12, "ellipse", fill=DOT[st],
                        line="FFFFFF", line_w=19050)
                s.text(f"{n} Milestone", x + 0.18, top, w - 0.18, h - 0.14,
                       [[r(ms_name, F_CELL, True, INK_STRONG)], [r(ms_date, F_SUB, False, MUTED)]],
                       anchor="t", line_pts=1150)

            elif key == "dims":
                ly = top
                for dim, st in list(zip(DIM_NAMES, uc["dims"])) + [("Supplier Activation", uc["sa"])]:
                    lines = DIM_LINES.get(dim, [dim])
                    s.shape(f"{n} Dim Dot {dim}", x, ly + 0.04, 0.09, 0.09, "ellipse", fill=DOT[st])
                    s.text(f"{n} Dim {dim}", x + 0.15, ly, w - 0.15, LINE * len(lines),
                           [[r(t, F_CELL, False, INK)] for t in lines], anchor="t", line_pts=1150)
                    ly += LINE * len(lines) + 0.02

            elif key == "sup":
                done, total = uc["sup"]
                pct = round(100 * done / total)
                paras = [[r(f"{done} / {total}", F_CELL, True, COMPLETED), r(f"  {pct}%", F_SUB, False, MUTED)]]
                if uc.get("note"):
                    paras.append([r("rescoped from 15", F_SUB, False, MUTED)])
                s.text(f"{n} Suppliers", x, top, w, h - 0.14, paras, anchor="t", line_pts=1150)

            elif key == "golive":
                value, note = GO_LIVE[n]
                if value:
                    paras = [[r(value[0], F_CELL, True, INK_STRONG)]]
                    if value[1]:
                        paras.append([r(value[1], F_SUB, False, MUTED)])
                else:
                    paras = [[r("–  ", F_CELL, False, FAINT), r(note, F_SUB, False, FAINT, None, True)]]
                s.text(f"{n} Go-live", x, top, w, h - 0.14, paras, anchor="t", line_pts=1150)

            elif key == "dec":
                if uc.get("decision"):
                    num = uc["decision"]
                    s.shape(f"{n} Decision Flag", x, top, 0.19, 0.19, "ellipse", fill=COMPLETED)
                    s.text(f"{n} Decision No", x, top, 0.19, 0.19,
                           [[r(str(num), F_SUB, True, "FFFFFF")]], algn="ctr")
                    s.text(f"{n} Decision", x + 0.25, top, w - 0.25, h - 0.14,
                           [[r(DECISIONS[num - 1][1], F_CELL, False, INK)]], anchor="t", line_pts=1150)
                else:
                    s.text(f"{n} Decision None", x, top, w, LINE, [[r("–", F_CELL, False, FAINT)]], anchor="t")
        y += h
    return s


def main():
    write_pptx(build().shapes, OUTPUT,
               "Use Case Pulse Management Overview CW 40 / 2026 – one column per use case (read top to bottom), "
               "one KPI per row, sorted by overall status.")


if __name__ == "__main__":
    main()
