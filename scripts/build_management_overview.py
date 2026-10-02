"""Build the "Use Case Pulse – Management Overview" slide (variant 2).

Transposed version of build_overview.py: one column per use case, one row per
KPI, so management can scan a single KPI across all use cases. Uses the same
template, colours and shape vocabulary as variant 1.

    python scripts/build_management_overview.py [template.pptx] [output.pptx]
"""

import sys

from build_overview import (
    BORDER, COMPLETED, DECISIONS, DIM_NAMES, DOT, FAINT, INK, INK_STRONG, MUTED, STATUS, TEMPLATE, USE_CASES,
    Slide, add_header, r, write_pptx,
)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Management_Overview_CW40_26.pptx"

# Management KPIs per use case (KPIS / FINAL GOAL boxes of the original PULSE slides).
# go_live: (main value, second line) — None = not defined in the source slide.
EXTRA = {
    "Quality": dict(go_live=("MVP Jan '27", "Scaling until 2030"), benefit=None),
    "Battery Passport": dict(go_live=("Feb 2027", None), benefit=None),
    "Product Passes": dict(go_live=None, go_live_note="not defined yet", benefit=None),
    "PURIS": dict(go_live=None, go_live_note="not specified", benefit="€32.4M"),
    "Business Partner Data Mgmt": dict(go_live=("Feb 2027", None), benefit=None),
    "Certificate Management": dict(go_live=("May 2028", None), benefit=None),
    "Product Carbon Footprint": dict(go_live=("01 Jan 2028", None), benefit=None),
}

PILL_W = {"On Track": 1.03, "Problem Solving": 1.33, "Escalation needed": 1.46, "Not Started": 1.07}

# Row key, label, base height (inches); spare card height is shared out evenly below
ROWS = [
    ("status", "Overall status", 0.42),
    ("phase", "Phase", 0.36),
    ("ms", "Next milestone", 0.62),
    ("dims", "Dimensions at risk", 0.82),
    ("sup", "Suppliers enabled 2026", 0.52),
    ("golive", "Goal: E2E go-live target", 0.50),
    ("benefit", "Goal: Final benefit", 0.42),
    ("dec", "Decision required", 1.00),
]
GROW = {k: 1 for k, _, _ in ROWS}


def dims_at_risk(uc):
    """Only Problem Solving / Escalation dimensions; green and not-started ones are left out."""
    named = list(zip(DIM_NAMES, uc["dims"])) + [("Supplier Activation", uc["sa"])]
    return [(name, st) for name, st in named if st in ("ps", "esc")]


def build():
    s = Slide()
    add_header(s, "Management Overview",
               f"{len(USE_CASES)} use cases · one KPI per row · sorted by overall status")

    card_x, card_y, card_w, card_bottom = 0.35, 1.07, 12.64, 7.24
    s.shape("Matrix Card", card_x, card_y, card_w, card_bottom - card_y, "roundRect", 2000, "FFFFFF", BORDER)

    label_x, label_w = 0.55, 1.15
    col_x0, col_x1 = 1.82, 12.79
    cw = (col_x1 - col_x0) / len(USE_CASES)
    pad = 0.05
    head_y, head_h = card_y + 0.10, 0.60

    # Distribute the remaining card height over the rows that wrap
    rows_top = head_y + head_h
    spare = (card_bottom - 0.08) - rows_top - sum(h for _, _, h in ROWS)
    weight = sum(GROW.values())
    heights = {k: h + spare * GROW.get(k, 0) / weight for k, _, h in ROWS}

    # Column headers: use case name (up to two lines, bottom-aligned) + abbreviation
    for c, uc in enumerate(USE_CASES):
        x = col_x0 + c * cw + pad
        s.text(f"Col {uc['name']} Name", x, head_y, cw - 2 * pad, 0.34,
               [[r(uc["name"], 8.75, True, INK)]], anchor="b", line_pts=1100)
        s.text(f"Col {uc['name']} Abbr", x, head_y + 0.37, cw - 2 * pad, 0.15,
               [[r(uc["abbr"], 7.5, False, MUTED)]])
    # Vertical column dividers
    for c in range(1, len(USE_CASES)):
        s.shape(f"Col Divider {c}", col_x0 + c * cw, head_y + 0.04, 0.007,
                card_bottom - 0.12 - head_y - 0.04, fill=BORDER)

    y = rows_top
    for key, label, _ in ROWS:
        h = heights[key]
        s.shape(f"Row Divider {label}", label_x, y, col_x1 - label_x, 0.007, fill=BORDER)
        s.text(f"Row Label {label}", label_x, y + 0.12, label_w, 0.30,
               [[r(label.upper(), 7, True, INK_STRONG, 40)]], anchor="t", line_pts=950)
        top = y + 0.12
        for c, uc in enumerate(USE_CASES):
            x = col_x0 + c * cw + pad
            w = cw - 2 * pad
            n = uc["name"]
            ex = EXTRA[n]

            if key == "status":
                lbl, fill, fg = STATUS[uc["status"]]
                pw = PILL_W[lbl]
                py = top - 0.05
                s.shape(f"{n} Status Pill", x, py, pw, 0.28, "roundRect", 50000, fill)
                s.shape(f"{n} Status Dot", x + 0.15, py + 0.115, 0.05, 0.05, "ellipse", fill=fg)
                s.text(f"{n} Status Text", x + 0.26, py + 0.07, pw - 0.30, 0.14,
                       [[r(lbl, 7.5, True, fg)]], wrap=False)

            elif key == "phase":
                s.text(f"{n} Phase", x, top, w, 0.17, [[r(uc["phase"], 8.5, True, INK_STRONG)]], anchor="t")

            elif key == "ms":
                st, ms_name, ms_date = uc["ms"]
                s.shape(f"{n} Milestone Dot", x, top + 0.005, 0.13, 0.13, "ellipse", fill=DOT[st],
                        line="FFFFFF", line_w=19050)
                s.text(f"{n} Milestone", x + 0.19, top - 0.01, w - 0.19, h - 0.14,
                       [[r(ms_name, 8, True, INK_STRONG)], [r(ms_date, 7.5, False, MUTED)]],
                       anchor="t", line_pts=1100)

            elif key == "dims":
                risk = dims_at_risk(uc)
                if not risk:
                    s.text(f"{n} No Dims", x, top, w, 0.16,
                           [[r("No dimensions at risk", 7.5, False, FAINT, None, True)]], anchor="t")
                for k, (dim, st) in enumerate(risk):
                    ly = top + k * 0.17
                    s.shape(f"{n} Risk Dot {dim}", x, ly + 0.03, 0.09, 0.09, "ellipse", fill=DOT[st])
                    s.text(f"{n} Risk {dim}", x + 0.14, ly, w - 0.14, 0.16,
                           [[r(dim, 7.5, False, INK)]], anchor="t")

            elif key == "sup":
                done, total = uc["sup"]
                pct = round(100 * done / total)
                paras = [[r(f"{done} / {total}", 11, True, COMPLETED), r(f"  {pct}%", 7.5, False, MUTED)]]
                if uc.get("note"):
                    paras.append([r("rescoped from 15", 7, False, MUTED)])
                s.text(f"{n} Suppliers", x, top - 0.04, w, h - 0.10, paras, anchor="t", line_pts=1300)

            elif key == "golive":
                gl = ex["go_live"]
                if gl:
                    paras = [[r(gl[0], 8.5, True, INK_STRONG)]]
                    if gl[1]:
                        paras.append([r(gl[1], 7.5, False, MUTED)])
                else:
                    paras = [[r("–  ", 8.5, False, FAINT), r(ex["go_live_note"], 7.5, False, FAINT, None, True)]]
                s.text(f"{n} Go-live", x, top, w, h - 0.14, paras, anchor="t", line_pts=1100)

            elif key == "benefit":
                if ex["benefit"]:
                    s.text(f"{n} Benefit", x, top - 0.04, w, 0.22, [[r(ex["benefit"], 11, True, COMPLETED)]],
                           anchor="t")
                else:
                    s.text(f"{n} Benefit", x, top, w, 0.16,
                           [[r("Not yet quantified", 7.5, False, FAINT, None, True)]], anchor="t")

            elif key == "dec":
                if uc.get("decision"):
                    num = uc["decision"]
                    s.shape(f"{n} Decision Flag", x, top - 0.01, 0.20, 0.20, "ellipse", fill=COMPLETED)
                    s.text(f"{n} Decision No", x, top - 0.01, 0.20, 0.20,
                           [[r(str(num), 7, True, "FFFFFF")]], algn="ctr")
                    s.text(f"{n} Decision", x + 0.26, top, w - 0.26, h - 0.16,
                           [[r(DECISIONS[num - 1][1], 7.5, False, INK)]], anchor="t", line_pts=1050)
                else:
                    s.text(f"{n} Decision None", x, top, w, 0.16, [[r("–", 8.5, False, FAINT)]], anchor="t")
        y += h
    return s


def main():
    write_pptx(build().shapes, OUTPUT,
               "Use Case Pulse Management Overview CW 40 / 2026 – one column per use case, one KPI per row, "
               "sorted by overall status. Dimensions at risk show only Problem Solving / Escalation.")


if __name__ == "__main__":
    main()
