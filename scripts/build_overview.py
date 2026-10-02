"""Build the "Use Case Pulse – Overview" slide.

Reuses the master, theme, background and notes of the Use Case PULSE template
and replaces the slide content with one row per use case. Every colour, font
size and corner radius is taken from the template's slide1.xml.

    python scripts/build_overview.py [template.pptx] [output.pptx]
"""

import re
import sys
import zipfile
from xml.sax.saxutils import escape

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "templates/Template_UseCase_PULSE.pptx"
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/UseCase_Pulse_Overview_CW40_26.pptx"

EMU = 914400
FONT = "Segoe UI"

# Status colours (template legend)
ON_TRACK, PROBLEM, ESCALATION, NOT_STARTED, COMPLETED = "53BE70", "F0B135", "E64343", "CFD6DB", "001E50"
# Neutrals (template)
INK, INK_STRONG, INK_TITLE = "292D31", "151B21", "10151A"
MUTED, FAINT = "63707A", "9AA6AD"
BORDER, TILE, TRACK, BAR = "E2E8EC", "F6F8FA", "E5E9EC", "2F7AB5"

STATUS = {
    "esc": ("Escalation needed", ESCALATION, "FFFFFF"),
    "ps": ("Problem Solving", PROBLEM, "FFFFFF"),
    "ok": ("On Track", ON_TRACK, "FFFFFF"),
    "ns": ("Not Started", NOT_STARTED, INK),
}
DOT = {"ok": ON_TRACK, "ps": PROBLEM, "esc": ESCALATION, "ns": NOT_STARTED, "done": COMPLETED}

# Sorted by overall status: Escalation needed -> Problem solving -> On track
USE_CASES = [
    dict(name="Quality", abbr="quality-cx", status="esc", phase="Piloting",
         ms=("ps", "Domain expertise aligned", "30 Sep 2026"),
         dims=["ok", "ps", "ok", "ps"], sa="ps", sup=(0, 1), note="rescoped", decision=1),
    dict(name="Battery Passport", abbr="Batt Pass", status="ps", phase="Pilot",
         ms=("ok", "Tech requirements done", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok"], sa="ok", sup=(0, 1)),
    dict(name="Product Passes", abbr="pass-cx", status="ps", phase="Definition",
         ms=("ps", "Klärung & Prio Produktpässe", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok"], sa="ok", sup=(0, 1)),
    dict(name="PURIS", abbr="PURIS", status="ps", phase="Piloting",
         ms=("ps", "Plan for existing challenges", "9 Oct 2026"),
         dims=["ok", "ps", "ps", "ps"], sa="ps", sup=(1, 6), decision=2),
    dict(name="Business Partner Data Mgmt", abbr="BPDM", status="ps", phase="Implementation",
         ms=("ok", "Define BPDM target picture", "31 Oct 2026"),
         dims=["ps", "ps", "ok", "ok"], sa="ns", sup=(0, 1)),
    dict(name="Certificate Management", abbr="cert-mgmt", status="ps", phase="Scaling",
         ms=("ok", "Official CX CCM release", "Sep 2026"),
         dims=["ps", "ok", "ok", "ok"], sa="ok", sup=(72, 100)),
    dict(name="Product Carbon Footprint", abbr="pcf-cx", status="ok", phase="Piloting",
         ms=("ok", "Supplier activation pilot 1", "15 Oct 2026"),
         dims=["ok", "ok", "ok", "ok"], sa="ok", sup=(0, 1)),
]
DIM_NAMES = ["External · CX Association", "Business", "Data Provisioning", "Technical"]

DECISIONS = [
    ("Quality:", "Rescope von 15 auf 1 Lieferant, bis Early Warning Production MVP live ist."),
    ("PURIS:", "Wie und wann wird über die WINGS-Connectivity-Rollout-Roadmap entschieden?"),
]


class Slide:
    def __init__(self):
        self.shapes = []
        self.next_id = 100

    def _id(self):
        self.next_id += 1
        return self.next_id

    def shape(self, name, x, y, w, h, geom="rect", adj=None, fill=None, line=None, line_w=6350):
        av = f'<a:gd name="adj" fmla="val {adj}"/>' if adj is not None else ""
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
        ln = (f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
              if line else "<a:ln><a:noFill/></a:ln>")
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self._id()}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{_xfrm(x, y, w, h)}<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>'
            f'{fill_xml}{ln}<a:effectLst/></p:spPr>'
            f'<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>')

    def text(self, name, x, y, w, h, runs, anchor="ctr", algn="l", wrap=True, line_pts=None):
        """runs: list of paragraphs, each a list of (text, size_pt, bold, color, spacing)."""
        paras = []
        for para in runs:
            lnspc = f'<a:lnSpc><a:spcPts val="{line_pts}"/></a:lnSpc>' if line_pts else ""
            rs = "".join(
                f'<a:r><a:rPr lang="de-DE" sz="{int(sz * 100)}" b="{1 if b else 0}" kern="0"'
                f'{f" spc={chr(34)}{spc}{chr(34)}" if spc else ""} dirty="0">'
                f'<a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
                f'<a:latin typeface="{FONT}" pitchFamily="34" charset="0"/>'
                f'<a:ea typeface="{FONT}" pitchFamily="34" charset="-122"/>'
                f'<a:cs typeface="{FONT}" pitchFamily="34" charset="-120"/></a:rPr>'
                f'<a:t>{escape(t)}</a:t></a:r>'
                for t, sz, b, c, spc in para)
            paras.append(f'<a:p><a:pPr marL="0" indent="0" algn="{algn}">{lnspc}<a:buNone/></a:pPr>{rs}</a:p>')
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self._id()}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{_xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="{"square" if wrap else "none"}" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}">'
            f'<a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')


def _xfrm(x, y, w, h):
    return (f'<a:xfrm><a:off x="{round(x * EMU)}" y="{round(y * EMU)}"/>'
            f'<a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>')


def r(text, sz, b=False, c=INK, spc=None):
    return (text, sz, b, c, spc)


def build():
    s = Slide()

    # ---------- Header (same positions/sizes as the PULSE template) ----------
    s.text("Eyebrow", 0.35, 0.26, 4.0, 0.12, [[r("USE CASE PULSE · CATENA-X", 7.5, True, COMPLETED, 75)]])
    s.text("Headline", 0.35, 0.42, 7.99, 0.38, [[r("Overview", 26, True, INK_TITLE, -26)]])
    s.text("Subline", 0.35, 0.84, 6.0, 0.16,
           [[r(f"{len(USE_CASES)} use cases · sorted by overall status", 9.5, False, MUTED)]])

    # Legend box + date box (top right, like Responsibilities/Date boxes)
    s.shape("Legend Box", 8.18, 0.26, 3.53, 0.73, "roundRect", 11363, "FFFFFF", BORDER, 9525)
    s.text("Legend Title", 8.30, 0.31, 3.3, 0.10, [[r("Status legend", 6.5, True, MUTED)]], anchor="t")
    legend = [("On Track", ON_TRACK), ("Problem Solving", PROBLEM), ("Escalation needed", ESCALATION),
              ("Not Started", NOT_STARTED), ("Completed", COMPLETED)]
    for i, (lbl, col) in enumerate(legend):
        lx = 8.30 + (i % 3) * 1.13
        ly = 0.55 + (i // 3) * 0.20
        s.shape(f"Legend Dot {lbl}", lx, ly + 0.015, 0.09, 0.09, "ellipse", fill=col)
        s.text(f"Legend {lbl}", lx + 0.14, ly, 0.98, 0.12, [[r(lbl, 6.5, False, "49545D")]])

    s.shape("Date Box", 11.79, 0.26, 1.19, 0.73, "roundRect", 11363, "FFFFFF", BORDER, 9525)
    s.text("Date Label", 11.89, 0.46, 1.00, 0.11, [[r("Date", 6.5, True, MUTED)]], anchor="t")
    s.text("Date Value", 11.89, 0.60, 1.00, 0.20, [[r("CW 40 / 2026", 8.5, True, INK_STRONG)]], anchor="t")

    # ---------- Main card ----------
    card_y, row_h, head_h = 1.07, 0.725, 0.34
    card_h = head_h + row_h * len(USE_CASES) + 0.10
    s.shape("Overview Card", 0.35, card_y, 12.64, card_h, "roundRect", 2000, "FFFFFF", BORDER)

    cols = dict(uc=0.55, st=2.70, ph=4.30, ms=5.48, dim=7.72, sup=11.15, dec=12.50)
    head = [("uc", "Use case"), ("st", "Overall status"), ("ph", "Phase"),
            ("ms", "Next milestone"), ("dim", "Dimensions"), ("sup", "Suppliers")]
    for key, lbl in head:
        s.text(f"Head {lbl}", cols[key], card_y + 0.12, 2.0, 0.14,
               [[r(lbl.upper(), 7.5, True, INK_STRONG, 51)]])

    y0 = card_y + head_h
    for i, uc in enumerate(USE_CASES):
        y = y0 + i * row_h
        s.shape(f"Row Divider {i + 1}", 0.55, y, 12.24, 0.007, fill=BORDER)
        cy = y + row_h / 2
        n = uc["name"]

        # Use case name + abbreviation
        s.text(f"{n} Name", cols["uc"], cy - 0.19, 2.15, 0.20, [[r(n, 8.75, True, INK)]])
        s.text(f"{n} Abbr", cols["uc"], cy + 0.02, 2.15, 0.15, [[r(uc["abbr"], 7.5, False, MUTED)]])

        # Overall status pill (wide pill: 0.28" high, white dot + 7.5pt bold text)
        lbl, fill, fg = STATUS[uc["status"]]
        # Template pill widths plus ~8% slack
        pw = {"On Track": 1.03, "Problem Solving": 1.33, "Escalation needed": 1.46, "Not Started": 1.07}[lbl]
        s.shape(f"{n} Status Pill", cols["st"], cy - 0.14, pw, 0.28, "roundRect", 50000, fill)
        s.shape(f"{n} Status Dot", cols["st"] + 0.15, cy - 0.025, 0.05, 0.05, "ellipse", fill=fg)
        s.text(f"{n} Status Text", cols["st"] + 0.26, cy - 0.07, pw - 0.30, 0.14,
               [[r(lbl, 7.5, True, fg)]], wrap=False)

        # Phase
        s.text(f"{n} Phase", cols["ph"], cy - 0.08, 1.15, 0.16, [[r(uc["phase"], 8.5, True, INK_STRONG)]])

        # Next milestone: timeline dot (19 px, white ring) + title + date
        ms_st, ms_name, ms_date = uc["ms"]
        s.shape(f"{n} Milestone Dot", cols["ms"], cy - 0.155, 0.13, 0.13, "ellipse", fill=DOT[ms_st],
                line="FFFFFF", line_w=19050)
        s.text(f"{n} Milestone", cols["ms"] + 0.20, cy - 0.19, 2.04, 0.20, [[r(ms_name, 8.25, True, INK_STRONG)]])
        s.text(f"{n} Milestone Date", cols["ms"] + 0.20, cy + 0.02, 2.04, 0.15, [[r(ms_date, 7.5, False, MUTED)]])

        # Dimensions: 2 x 2 grid + Supplier Activation as separate tile below
        dy = y + 0.075
        for k, st in enumerate(uc["dims"]):
            dx = cols["dim"] + (k % 2) * 1.70
            ly = dy + (k // 2) * 0.185
            s.shape(f"{n} Dim Dot {DIM_NAMES[k]}", dx + 0.06, ly + 0.035, 0.09, 0.09, "ellipse", fill=DOT[st])
            s.text(f"{n} Dim {DIM_NAMES[k]}", dx + 0.20, ly, 1.56, 0.16, [[r(DIM_NAMES[k], 7.5, False, INK)]])
        ty = dy + 2 * 0.185 + 0.01
        s.shape(f"{n} Supplier Activation Tile", cols["dim"], ty, 3.25, 0.19, "roundRect", 30000, TILE, BORDER)
        s.shape(f"{n} Dim Dot Supplier Activation", cols["dim"] + 0.06, ty + 0.05, 0.09, 0.09, "ellipse",
                fill=DOT[uc["sa"]])
        sa_runs = [r("Supplier Activation", 7.5, True, INK_STRONG)]
        if uc["sa"] == "ns":
            sa_runs.append(r("  · not started", 7.5, False, FAINT))
        s.text(f"{n} Dim Supplier Activation", cols["dim"] + 0.20, ty + 0.015, 3.0, 0.16, [sa_runs])

        # Suppliers: "x / y" + percentage + enablement bar
        done, total = uc["sup"]
        pct = round(100 * done / total)
        s.text(f"{n} Suppliers", cols["sup"], cy - 0.21, 1.25, 0.20,
               [[r(f"{done} / {total}", 11, True, COMPLETED),
                 r(f"  {pct}%" + (" · rescoped" if uc.get("note") else ""), 7.5, False, MUTED)]])
        bw = 1.18
        s.shape(f"{n} Supplier Track", cols["sup"], cy + 0.07, bw, 0.07, "roundRect", 50000, TRACK)
        if pct:
            s.shape(f"{n} Supplier Bar", cols["sup"], cy + 0.07, bw * pct / 100, 0.07, "roundRect", 50000, BAR)
        s.shape(f"{n} Supplier Knob", cols["sup"] + max(0, bw * pct / 100 - 0.075), cy + 0.03, 0.15, 0.15, "ellipse",
                fill="FFFFFF", line=ON_TRACK, line_w=25400)

        # Decision flag (numbered, links to the decision list below)
        if uc.get("decision"):
            s.shape(f"{n} Decision Flag", cols["dec"], cy - 0.11, 0.22, 0.22, "ellipse", fill=COMPLETED)
            s.text(f"{n} Decision No", cols["dec"], cy - 0.11, 0.22, 0.22,
                   [[r(str(uc["decision"]), 7.5, True, "FFFFFF")]], algn="ctr")
        else:
            s.text(f"{n} Decision None", cols["dec"], cy - 0.08, 0.22, 0.16, [[r("–", 8.5, False, FAINT)]], algn="ctr")

    # ---------- Decisions card ----------
    dec_y = card_y + card_h + 0.08
    dec_h = 7.24 - dec_y
    s.shape("Decisions Card", 0.35, dec_y, 12.64, dec_h, "roundRect", 20000, "FFFFFF", BORDER)
    s.text("Decisions Title", 0.55, dec_y + 0.09, 3.0, 0.14, [[r("Decisions required:", 8.5, True, INK_STRONG, 20)]])
    for j, (who, what) in enumerate(DECISIONS):
        dx = 0.55 + j * 6.17
        ry = dec_y + dec_h - 0.27
        s.shape(f"Decision {j + 1} Flag", dx, ry, 0.20, 0.20, "ellipse", fill=COMPLETED)
        s.text(f"Decision {j + 1} No", dx, ry, 0.20, 0.20, [[r(str(j + 1), 7, True, "FFFFFF")]], algn="ctr")
        s.text(f"Decision {j + 1}", dx + 0.28, ry + 0.01, 5.80, 0.18,
               [[r(who + " ", 8.5, True, INK), r(what, 8.5, False, INK)]])
    return s


def main():
    s = build()
    zin = zipfile.ZipFile(TEMPLATE)
    slide = zin.read("ppt/slides/slide1.xml").decode("utf8")
    # Keep group header + hidden think-cell frame, replace every visible shape
    head_end = slide.index("</p:graphicFrame>") + len("</p:graphicFrame>")
    tail_start = slide.rindex("</p:spTree>")
    slide = slide[:head_end] + "".join(s.shapes) + slide[tail_start:]

    notes = zin.read("ppt/notesSlides/notesSlide1.xml").decode("utf8")
    notes = re.sub(r"<a:t>Vorlage für Confluence[^<]*</a:t>",
                   "<a:t>Use Case Pulse Overview CW 40 / 2026 – one row per use case, "
                   "sorted by overall status.</a:t>", notes)

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                data = slide.encode("utf8")
            elif item.filename == "ppt/notesSlides/notesSlide1.xml":
                data = notes.encode("utf8")
            zout.writestr(item, data)
    print("wrote", OUTPUT)


if __name__ == "__main__":
    main()
