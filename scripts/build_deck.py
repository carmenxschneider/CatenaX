"""Build one deck with both overview variants.

    slide 1: Use Case Pulse – Overview (one row per use case, build_overview.py)
    slide 2: Use Case Pulse – Management Overview (one card per use case, build_management_overview.py)

    python scripts/build_deck.py [output.pptx]
"""

import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import build_management_overview  # noqa: E402
import build_overview  # noqa: E402
from xml.sax.saxutils import escape  # noqa: E402

TEMPLATE = "templates/Template_UseCase_PULSE.pptx"
OUTPUT = sys.argv[1] if len(sys.argv) > 1 else "output/UseCase_Pulse_Overview_Varianten_CW40_26.pptx"

SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def main():
    zin = zipfile.ZipFile(TEMPLATE)
    read = lambda name: zin.read(name).decode("utf8")  # noqa: E731

    slide = read("ppt/slides/slide1.xml")
    graphic_start = slide.index("<p:graphicFrame>")
    head_end = slide.index("</p:graphicFrame>") + len("</p:graphicFrame>")
    tail_start = slide.rindex("</p:spTree>")

    # Slide 1 keeps the template's hidden think-cell frame; slide 2 gets no OLE parts of its own
    slide1 = slide[:head_end] + "".join(build_overview.build().shapes) + slide[tail_start:]
    slide2 = (slide[:graphic_start] + "".join(build_management_overview.build().shapes) + slide[tail_start:])

    layout = re.search(r'Target="(\.\./slideLayouts/[^"]+)"', read("ppt/slides/_rels/slide1.xml.rels")).group(1)
    slide2_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'<Relationship Id="rId1" Type="{LAYOUT_REL}" Target="{layout}"/></Relationships>')

    pres_rels = read("ppt/_rels/presentation.xml.rels")
    new_rid = "rId%d" % (max(int(i) for i in re.findall(r'Id="rId(\d+)"', pres_rels)) + 1)
    pres_rels = pres_rels.replace(
        "</Relationships>", f'<Relationship Id="{new_rid}" Type="{SLIDE_REL}" Target="slides/slide2.xml"/></Relationships>')

    pres = read("ppt/presentation.xml")
    new_id = max(int(i) for i in re.findall(r'<p:sldId id="(\d+)"', pres)) + 1
    pres = pres.replace("</p:sldIdLst>", f'<p:sldId id="{new_id}" r:id="{new_rid}"/></p:sldIdLst>')

    ct = read("[Content_Types].xml").replace(
        "</Types>", f'<Override PartName="/ppt/slides/slide2.xml" ContentType="{SLIDE_CT}"/></Types>')
    app = re.sub(r"<Slides>\d+</Slides>", "<Slides>2</Slides>", read("docProps/app.xml"))

    notes = re.sub(r"<a:t>Vorlage für Confluence[^<]*</a:t>",
                   "<a:t>" + escape("Use Case Pulse CW 40 / 2026 – slide 1: one row per use case; "
                                    "slide 2: Management Overview, one card per use case (read top to bottom).")
                   + "</a:t>", read("ppt/notesSlides/notesSlide1.xml"))

    replace = {
        "ppt/slides/slide1.xml": slide1,
        "ppt/_rels/presentation.xml.rels": pres_rels,
        "ppt/presentation.xml": pres,
        "[Content_Types].xml": ct,
        "docProps/app.xml": app,
        "ppt/notesSlides/notesSlide1.xml": notes,
    }
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replace.get(item.filename)
            zout.writestr(item, data.encode("utf8") if data is not None else zin.read(item.filename))
        zout.writestr("ppt/slides/slide2.xml", slide2.encode("utf8"))
        zout.writestr("ppt/slides/_rels/slide2.xml.rels", slide2_rels.encode("utf8"))
    print("wrote", OUTPUT)


if __name__ == "__main__":
    main()
